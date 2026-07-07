"""
Google Cloud Function 진입점

트리거: Pub/Sub (입시_트래킹 변경 감지 후 메시지 발행)
동작: generate_final_sheets.py의 로직 실행

배포 방법:
  gcloud functions deploy update_final_sheets \
    --runtime python312 \
    --trigger-topic high-school-tracking-update \
    --entry-point main \
    --service-account-email=school-bot@gen-lang-client-0367740438.iam.gserviceaccount.com

환경 변수 설정:
  SPREADSHEET_ID: 14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0
"""

import os
import json
import base64
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0")

# 전기고/후기고 유형 및 섹션 정의
EARLY_TYPES = {"영재고", "과학고", "예술계고", "특성화고"}

EARLY_SECTIONS = [
    ("과학고", ["연번", "반", "이름", "성별", "지원교", "1차", "2차", "최종"]),
    ("예고", ["연번", "반", "성명", "성별", "예술계고", "1차", "2차", "최종"]),
    ("특성화고 / 마이스터고", ["연번", "반", "이름", "성별", "지원고등학교", "배정학과", "1차", "2차", "최종"]),
]

LATE_SECTIONS = [
    ("자사고", ["연번", "반", "이름", "성별", "지원교", "1차", "2차", "최종"]),
    ("외고/국제고", ["연번", "반", "성명", "성별", "지원교", "1차", "2차", "최종"]),
    ("비평준화고 / 중점고", ["연번", "반", "이름", "성별", "지원고등학교", "1차", "2차", "최종"]),
]


def update_final_sheets(event, context):
    """
    Pub/Sub 메시지 처리 및 최종 합불 시트 업데이트

    event: Pub/Sub 메시지
    context: Cloud Functions 컨텍스트
    """
    print(f"Received event: {json.dumps(event)}")

    try:
        # 서비스 계정 인증 (Cloud Function 내장 인증 사용)
        creds = Credentials.from_service_account_info(
            json.loads(os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON", "{}")),
            scopes=SCOPES,
        )
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(SPREADSHEET_ID)

        # 입시_트래킹 시트 읽기
        tracking_sht = ss.worksheet("입시_트래킹")
        rows = tracking_sht.get_all_values()

        header = rows[0]
        col_map = {h: i for i, h in enumerate(header)}

        # 합격자 분류
        passed = defaultdict(lambda: defaultdict(list))

        for r in rows[1:]:
            if len(r) < 3 or not r[col_map.get("성명", -1) if col_map.get("성명", -1) < len(r) else ""]:
                continue

            final = r[col_map.get("최종", -1)].strip().lower() if col_map.get("최종", -1) < len(r) else ""
            if final not in ["합격", "pass", "o", "yes", "v"]:
                continue

            student = {
                "반": r[col_map.get("반", -1)] if col_map.get("반", -1) < len(r) else "",
                "번호": r[col_map.get("번호", -1)] if col_map.get("번호", -1) < len(r) else "",
                "성명": r[col_map.get("성명", -1)] if col_map.get("성명", -1) < len(r) else "",
                "성별": r[col_map.get("성별", -1)] if col_map.get("성별", -1) < len(r) else "",
                "유형": r[col_map.get("유형", -1)] if col_map.get("유형", -1) < len(r) else "",
                "지원학교": r[col_map.get("지원학교", -1)] if col_map.get("지원학교", -1) < len(r) else "",
                "학과": r[col_map.get("학과", -1)] if col_map.get("학과", -1) < len(r) else "",
                "1차": r[col_map.get("1차", -1)] if col_map.get("1차", -1) < len(r) else "",
                "2차": r[col_map.get("2차", -1)] if col_map.get("2차", -1) < len(r) else "",
            }

            school_type = student["유형"]
            if school_type in EARLY_TYPES:
                passed["early"][school_type].append(student)
            else:
                passed["late"][school_type].append(student)

        # 전기고 최종 합불 시트 업데이트
        early_sht = ss.worksheet("전기고 최종 합불")
        early_sht.clear()
        early_data = []

        for section_name, section_cols in EARLY_SECTIONS:
            early_data.append(["" if i > 0 else section_name] + [""] * (len(section_cols) - 1))
            early_data.append([""] * len(section_cols))
            early_data.append(section_cols)

            section_type = section_name.split(" / ")[0]
            students = passed["early"].get(section_type, [])

            for seq, student in enumerate(students, 1):
                row_data = [
                    str(seq),
                    student["반"],
                    student["성명"],
                    student["성별"],
                    student["지원학교"],
                    student["1차"],
                    student["2차"],
                    "합격",
                ]
                if section_type == "특성화고":
                    row_data.insert(5, student.get("학과", ""))
                early_data.append(row_data)

        early_sht.update(values=early_data, range_name="A1")

        # 후기고 최종 합불 시트 업데이트
        late_sht = ss.worksheet("후기고 최종 합불")
        late_sht.clear()
        late_data = []

        for section_name, section_cols in LATE_SECTIONS:
            late_data.append(["" if i > 0 else section_name] + [""] * (len(section_cols) - 1))
            late_data.append([""] * len(section_cols))
            late_data.append(section_cols)

            section_type = section_name.split(" / ")[0]
            students = passed["late"].get(section_type, [])

            for seq, student in enumerate(students, 1):
                row_data = [
                    str(seq),
                    student["반"],
                    student["성명"],
                    student["성별"],
                    student["지원학교"],
                    student["1차"],
                    student["2차"],
                    "합격",
                ]
                late_data.append(row_data)

        late_sht.update(values=late_data, range_name="A1")

        print("✅ 최종 합불 시트 업데이트 완료")
        return {"status": "success"}

    except Exception as e:
        print(f"❌ 오류: {e}")
        return {"status": "error", "message": str(e)}, 500
