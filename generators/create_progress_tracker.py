#!/usr/bin/env python3
"""
입시 진행 현황 시트 생성 (방안 A)

반별 시트의 모든 지원 유형을 감지하여
학생별 지원 유형별로 한 행씩 생성

실행:
  python create_progress_tracker.py 2026
  또는
  python create_progress_tracker.py [SPREADSHEET_ID]
"""

import sys
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_IDS = {
    "2025": "1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI",
    "2026": "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0",
}

# 지원 유형 매핑 (컬럼 인덱스)
TYPE_MAP = {
    7: "영재고",
    8: "과학고",
    9: "예술계고",
    10: "특성화고",
    12: "자사고",
    13: "외고/국제고",
    14: "일반고",
    15: "기타/대안",
}

# 우선순위 (전기고가 우선)
PRIORITY = ["영재고", "과학고", "예술계고", "특성화고", "자사고", "외고/국제고", "기타/대안", "일반고"]


def detect_types(row):
    """행 데이터에서 모든 지원 유형 감지 (복수 지원 가능)"""
    row = row + [""] * (30 - len(row))
    types = []

    for col_idx, type_name in TYPE_MAP.items():
        val = row[col_idx].strip()
        if val and val not in ["", "X", "x", "0"]:
            types.append(type_name)
        elif val in ["○", "O", "o"]:
            types.append(type_name)

    return types


def get_final_type(types_list):
    """복수 지원한 유형 중 최종 유형 판정 (우선순위 기반)"""
    if not types_list:
        return ""
    for ptype in PRIORITY:
        if ptype in types_list:
            return ptype
    return types_list[0]


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    ss_id = SPREADSHEET_IDS.get(year)

    if not ss_id:
        print(f"❌ 잘못된 연도: {year}")
        return

    print(f"{'='*70}")
    print(f" {year}학년도 입시 진행 현황 시트 생성")
    print(f"{'='*70}")

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    # ── Step 1. 반별 데이터 수집 및 모든 지원 유형 감지 ─────────
    print(f"\n[1/3] 반별 데이터에서 지원 유형 감지 중...")

    # (반, 번호) → {name, gender, types}
    student_data = {}
    type_stats = defaultdict(int)

    target_sheet_pattern = "진학희망 및 지원유형 조사" if year == "2025" else "Sheet1"

    for sht in ss.worksheets():
        # 반별 시트 찾기
        if "진학희망 및 지원유형 조사" not in sht.title and year == "2025":
            continue
        if year == "2026" and not (
            "진학희망 및 지원유형 조사" in sht.title or sht.title.isdigit()
        ):
            continue

        rows = sht.get_all_values()
        if len(rows) < 3:
            continue

        # 반 번호 추출
        if year == "2025":
            try:
                cls_num = sht.title.split("(")[1].split(")")[0]  # (301) → 301
            except:
                continue
        else:
            cls_num = sht.title if sht.title.isdigit() else None
            if not cls_num:
                continue

        students_in_class = 0
        for r in rows[2:]:
            if len(r) < 3 or not r[2].strip():
                continue

            key = (cls_num, r[1])  # (반, 번호)
            types = detect_types(r)

            if types:
                student_data[key] = {
                    "name": r[2],
                    "gender": r[3] if len(r) > 3 else "",
                    "types": types,
                }

                for t in types:
                    type_stats[t] += 1

                students_in_class += 1

        print(f"  → {cls_num}반: {students_in_class}명 ({len([k for k in student_data if k[0] == cls_num])}명)")

    total_students = len(student_data)
    print(f"  → 총 {total_students}명 감지 완료")

    # ── Step 2. 입시_트래킹에서 최종 결과 정보 수집 ──────
    print(f"\n[2/3] 입시_트래킹에서 최종 결과 정보 수집 중...")

    tracking_sht = ss.worksheet("입시_트래킹")
    tracking_rows = tracking_sht.get_all_values()

    # (반, 번호) → {유형, 1차, 2차, 최종}
    tracking_data = {}

    for row in tracking_rows[1:]:
        if len(row) < 3 or not row[2]:
            continue
        key = (row[0], row[1])  # (반, 번호)
        tracking_data[key] = {
            "type": row[4] if len(row) > 4 else "",
            "1차": row[7] if len(row) > 7 else "",
            "2차": row[8] if len(row) > 8 else "",
            "최종": row[9] if len(row) > 9 else "",
        }

    print(f"  → {len(tracking_data)}명의 최종 결과 정보 로드 완료")

    # ── Step 3. 입시 진행 현황 데이터 생성 ──────────────
    print(f"\n[3/3] 입시 진행 현황 시트 생성 중...")

    progress_data = [
        ["반", "번호", "성명", "성별", "지원유형", "1차", "2차", "최종", "비고"]
    ]

    matched_count = 0
    for (cls_num, student_num), info in sorted(student_data.items()):
        name = info["name"]
        gender = info["gender"]
        types = info["types"]

        # 각 지원 유형별로 한 행씩 생성
        for type_name in types:
            row_data = [
                cls_num,
                student_num,
                name,
                gender,
                type_name,
                "",
                "",
                "",
                ""
            ]

            # 해당 학생의 최종 유형과 일치하면 입시_트래킹 데이터 추가
            key = (cls_num, student_num)
            if key in tracking_data:
                tracking_info = tracking_data[key]
                final_type = tracking_info["type"]

                # 현재 지원유형이 최종 유형과 일치하면
                if type_name == final_type:
                    row_data[5] = tracking_info["1차"]
                    row_data[6] = tracking_info["2차"]
                    row_data[7] = tracking_info["최종"]
                    matched_count += 1

            progress_data.append(row_data)

    print(f"  → {len(progress_data)-1}행 생성 완료 ({matched_count}명과 입시_트래킹 매칭)")

    # ── Step 4. 입시 진행 현황 시트 업데이트 ──────────
    try:
        progress_sht = ss.worksheet("입시 진행 현황")
        progress_sht.clear()
        print(f"  → 기존 입시 진행 현황 시트 초기화")
    except gspread.exceptions.WorksheetNotFound:
        progress_sht = ss.add_worksheet("입시 진행 현황", rows=5000, cols=10)
        print(f"  → 새 입시 진행 현황 시트 생성")

    progress_sht.update(values=progress_data, range_name="A1")

    print(f"\n{'='*70}")
    print(f" 완료! ✅")
    print(f"")
    print(f" 입시 진행 현황 시트가 생성/업데이트됐습니다.")
    print(f" - 총 {len(progress_data)-1}행 (학생-유형별 조합)")
    print(f" - {matched_count}명이 입시_트래킹과 매칭됨")
    print(f"")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
