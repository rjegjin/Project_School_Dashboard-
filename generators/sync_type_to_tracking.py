#!/usr/bin/env python3
"""
반별 지원 데이터에서 유형 자동 감지 → 입시_트래킹 시트 자동 채우기

실행:
  python sync_type_to_tracking.py 2026
  또는
  python sync_type_to_tracking.py [SPREADSHEET_ID]
"""

import sys
import gspread
from google.oauth2.service_account import Credentials

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 2025/2026 선택
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


def detect_type(row):
    """행 데이터에서 지원 유형 감지"""
    row = row + [""] * (30 - len(row))

    for priority_type in PRIORITY:
        for col_idx, type_name in TYPE_MAP.items():
            if type_name != priority_type:
                continue
            val = row[col_idx].strip()
            if val and val not in ["", "X", "x", "0"]:
                return type_name
            elif val in ["○", "O", "o"]:
                return type_name

    return ""


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    ss_id = SPREADSHEET_IDS.get(year)

    if not ss_id:
        print(f"❌ 잘못된 연도: {year}")
        print(f"   사용 가능: {', '.join(SPREADSHEET_IDS.keys())}")
        return

    print(f"{'='*60}")
    print(f" {year}학년도 유형 자동 감지 및 입시_트래킹 동기화")
    print(f"{'='*60}")

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    # ── Step 1. 반별 데이터 수집 및 유형 감지 ─────────
    print(f"\n[1/2] 반별 데이터에서 유형 감지 중...")

    # 반별 시트 찾기 (시트명이 "301"~"314" 형식)
    class_data = {}
    for sht in ss.worksheets():
        title = sht.title.strip()
        # "301"~"314" 숫자명 시트만 처리
        if not (title.isdigit() and len(title) == 3 and title.startswith("3")):
            continue

        cls_num = int(title[1:])  # "301" → 1, "314" → 14

        rows = sht.get_all_values()
        # 반별 시트는 2행 헤더 (row0=메인헤더, row1=서브헤더, row2+=학생)
        data_start = 2 if len(rows) > 2 and not rows[1][2].strip() else 1
        if len(rows) <= data_start:
            continue

        for r in rows[data_start:]:
            if len(r) < 3 or not r[2].strip():
                continue

            key = (r[0].strip(), r[1].strip())  # (반, 번호)
            school_type = detect_type(r)

            # 지원학교명: 유형이 감지된 컬럼의 값이 "O"/"o"/"○"이 아니면 학교명으로 사용
            school_name = ""
            for col_idx, type_name in TYPE_MAP.items():
                if type_name == school_type and col_idx < len(r):
                    val = r[col_idx].strip()
                    if val and val not in ["", "X", "x", "0", "O", "o", "○"]:
                        school_name = val
                    break

            class_data[key] = {
                "name": r[2],
                "gender": r[3] if len(r) > 3 else "",
                "type": school_type,
                "school": school_name,
            }

        count = len([k for k in class_data if k[0] == str(cls_num)])
        print(f"  → {title}반: {count}명")

    total = len(class_data)
    print(f"  → 총 {total}명 유형 감지 완료")

    # ── Step 2. 입시_트래킹 시트 업데이트 ─────────────
    print(f"\n[2/2] 입시_트래킹 시트 업데이트 중...")

    tracking_sht = ss.worksheet("입시_트래킹")
    rows = tracking_sht.get_all_values()

    # 헤더 행 확인
    if len(rows) < 1:
        print("❌ 입시_트래킹 시트가 비어 있습니다.")
        return

    header = rows[0]

    # "유형" 컬럼 인덱스
    try:
        type_col_idx = header.index("유형")
    except ValueError:
        print("❌ '유형' 컬럼을 찾을 수 없습니다.")
        return

    # "지원학교" 컬럼 인덱스 (없으면 None)
    school_col_idx = header.index("지원학교") if "지원학교" in header else None

    # 데이터 업데이트
    updates = []
    updated_types = 0
    updated_schools = 0

    for i, row in enumerate(rows[1:], start=2):  # 1-based row index
        if len(row) < 3 or not row[2]:
            continue

        cls_num = row[0].strip()
        student_num = row[1].strip()
        key = (cls_num, student_num)

        if key not in class_data:
            continue

        student = class_data[key]

        # 유형 업데이트 (이미 값 있으면 덮어쓰지 않음)
        existing_type = row[type_col_idx].strip() if type_col_idx < len(row) else ""
        if student["type"] and not existing_type:
            cell_addr = f"{chr(65 + type_col_idx)}{i}"
            updates.append((cell_addr, student["type"]))
            updated_types += 1

        # 지원학교 업데이트 (컬럼 있고, 값 있고, 기존 비어 있을 때)
        if school_col_idx is not None and student["school"]:
            existing_school = row[school_col_idx].strip() if school_col_idx < len(row) else ""
            if not existing_school:
                cell_addr = f"{chr(65 + school_col_idx)}{i}"
                updates.append((cell_addr, student["school"]))
                updated_schools += 1

    # 배치 업데이트 (최대 100개씩)
    for batch in [updates[i:i+100] for i in range(0, len(updates), 100)]:
        update_data = [{"range": cell, "values": [[val]]} for cell, val in batch]
        if update_data:
            tracking_sht.batch_update(update_data)

    print(f"  → 유형 동기화: {updated_types}명 / 지원학교 동기화: {updated_schools}명")

    print(f"\n{'='*60}")
    print(f" 완료!")
    print(f" 입시_트래킹 시트의 '유형' 컬럼이 업데이트됐습니다.")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
