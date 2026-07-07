#!/usr/bin/env python3
"""
입시_트래킹 시트의 '최종' 컬럼에서 합격자 추출
→ 유형별로 분류해서 전기고/후기고 최종 합불 시트에 자동 입력

실행:
  python generate_final_sheets.py 2026
  또는
  python generate_final_sheets.py [SPREADSHEET_ID]
"""

import sys
from pathlib import Path
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials

# 슬롯 기반 학교 분류
sys.path.insert(0, str(Path(__file__).resolve().parent))
from school_types import classify_school

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_IDS = {
    "2025": "1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI",
    "2026": "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0",
}

# 전기고 유형
EARLY_TYPES = {"영재고", "과학고", "예술계고", "특성화고"}

# 최종 합불 시트 섹션 정의
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


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    ss_id = SPREADSHEET_IDS.get(year)

    if not ss_id:
        print(f"❌ 잘못된 연도: {year}")
        return

    print(f"{'='*60}")
    print(f" {year}학년도 최종 합불 시트 자동 생성")
    print(f"{'='*60}")

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    # ── Step 1. 입시_트래킹 데이터 수집 ────────────────
    print(f"\n[1/3] 입시_트래킹 데이터 수집 중...")

    tracking_sht = ss.worksheet("입시_트래킹")
    rows = tracking_sht.get_all_values()

    header = rows[0]
    col_map = {h: i for i, h in enumerate(header)}

    # 필수 컬럼 확인
    required = ["반", "번호", "성명", "성별"]
    for req in required:
        if req not in col_map:
            print(f"❌ '{req}' 컬럼을 찾을 수 없습니다.")
            return

    # 슬롯 컬럼 정의: (접수_컬럼, 결과_컬럼, 고정_유형)
    SLOT_FIELDS = [
        ("영재고_접수", "영재고_결과", "영재고"),
        ("전기_접수학교", "전기_결과", None),
        ("후기_접수학교", "후기_결과", None),
    ]

    def get(row, col_name=None, idx=None):
        col_idx = col_map.get(col_name) if col_name is not None else idx
        return row[col_idx].strip() if col_idx is not None and col_idx < len(row) else ""

    # classify_school 결과를 섹션 타입으로 매핑
    def map_to_section_type(classified_type):
        """classify_school 반환값을 EARLY/LATE_SECTIONS의 섹션 타입으로 변환. 미분류("") 반환 시 None."""
        if classified_type == "예술계고":
            return "예고"
        elif classified_type == "영재고":
            # ponytail: 영재고 slot 전용; EARLY_SECTIONS에 전용 섹션 없으므로 과학고 섹션으로 통합
            return "과학고"
        elif classified_type == "":
            return None  # 미분류 플래그: 별도 수집
        return classified_type

    passed = defaultdict(lambda: defaultdict(list))  # [early/late][type] = [students]
    unclassified = []  # 미분류 학교들 (조용한 누락 방지)

    for r in rows[1:]:
        if len(r) < 3 or not get(r, "성명"):
            continue

        for idx, (school_col, result_col, fixed_type) in enumerate(SLOT_FIELDS):
            school = get(r, school_col)
            result = get(r, result_col)

            # 접수가 있고 결과가 "최종합"인 경우만 합격 판정
            if not school or result != "최종합":
                continue

            # 학교 유형 결정: 고정 유형 또는 classify_school
            raw_type = fixed_type or classify_school(school.split(",")[0])

            # 섹션 타입으로 매핑 (미분류는 None)
            school_type = map_to_section_type(raw_type)

            if school_type is None:
                # 미분류: 경고 리스트에 수집, 시트엔 기재하지 않음
                slot_name = school_col.replace("_접수", "")
                unclassified.append({
                    "반": get(r, "반"),
                    "성명": get(r, "성명"),
                    "슬롯": slot_name,
                    "학교명": school,
                })
                continue

            student = {
                "반": get(r, "반"),
                "번호": get(r, "번호"),
                "성명": get(r, "성명"),
                "성별": get(r, "성별"),
                "유형": school_type,
                "지원학교": school,
                "학과": get(r, "학과"),
                "1차": "",
                "2차": "",
            }

            # 슬롯 위치로 early/late 결정: 영재고(0), 전기(1) = early; 후기(2) = late
            if idx < 2:
                passed["early"][school_type].append(student)
            else:
                passed["late"][school_type].append(student)

    early_total = sum(len(v) for v in passed["early"].values())
    late_total = sum(len(v) for v in passed["late"].values())
    print(f"  → 전기고 합격자: {early_total}명")
    print(f"  → 후기고 합격자: {late_total}명")

    # 미분류 경고 (조용한 누락 방지)
    if unclassified:
        print(f"\n{'⚠'*30}")
        print(f"⚠ 미분류 학교 — 시트 미기재, school_types.py SCHOOL_TYPE_OVERRIDES에 추가 후 재실행 필요:")
        for item in unclassified:
            print(f"  {item['반']}/{item['성명']} ({item['슬롯']}): {item['학교명']}")
        print(f"{'⚠'*30}\n")

    # ── Step 2. 전기고 최종 합불 시트 업데이트 ────────
    print(f"\n[2/3] 전기고 최종 합불 시트 업데이트 중...")

    early_sht = ss.worksheet("전기고_최종")
    early_sht.clear()

    early_data = []
    col_offset = 0

    for section_name, section_cols in EARLY_SECTIONS:
        # 섹션 헤더
        early_data.append([section_name] + [""] * (len(section_cols) - 1))
        early_data.append([""] * len(section_cols))
        early_data.append(section_cols)

        # 학생 데이터
        section_type = section_name.split(" / ")[0]  # "과학고", "예고", "특성화고"
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
            # 특성화고는 학과 컬럼 추가
            if section_type == "특성화고":
                row_data.insert(5, student.get("학과", ""))
            early_data.append(row_data)

        print(f"  → {section_name}: {len(students)}명")

    early_sht.update(values=early_data, range_name="A1")

    # ── Step 3. 후기고 최종 합불 시트 업데이트 ────────
    print(f"\n[3/3] 후기고 최종 합불 시트 업데이트 중...")

    late_sht = ss.worksheet("후기고_최종")
    late_sht.clear()

    late_data = []

    for section_name, section_cols in LATE_SECTIONS:
        # 섹션 헤더
        late_data.append([section_name] + [""] * (len(section_cols) - 1))
        late_data.append([""] * len(section_cols))
        late_data.append(section_cols)

        # 학생 데이터
        section_type = section_name.split(" / ")[0]  # "자사고", "외고/국제고", "비평준화고"
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

        print(f"  → {section_name}: {len(students)}명")

    late_sht.update(values=late_data, range_name="A1")

    print(f"\n{'='*60}")
    print(f" 완료!")
    print(f" 전기고 최종 합불 / 후기고 최종 합불 시트가 업데이트됐습니다.")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
