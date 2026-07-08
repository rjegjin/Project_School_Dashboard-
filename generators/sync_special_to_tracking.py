#!/usr/bin/env python3
"""
반별 시트(301~314) 특별전형 컬럼 → 특별전형_트래킹 시트 동기화

특별전형 7종:
  사회통합전형(E), 특례(F), 보훈(G)
  쌍둥이(Q), 학폭(R), 교직원자녀(S), 장애(T)

실행:
  python sync_special_to_tracking.py 2026
"""

import sys
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

# 반별 시트 0-based 컬럼 인덱스 → 특별전형 이름
# 다자녀 컬럼(20)은 add_dajanyeo_column.py 실행 후 추가됨
SPECIAL_COLS = {
    4:  "사회통합전형",
    5:  "특례",
    6:  "보훈",
    16: "쌍둥이",
    17: "학폭",
    18: "교직원자녀",
    19: "장애",
    20: "다자녀(3인+)",
}

SPECIAL_NAMES = list(SPECIAL_COLS.values())  # 컬럼 순서 유지

TRACKING_SHEET = "특별전형_트래킹"
TRACKING_HEADER = ["반", "번호", "이름", "성별"] + SPECIAL_NAMES

# 체크된 것으로 인정하는 값
POSITIVE_VALUES = {"o", "O", "○", "V", "v", "✓", "1", "y", "Y", "예", "있음"}


def is_checked(val: str) -> bool:
    return val.strip() in POSITIVE_VALUES


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    ss_id = SPREADSHEET_IDS.get(year)
    if not ss_id:
        print(f"❌ 잘못된 연도: {year}")
        return

    print("=" * 60)
    print(f" {year} 특별전형_트래킹 동기화")
    print("=" * 60)

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    # ── Step 1. 반별 시트에서 특별전형 데이터 수집 ────────────
    print("\n[1/2] 반별 시트 특별전형 데이터 수집 중...")

    students = []  # list of dicts

    # ponytail: 14개 시트 개별 read를 1회 batch로 통합
    class_titles = [
        sht.title.strip()
        for sht in sorted(ss.worksheets(), key=lambda s: s.title)
        if sht.title.strip().isdigit() and len(sht.title.strip()) == 3 and sht.title.strip().startswith("3")
    ]
    if not class_titles:
        total_students = 0
        total_special = 0
        print(f"  → 전체 {total_students}명 중 특별전형 해당 {total_special}명")
    else:
        # ponytail: 80행 상한 (범위 확장 필요시 A1:AB{MAX_ROWS} 수정)
        response = ss.values_batch_get(ranges=[f"'{t}'!A1:AB80" for t in class_titles])

        for title, value_range in zip(class_titles, response["valueRanges"]):
            rows = value_range.get("values", [])
            if len(rows) >= 80:
                print(f"⚠ {title}: 80행 상한 도달 — 범위 확장 필요")
            # 반별 시트 헤더: row0=HEADER1, row1=HEADER2, row2+=학생
            data_start = 2
            if len(rows) <= data_start:
                print(f"  → {title}반: 0명 특별전형 해당")
                continue

            cls_num = int(title[1:])  # "301" → 1

            found = 0
            for r in rows[data_start:]:
                # 최소 이름 있어야
                if len(r) < 3 or not r[2].strip():
                    continue

                special = {}
                has_any = False
                for col_idx, name in SPECIAL_COLS.items():
                    val = r[col_idx].strip() if col_idx < len(r) else ""
                    checked = is_checked(val)
                    special[name] = "O" if checked else ""
                    if checked:
                        has_any = True

                students.append({
                    "cls":    r[0].strip() or str(cls_num),
                    "num":    r[1].strip(),
                    "name":   r[2].strip(),
                    "gender": r[3].strip() if len(r) > 3 else "",
                    **special,
                    "_any":  has_any,
                })
                if has_any:
                    found += 1

            print(f"  → {title}반: {found}명 특별전형 해당")

        total_students = len(students)
        total_special = sum(1 for s in students if s["_any"])
        print(f"  → 전체 {total_students}명 중 특별전형 해당 {total_special}명")

    # ── Step 2. 특별전형_트래킹 시트 생성/갱신 ───────────────
    print(f"\n[2/2] {TRACKING_SHEET} 시트 업데이트 중...")

    try:
        tracking = ss.worksheet(TRACKING_SHEET)
        tracking.clear()
        print("  → 기존 시트 초기화")
    except gspread.exceptions.WorksheetNotFound:
        tracking = ss.add_worksheet(TRACKING_SHEET, rows=500, cols=len(TRACKING_HEADER))
        print("  → 새 시트 생성")

    # 전체 학생 기록 (특별전형 없는 경우 빈칸 행으로 포함 — 담임 확인용)
    write_rows = [TRACKING_HEADER]
    for s in students:
        row = [s["cls"], s["num"], s["name"], s["gender"]]
        row += [s.get(name, "") for name in SPECIAL_NAMES]
        write_rows.append(row)

    tracking.update(values=write_rows, range_name="A1")

    # 헤더 굵게 (formatting은 API 직접 호출 필요 — 생략 가능)
    print(f"  → {len(write_rows)-1}명 데이터 기록 완료")
    print(f"  → 특별전형 해당: {total_special}명")

    # 범주별 집계 출력
    print("\n[집계]")
    for name in SPECIAL_NAMES:
        cnt = sum(1 for s in students if s.get(name) == "O")
        if cnt:
            print(f"  · {name}: {cnt}명")

    print(f"\n{'='*60}")
    print(f" 완료! {TRACKING_SHEET} 시트 업데이트됨")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
