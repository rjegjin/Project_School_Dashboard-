#!/usr/bin/env python3
"""
반별 시트(301~314)에 '다자녀(3인+)' 컬럼 추가

현재 구조:
  col 19 (T) = 장애
  col 20 (U) = 영재고 합불   ← 여기에 삽입

삽입 후:
  col 19 (T) = 장애
  col 20 (U) = 다자녀(3인+)  ← NEW
  col 21 (V) = 영재고 합불
  ...

주의: insertDimension은 해당 시트의 합불/1차/2차/최종 컬럼 위치를 오른쪽으로 밀지만
       sync_type_to_tracking 은 컬럼 7~15만 사용하므로 영향 없음.

실행:
  python add_dajanyeo_column.py 2026
  python add_dajanyeo_column.py 2026 --dry-run
"""

import sys
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_IDS = {
    "2026": "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0",
}

INSERT_AT_COL = 20   # 0-based; 장애(19) 바로 다음
NEW_COL_HEADER = "다자녀\n(3인+)"

# 헤더가 있는 행(0-based): HEADER1=0, HEADER2=1
HEADER2_ROW = 1   # 소분류 헤더 행


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    dry_run = "--dry-run" in sys.argv

    ss_id = SPREADSHEET_IDS.get(year)
    if not ss_id:
        print(f"❌ 지원하지 않는 연도: {year}")
        return

    print("=" * 60)
    print(f" {year} 반별 시트 다자녀 컬럼 추가{'  [DRY RUN]' if dry_run else ''}")
    print("=" * 60)

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheets_api = build("sheets", "v4", credentials=creds)
    ss = gc.open_by_key(ss_id)

    class_sheets = sorted(
        [s for s in ss.worksheets() if s.title.isdigit() and len(s.title) == 3 and s.title.startswith("3")],
        key=lambda s: s.title
    )
    print(f"\n대상 시트: {[s.title for s in class_sheets]}")
    print()

    for sht in class_sheets:
        print(f"  [{sht.title}]", end="")

        # 이미 추가됐는지 확인
        header_row = sht.row_values(2)  # HEADER2 (1-based = row 2)
        if "다자녀" in " ".join(header_row):
            print(" → 이미 '다자녀' 컬럼 존재, 건너뜀")
            continue

        # 현재 col 20 확인
        current_col20 = header_row[INSERT_AT_COL] if len(header_row) > INSERT_AT_COL else "(없음)"
        print(f" col{INSERT_AT_COL} 현재값='{current_col20}'", end="")

        if dry_run:
            print(" → [dry-run] 삽입 건너뜀")
            continue

        # 1) insertDimension: col 20 위치에 빈 열 삽입
        sheets_api.spreadsheets().batchUpdate(
            spreadsheetId=ss_id,
            body={
                "requests": [{
                    "insertDimension": {
                        "range": {
                            "sheetId": sht.id,
                            "dimension": "COLUMNS",
                            "startIndex": INSERT_AT_COL,   # 0-based, inclusive
                            "endIndex":   INSERT_AT_COL + 1,  # exclusive
                        },
                        "inheritFromBefore": True,
                    }
                }]
            }
        ).execute()

        # 2) HEADER2 행의 새 열에 헤더 텍스트 입력 (row 2, col U = index 20)
        col_letter = chr(65 + INSERT_AT_COL)  # 'U'
        cell_addr = f"{col_letter}{HEADER2_ROW + 1}"   # 1-based "U2"
        sht.update(values=[[NEW_COL_HEADER]], range_name=cell_addr)

        print(f" → ✅ 삽입 완료 ({col_letter}열)")

    print()
    if dry_run:
        print("DRY RUN 완료 — 실제 변경 없음")
    else:
        print("완료! 모든 반별 시트에 다자녀(3인+) 컬럼이 추가되었습니다.")
        print(f"이제 sync_special_to_tracking.py 가 col {INSERT_AT_COL}을 자동으로 읽습니다.")


if __name__ == "__main__":
    main()
