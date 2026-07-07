#!/usr/bin/env python3
"""
입시 진행 현황 시트 생성 (입시_트래킹 기반)

입시_트래킹의 최종 지원 유형 데이터를 활용하여
입시 진행 현황 시트 생성

현재 입시_트래킹에는 이미 "최종 유형"이 결정된 상태
이를 입시 진행 현황으로 그대로 전환

실행:
  python build_progress_from_tracking.py 2026
  또는
  python build_progress_from_tracking.py [SPREADSHEET_ID]
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


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    ss_id = SPREADSHEET_IDS.get(year)

    if not ss_id:
        print(f"❌ 잘못된 연도: {year}")
        return

    print(f"{'='*70}")
    print(f" {year}학년도 입시 진행 현황 시트 생성")
    print(f" (입시_트래킹 기반)")
    print(f"{'='*70}")

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    # ── Step 1. 입시_트래킹 데이터 수집 ──────────────────
    print(f"\n[1/2] 입시_트래킹 시트에서 데이터 수집 중...")

    tracking_sht = ss.worksheet("입시_트래킹")
    tracking_rows = tracking_sht.get_all_values()

    header = tracking_rows[0]
    print(f"  헤더: {header}")

    # 컬럼 인덱스 확인
    col_map = {h: i for i, h in enumerate(header)}
    required_cols = ["반", "번호", "성명", "성별"]

    for col in required_cols:
        if col not in col_map:
            print(f"❌ '{col}' 컬럼을 찾을 수 없습니다.")
            return

    def get(row, name):
        idx = col_map.get(name)
        return str(row[idx]).strip() if idx is not None and idx < len(row) else ""

    def slot_cell(row, school_col, result_col):
        school, result = get(row, school_col), get(row, result_col)
        return f"{school} {result}".strip()

    # 입시_트래킹 데이터 읽기
    progress_data = [
        ["반", "번호", "성명", "성별", "희망유형", "영재고", "전기", "후기", "최종배정", "비고"]
    ]

    progress_rows = []
    for row in tracking_rows[1:]:
        hope = get(row, "희망유형")
        gifted = slot_cell(row, "영재고_접수", "영재고_결과")
        early = slot_cell(row, "전기_접수학교", "전기_결과")
        late = slot_cell(row, "후기_접수학교", "후기_결과")
        assigned = get(row, "최종배정학교")
        if not (hope or gifted or early or late or assigned):
            continue
        progress_rows.append([
            get(row, "반"), get(row, "번호"), get(row, "성명"), get(row, "성별"),
            hope, gifted, early, late, assigned, get(row, "데이터상태"),
        ])

    valid_count = len(progress_rows)
    for row_data in progress_rows:
        progress_data.append(row_data)

    print(f"  → {valid_count}명의 데이터 수집 완료")

    # ── Step 2. 입시 진행 현황 시트 생성/업데이트 ──────────
    print(f"\n[2/2] 입시 진행 현황 시트 생성/업데이트 중...")

    try:
        progress_sht = ss.worksheet("입시 진행 현황")
        progress_sht.clear()
        print(f"  → 기존 입시 진행 현황 시트 초기화")
    except gspread.exceptions.WorksheetNotFound:
        progress_sht = ss.add_worksheet("입시 진행 현황", rows=5000, cols=10)
        print(f"  → 새 입시 진행 현황 시트 생성")

    # 데이터 업로드
    progress_sht.update(values=progress_data, range_name="A1")
    print(f"  → {len(progress_data)-1}행 데이터 업로드 완료")

    # ── Step 3. 통계 출력 ──────────────────────────────
    print(f"\n[3/2] 통계 계산 중...")

    from collections import Counter
    hope_stats = Counter(row[4] for row in progress_data[1:] if row[4])

    print(f"\n  희망 유형별 분포:")
    for school_type in sorted(hope_stats.keys()):
        count = hope_stats[school_type]
        pct = count / valid_count * 100 if valid_count > 0 else 0
        print(f"    {school_type}: {count}명 ({pct:.1f}%)")

    print(f"\n{'='*70}")
    print(f" 완료! ✅")
    print(f"")
    print(f" 입시 진행 현황 시트가 생성/업데이트됐습니다.")
    print(f" - 총 {valid_count}명의 데이터 (최종 지원 유형별)")
    print(f"")
    print(f" 📌 참고:")
    print(f"    - 각 학생은 '최종 지원 유형'으로 1행으로 표시됩니다")
    print(f"    - 복수 지원 이력은 반별 시트 입력 후 추가 분석 예정")
    print(f"")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
