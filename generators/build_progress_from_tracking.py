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
    required_cols = ["반", "번호", "성명", "성별", "유형", "1차", "2차", "최종"]

    for col in required_cols:
        if col not in col_map:
            print(f"❌ '{col}' 컬럼을 찾을 수 없습니다.")
            return

    # 입시_트래킹 데이터 읽기
    progress_data = [
        ["반", "번호", "성명", "성별", "지원유형", "1차", "2차", "최종", "비고"]
    ]

    valid_count = 0
    for row in tracking_rows[1:]:
        if len(row) < 3 or not row[col_map["성명"]]:
            continue

        cls = row[col_map["반"]].strip()
        num = row[col_map["번호"]].strip()
        name = row[col_map["성명"]].strip()
        gender = row[col_map["성별"]].strip() if col_map["성별"] < len(row) else ""
        school_type = row[col_map["유형"]].strip() if col_map["유형"] < len(row) else ""
        first = row[col_map["1차"]].strip() if col_map["1차"] < len(row) else ""
        second = row[col_map["2차"]].strip() if col_map["2차"] < len(row) else ""
        final = row[col_map["최종"]].strip() if col_map["최종"] < len(row) else ""

        # 최종 유형이 없으면 스킵
        if not school_type:
            continue

        # 지원학교, 학과 정보 수집 (비고에 추가)
        school = row[col_map.get("지원학교", -1)].strip() if col_map.get("지원학교", -1) < len(row) else ""
        major = row[col_map.get("학과", -1)].strip() if col_map.get("학과", -1) < len(row) else ""

        remark = ""
        if school:
            remark = f"지원: {school}"
        if major:
            remark += f" / {major}" if remark else f"학과: {major}"

        row_data = [
            cls,
            num,
            name,
            gender,
            school_type,
            first,
            second,
            final,
            remark
        ]

        progress_data.append(row_data)
        valid_count += 1

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

    type_stats = {}
    final_stats = {}

    for row in progress_data[1:]:
        school_type = row[4]
        final_result = row[7]

        if school_type not in type_stats:
            type_stats[school_type] = 0
        type_stats[school_type] += 1

        if final_result not in final_stats:
            final_stats[final_result] = 0
        final_stats[final_result] += 1

    print(f"\n  지원 유형별 분포:")
    for school_type in sorted(type_stats.keys()):
        count = type_stats[school_type]
        pct = count / valid_count * 100
        print(f"    {school_type}: {count}명 ({pct:.1f}%)")

    print(f"\n  최종 결과별 분포:")
    for result in sorted(final_stats.keys()):
        count = final_stats[result]
        pct = count / valid_count * 100 if valid_count > 0 else 0
        print(f"    {result}: {count}명 ({pct:.1f}%)")

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
