#!/usr/bin/env python3
"""
validate_2025_master.py
2025_입시결과_마스터 시트 검증 스크립트

- 마스터 시트 로드 확인
- 데이터 완결성 검증
- 통계 출력
"""

import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

def get_sheets_client():
    """Google Sheets 클라이언트"""
    KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

def validate():
    """마스터 시트 검증"""
    print("\n" + "=" * 80)
    print("2025_입시결과_마스터 시트 검증")
    print("=" * 80)

    try:
        client = get_sheets_client()
        doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')
        sht = doc.worksheet("2025_입시결과_마스터")
        rows = sht.get_all_values()

        if len(rows) < 2:
            print("✗ 시트가 비어있습니다")
            return False

        df = pd.DataFrame(rows[1:], columns=rows[0])

        print(f"\n✓ 시트 로드 성공: {len(df)}명")
        print(f"\n[컬럼]")
        for i, col in enumerate(df.columns):
            print(f"  {i}: {col}")

        # 검증
        print(f"\n[데이터 검증]")
        print(f"  총 학생 수: {len(df)}")

        # 필수 컬럼 확인
        required_cols = ['성명', '반', '번호', '최종유형', '배정학교', '합격여부']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            print(f"  ✗ 누락된 컬럼: {missing}")
            return False
        print(f"  ✓ 필수 컬럼 완성")

        # 성명 결측치
        empty_names = df[df['성명'].str.strip() == ''].shape[0]
        if empty_names > 0:
            print(f"  ⚠ 성명 공란: {empty_names}명")
        else:
            print(f"  ✓ 성명 완전")

        # 최종유형 분포
        print(f"\n[최종유형 분포]")
        type_dist = df['최종유형'].value_counts().sort_values(ascending=False)
        for utype, count in type_dist.items():
            print(f"  {utype if utype else '(미분류)'}: {count}명")

        # 합격여부 분포
        print(f"\n[합격여부 분포]")
        result_dist = df['합격여부'].value_counts().sort_values(ascending=False)
        for result, count in result_dist.items():
            print(f"  {result}: {count}명")

        # 데이터출처 분포 (있으면)
        if '데이터출처' in df.columns:
            print(f"\n[데이터출처 분포]")
            source_dist = df['데이터출처'].value_counts().sort_values(ascending=False)
            for source, count in source_dist.items():
                print(f"  {source}: {count}명")

        # 샘플 데이터
        print(f"\n[샘플 데이터 - 첫 5행]")
        for idx, row in df.head(5).iterrows():
            print(f"  {row['성명']}: {row['최종유형']} → {row['배정학교']} ({row['합격여부']})")

        print(f"\n✓ 검증 완료")
        return True

    except Exception as e:
        print(f"✗ 오류: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    try:
        success = validate()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n중단됨")
        sys.exit(1)
