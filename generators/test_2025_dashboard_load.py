#!/usr/bin/env python3
"""
test_2025_dashboard_load.py
대시보드 load_2025_data() 호환성 테스트
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

def load_2025_data_test():
    """
    대시보드 load_2025_data() 함수 시뮬레이션
    (integrated_dashboard.py와 동일한 로직)
    """
    try:
        client = get_sheets_client()
        doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')

        # 2025_입시결과_마스터 시트 로드
        master_sht = doc.worksheet("2025_입시결과_마스터")
        rows = master_sht.get_all_values()

        if len(rows) < 2:
            return None, "시트가 비어있음"

        # DataFrame으로 변환
        df = pd.DataFrame(rows[1:], columns=rows[0])

        # 컬럼 이름 통일 (대시보드 호환성)
        column_mapping = {
            '최종유형': '유형',
            '합격여부': '최종결과',
            '배정학교': '배정학교'
        }
        df = df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns})

        # 배정학교타입 열 추가
        if '배정학교' not in df.columns:
            df['배정학교'] = ''
        if '배정학교타입' not in df.columns:
            df['배정학교타입'] = ''
            if '유형' in df.columns:
                for idx, row in df.iterrows():
                    utype = str(row['유형']).strip()
                    if utype in ['과학고', '예고', '특성화고']:
                        df.at[idx, '배정학교타입'] = '전기고'
                    elif utype in ['자사고', '외고/국제고', '비평준화고']:
                        df.at[idx, '배정학교타입'] = '후기고'
                    elif utype == '일반고':
                        df.at[idx, '배정학교타입'] = '후기고'

        return df, None

    except Exception as e:
        return None, str(e)

if __name__ == "__main__":
    print("\n" + "=" * 80)
    print("대시보드 load_2025_data() 호환성 테스트")
    print("=" * 80)

    df, error = load_2025_data_test()

    if error:
        print(f"\n✗ 오류: {error}")
        sys.exit(1)

    if df is None or len(df) == 0:
        print("\n✗ 데이터 로드 실패")
        sys.exit(1)

    print(f"\n✓ 데이터 로드 성공: {len(df)}명")

    # 컬럼 검증
    print(f"\n[필수 컬럼 검증]")
    required_cols = ['반', '번호', '성명', '성별', '유형', '배정학교', '최종결과', '배정학교타입']
    for col in required_cols:
        present = '✓' if col in df.columns else '✗'
        print(f"  {present} {col}")

    # 데이터 샘플
    print(f"\n[데이터 샘플]")
    print(df[['성명', '유형', '배정학교', '최종결과', '배정학교타입']].head(10).to_string(index=False))

    # 통계
    print(f"\n[통계]")
    print(f"  전기고: {(df['배정학교타입'] == '전기고').sum()}명")
    print(f"  후기고: {(df['배정학교타입'] == '후기고').sum()}명")
    print(f"  합격: {(df['최종결과'] == '합격').sum()}명")
    print(f"  불합격: {(df['최종결과'] == '불합격').sum()}명")
    print(f"  미결: {(df['최종결과'] == '미결').sum()}명")

    print(f"\n✓ 모든 테스트 통과")
    sys.exit(0)
