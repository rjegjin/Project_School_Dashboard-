#!/usr/bin/env python3
"""
inspect_2025_sources.py
2025학년도 입시 데이터 소스 실태 파악
- SS2 "전기고 최종 합불" raw 행 출력
- SS2 "후기고 최종 합불" raw 행 출력
- xlsx 파일 구조 + 샘플 행 출력
"""

import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import os

def get_sheets_client():
    """Google Sheets 클라이언트 생성"""
    KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

def inspect_final_results_sheets():
    """SS2의 전기고/후기고 최종 합불 시트 검사"""
    print("\n" + "=" * 80)
    print("SS2 '전기고 최종 합불' 시트 분석")
    print("=" * 80)

    try:
        client = get_sheets_client()
        doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')
        sht = doc.worksheet("전기고 최종 합불")
        rows = sht.get_all_values()

        print(f"전체 행 수: {len(rows)}")
        print(f"최대 컬럼 수: {max(len(r) for r in rows) if rows else 0}")
        print("\n[Raw 데이터 - 첫 25행]")
        for i, row in enumerate(rows[:25]):
            print(f"행 {i}: {row}")
    except Exception as e:
        print(f"ERROR: {e}")

    print("\n" + "=" * 80)
    print("SS2 '후기고 최종 합불' 시트 분석")
    print("=" * 80)

    try:
        client = get_sheets_client()
        doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')
        sht = doc.worksheet("후기고 최종 합불")
        rows = sht.get_all_values()

        print(f"전체 행 수: {len(rows)}")
        print(f"최대 컬럼 수: {max(len(r) for r in rows) if rows else 0}")
        print("\n[Raw 데이터 - 첫 25행]")
        for i, row in enumerate(rows[:25]):
            print(f"행 {i}: {row}")
    except Exception as e:
        print(f"ERROR: {e}")

def inspect_xlsx_file():
    """일반고 xlsx 파일 검사"""
    print("\n" + "=" * 80)
    print("일반고 배정 xlsx 파일 분석")
    print("=" * 80)

    final_file = "/home/rjegj/다운로드/2026학년도 후기고 최종결과.xlsx"

    if not os.path.exists(final_file):
        print(f"파일 없음: {final_file}")
        return

    try:
        df = pd.read_excel(final_file, sheet_name=0)
        print(f"파일: {final_file}")
        print(f"행 수: {len(df)}")
        print(f"컬럼 수: {len(df.columns)}")
        print(f"\n[컬럼명]")
        for i, col in enumerate(df.columns):
            print(f"  {i}: {col}")

        print(f"\n[첫 5행 데이터]")
        print(df.head(5).to_string())
    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    try:
        print("2025학년도 입시 데이터 원본 분석 시작...")
        inspect_final_results_sheets()
        inspect_xlsx_file()
        print("\n✓ 분석 완료")
    except KeyboardInterrupt:
        print("\n중단됨")
        sys.exit(1)
