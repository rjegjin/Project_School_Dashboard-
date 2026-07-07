#!/usr/bin/env python3
"""
스프레드시트 시트명 간결화

변경 전: 진학희망 및 지원유형 조사(301)_Sheet1 → 변경 후: 301
        전기고 최종 합불 → 전기고_최종
        후기고 최종 합불 → 후기고_최종
        입시_트래킹 → 입시_트래킹 (유지)
"""

import gspread
from google.oauth2.service_account import Credentials

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"

# 시트명 변경 매핑
RENAME_MAP = {
    "진학희망 및 지원유형 조사(301)_Sheet1": "301",
    "진학희망 및 지원유형 조사(302)_Sheet1": "302",
    "진학희망 및 지원유형 조사(303)_Sheet1": "303",
    "진학희망 및 지원유형 조사(304)_Sheet1": "304",
    "진학희망 및 지원유형 조사(305)_Sheet1": "305",
    "진학희망 및 지원유형 조사(306)_Sheet1": "306",
    "진학희망 및 지원유형 조사(307)_Sheet1": "307",
    "진학희망 및 지원유형 조사(308)_Sheet1": "308",
    "진학희망 및 지원유형 조사(309)_Sheet1": "309",
    "진학희망 및 지원유형 조사(310)_Sheet1": "310",
    "진학희망 및 지원유형 조사(311)_Sheet1": "311",
    "진학희망 및 지원유형 조사(312)_Sheet1": "312",
    "진학희망 및 지원유형 조사(313)_Sheet1": "313",
    "진학희망 및 지원유형 조사(314)_Sheet1": "314",
    "전기고 최종 합불": "전기고_최종",
    "후기고 최종 합불": "후기고_최종",
}


def main():
    print("=" * 60)
    print(" 스프레드시트 시트명 간결화")
    print("=" * 60)

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SPREADSHEET_ID)

    print(f"\n[1/2] 현재 시트명 확인 중...")
    sheets = ss.worksheets()
    for sht in sheets:
        print(f"  {sht.title}")

    print(f"\n[2/2] 시트명 변경 중...")
    for sht in sheets:
        if sht.title in RENAME_MAP:
            new_name = RENAME_MAP[sht.title]
            try:
                sht.update_title(new_name)
                print(f"  ✅ '{sht.title}' → '{new_name}'")
            except Exception as e:
                print(f"  ❌ '{sht.title}' 변경 실패: {e}")
        else:
            print(f"  ⏭️  '{sht.title}' (유지)")

    print(f"\n" + "=" * 60)
    print(f" 완료!")
    print(f" URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print("=" * 60)


if __name__ == "__main__":
    main()
