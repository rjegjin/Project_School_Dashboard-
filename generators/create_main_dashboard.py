#!/usr/bin/env python3
"""
메인 대시보드 시트 생성 (gid 직접 링크 방식)

- 상세한 시스템 안내
- 반별 시트로의 직접 링크 (gid 기반)
- 시트 맨 앞으로 이동
"""

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"

# 반별 시트 gid 매핑
GID_MAP = {
    301: 1002263472,
    302: 2019095679,
    303: 954874646,
    304: 912057250,
    305: 337025937,
    306: 180175872,
    307: 1612233584,
    308: 119626419,
    309: 900587049,
    310: 39575293,
    311: 2068817738,
    312: 294818561,
    313: 2098795156,
    314: 2140330898,
}


def main():
    import sys
    force = "--yes" in sys.argv or "--force" in sys.argv

    print("=" * 60)
    print(" 메인 대시보드 시트 생성 (gid 직접 링크 방식)")
    print("=" * 60)
    print()
    print("  ⚠️  수정 대상: 스프레드시트의 [📌 메인] 탭만")
    print("  ℹ️  입시_트래킹, 301~314반 시트 등 데이터 시트는 건드리지 않음")
    print()

    if not force:
        ans = input("  계속 진행하시겠습니까? (y/N): ").strip().lower()
        if ans != "y":
            print("  취소되었습니다.")
            return

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SPREADSHEET_ID)
    sheets_api = build("sheets", "v4", credentials=creds)

    # Step 1: 메인 시트 생성 또는 기존 시트 확인
    print("\n[1/3] 메인 시트 준비 중...")
    try:
        main_sht = ss.worksheet("📌 메인")
        main_sht.clear()
        print("  → 기존 메인 시트 초기화")
    except gspread.exceptions.WorksheetNotFound:
        main_sht = ss.add_worksheet("📌 메인", rows=100, cols=3)
        print("  → 메인 시트 생성 완료")

    # Step 2: 콘텐츠 작성 (URL 링크 사용)
    print("\n[2/3] 대시보드 콘텐츠 작성 중...")

    base_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit?gid="

    data = [
        ["2026학년도 목일중 고입 진학현황 관리"],
        [""],
        ["아래에서 자신의 반을 클릭하여 학생 정보를 입력하세요."],
        [""],
    ]

    # 반별 시트 링크
    data.append(["반", ""])

    # 301~314
    for cls_num in range(301, 315):
        gid = GID_MAP[cls_num]
        cls_display = f"3-{cls_num % 100}"
        link = f"{base_url}{gid}"
        data.append([f"{cls_display}반", link])

    data.extend([
        [""],
        ["입력 주의사항"],
        [""],
        ["학교명: 띄어쓰기/괄호 없이 정확하게 (한성과고 O, 한성 과고 X)"],
        ["결과: 합격/불합격 또는 빈칸"],
        ["모를 경우: ○ 표시만 하면 관리자가 확인합니다"],
        [""],
        ["문의: 관리자에게 연락"],
    ])

    # 데이터 쓰기
    main_sht.update(values=data, range_name="A1")
    print(f"  → 콘텐츠 작성 완료 ({len(data)}행)")

    # Step 3: 메인 시트를 맨 앞으로 이동
    print("\n[3/3] 메인 시트를 맨 앞으로 이동 중...")

    meta = sheets_api.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    main_sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == "📌 메인":
            main_sheet_id = s["properties"]["sheetId"]
            break

    if main_sheet_id is not None:
        sheets_api.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": main_sheet_id,
                                "index": 0,
                            },
                            "fields": "index",
                        }
                    }
                ]
            }
        ).execute()
        print(f"  → 메인 시트가 맨 앞으로 이동되었습니다")

    print(f"\n" + "=" * 60)
    print(f" 완료! ✅")
    print(f"")
    print(f" URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print(f"")
    print(f" 시트를 열면 '📌 메인' 탭이 가장 앞에 있습니다.")
    print(f" 담임들은 메인 시트의 반별 입력 시트 표에서 URL을 복사하여")
    print(f" 브라우저에 붙여넣거나, 반 이름을 클릭하여 이동할 수 있습니다.")
    print("=" * 60)


if __name__ == "__main__":
    main()
