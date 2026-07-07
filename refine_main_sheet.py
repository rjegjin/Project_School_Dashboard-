# refine_main_sheet.py
import gspread
import os
import json
from google.oauth2.service_account import Credentials
import time

# --- Configuration ---
SERVICE_ACCOUNT_FILE = os.path.expanduser('~/.config/gspread/service_account.json')
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]
TARGET_SS_ID = '14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0'

# --- Helper Functions ---
def get_credentials():
    """Gets credentials from the service account file."""
    # Check for the secret file in the current directory's .secret folder first
    secret_file = os.path.join('.secret', 'service_account.json')
    if os.path.exists(secret_file):
        return Credentials.from_service_account_file(secret_file, scopes=SCOPES)
    # Fallback to the default gspread location
    elif os.path.exists(SERVICE_ACCOUNT_FILE):
        return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    raise FileNotFoundError(f"Service account key not found in .secret/ or {SERVICE_ACCOUNT_FILE}")

def connect_google_api():
    """Connects to Google Sheets API and returns the client."""
    creds = get_credentials()
    return gspread.authorize(creds)

# --- Main Logic ---
def main():
    """
    Updates the main sheet with a user-friendly layout, hyperlinks,
    and moves it to the first position.
    """
    try:
        print("🚀 Google Sheets API에 연결 중...")
        client = connect_google_api()
        sh = client.open_by_key(TARGET_SS_ID)
        print(f"✅ 스프레드시트 '{sh.title}'에 연결되었습니다.")

        # 1. Get all worksheets to build hyperlinks
        all_worksheets = sh.worksheets()
        spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{TARGET_SS_ID}/edit#gid="

        # 2. Prepare the new, more user-friendly content
        print("🎨 메인 시트 콘텐츠를 생성 중...")
        
        content = [
            ["🏫 2026학년도 목일중 고입 진학현황 관리 시스템"],
            [],
            ["안녕하세요, 선생님! 2026학년도 고입 진학현황 관리를 위한 통합 대시보드입니다."],
            ["아래 안내에 따라 각 반의 데이터를 손쉽게 입력하고 전체 현황을 확인하세요."],
            [],
            ["✅ Step 1: 담임 선생님용 - 반별 데이터 입력"],
            ["담당하시는 반의 번호를 클릭하여 학생들의 지원 현황을 입력해주세요. 실시간으로 중앙 관리 시트에 반영됩니다."],
        ]
        
        class_sheets = sorted([w for w in all_worksheets if w.title.isdigit()], key=lambda x: int(x.title))
        class_links_rows = []
        row = []
        for i, ws in enumerate(class_sheets):
            link_formula = f'=HYPERLINK("{spreadsheet_url}{ws.id}", "{ws.title}반 바로가기")'
            row.append(link_formula)
            if (i + 1) % 4 == 0 or (i + 1) == len(class_sheets):
                class_links_rows.append(row)
                row = []
        content.extend(class_links_rows)
        
        content.extend([
            [],
            ["📊 Step 2: 관리자용 - 전체 현황 확인"],
            ["아래 링크를 통해 전체 학생들의 입시 진행 상황과 최종 합격자 현황을 한눈에 파악할 수 있습니다."],
        ])

        # Find management sheets and create hyperlinks
        mgmt_sheets_to_find = ["입시_트래킹", "전기고_최종", "후기고_최종"]
        mgmt_links = []
        for title in mgmt_sheets_to_find:
            try:
                ws = sh.worksheet(title)
                mgmt_links.append(f'=HYPERLINK("{spreadsheet_url}{ws.id}", "{title}")')
            except gspread.WorksheetNotFound:
                print(f"경고: '{title}' 시트를 찾을 수 없어 링크를 생성하지 못했습니다.")
        content.append(mgmt_links)

        content.extend([
            [],
            ["💡 사용 방법 요약"],
            ["1. [반별 데이터 입력]에서 담당 반의 링크를 클릭합니다."],
            ["2. 해당 반 시트에서 학생의 지원 학교 칸에 'O' 또는 학교명을 입력합니다."],
            ["3. 전형 단계(1차, 2차, 최종) 결과가 나오면 해당 칸에 '합격', '불합격', '대기' 등으로 상태를 업데이트합니다."],
            ["4. 모든 내용은 자동으로 저장되며, '입시_트래킹' 시트에 취합됩니다."],
            [],
            ["⚠️ 주의사항"],
            ["- 학생의 개인정보 보호에 유의해주시기 바랍니다."],
            ["- '입시_트래킹', '전기고_최종', '후기고_최종' 시트는 자동 생성되므로 직접 수정하지 않는 것을 권장합니다."],
            ["- 문의사항은 정보 담당자에게 연락주세요."],
        ])

        # 3. Update the sheet
        main_sheet_title = '메인'
        try:
            main_sheet = sh.worksheet(main_sheet_title)
            print(f"'{main_sheet_title}' 시트를 업데이트하는 중...")
        except gspread.WorksheetNotFound:
            print(f"'{main_sheet_title}' 시트를 찾을 수 없어 새로 생성합니다.")
            main_sheet = sh.add_worksheet(title=main_sheet_title, rows=100, cols=20)

        main_sheet.clear()
        main_sheet.update(content, value_input_option='USER_ENTERED')
        print("✅ 콘텐츠 업데이트 완료.")

        # 4. Format the sheet for better readability
        print("💅 시트 서식을 적용하는 중...")
        requests = [
            {"updateSheetProperties": {"properties": {"sheetId": main_sheet.id, "gridProperties": {"frozenRowCount": 1}}, "fields": "gridProperties.frozenRowCount"}},
            {"repeatCell": {"range": {"sheetId": main_sheet.id, "startRowIndex": 0, "endRowIndex": 1}, "cell": {"userEnteredFormat": {"textFormat": {"fontSize": 18, "bold": True}}}, "fields": "userEnteredFormat(textFormat)"}},
            {"repeatCell": {"range": {"sheetId": main_sheet.id, "startRowIndex": 5, "endRowIndex": 6}, "cell": {"userEnteredFormat": {"textFormat": {"fontSize": 12, "bold": True}}}, "fields": "userEnteredFormat(textFormat)"}},
            {"repeatCell": {"range": {"sheetId": main_sheet.id, "startRowIndex": 12, "endRowIndex": 13}, "cell": {"userEnteredFormat": {"textFormat": {"fontSize": 12, "bold": True}}}, "fields": "userEnteredFormat(textFormat)"}},
            {"repeatCell": {"range": {"sheetId": main_sheet.id, "startRowIndex": 15, "endRowIndex": 16}, "cell": {"userEnteredFormat": {"textFormat": {"fontSize": 12, "bold": True}}}, "fields": "userEnteredFormat(textFormat)"}},
            {"repeatCell": {"range": {"sheetId": main_sheet.id, "startRowIndex": 21, "endRowIndex": 22}, "cell": {"userEnteredFormat": {"textFormat": {"fontSize": 12, "bold": True}}}, "fields": "userEnteredFormat(textFormat)"}},
            {"mergeCells": {"range": {"sheetId": main_sheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 5}}},
            {"updateDimensionProperties": {"range": {"sheetId": main_sheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 450}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {"range": {"sheetId": main_sheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 5}, "properties": {"pixelSize": 150}, "fields": "pixelSize"}},
        ]
        sh.batch_update({"requests": requests})
        print("✅ 서식 적용 완료.")

        # 5. Move the sheet to the first position
        print(f"'{main_sheet_title}' 시트를 맨 앞으로 이동 중...")
        # Re-fetch worksheets to ensure correct order after potential creation
        all_worksheets = sh.worksheets()
        if all_worksheets[0].title != main_sheet_title:
            main_sheet_obj = sh.worksheet(main_sheet_title)
            other_sheets = [w for w in all_worksheets if w.title != main_sheet_title]
            new_order = [main_sheet_obj] + other_sheets
            sh.reorder_worksheets(new_order)
            print("✅ 시트 이동 완료! '메인' 시트가 이제 첫 번째 탭입니다.")
        else:
            print("✅ '메인' 시트가 이미 첫 번째 위치에 있습니다.")


    except gspread.exceptions.APIError as e:
        print(f"❌ Google Sheets API 에러 발생: {e}")
    except Exception as e:
        print(f"❌ 예상치 못한 에러 발생: {e}")

if __name__ == '__main__':
    main()
