#!/usr/bin/env python3
"""
2026학년도 목일중 고입 진학현황 스프레드시트 초기화
- 2025 양식 구조 복제
- 2026 학생 명렬표에서 학생 데이터 삽입 (지원 내용은 빈칸)
- 1차 / 2차 / 최종 트래킹 컬럼 추가
- 입시_트래킹 마스터 시트 생성
- rjegjin@gmail.com 에 편집 권한 공유
"""

import time
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread

# ── 설정 ──────────────────────────────────────────────
KEY_FILE    = "/home/rjegj/projects/.secrets/service_key.json"
ROSTER_ID   = "1QiniYj9Q4J0MR1y5TDKTydQdft6WoOEKb4u0P-AJnRI"
USER_EMAIL  = "rjegjin@gmail.com"
SCOPES      = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── 사전 준비 안내 ────────────────────────────────────
# 서비스 계정 Drive 할당량 초과로 직접 시트 생성 불가.
# 실행 전:
#   1. rjegjin@gmail.com 으로 빈 Google Sheets 파일 생성
#   2. 공유 → school-bot@gen-lang-client-0367740438.iam.gserviceaccount.com 편집자 추가
#   3. 아래 TARGET_SS_ID 에 생성된 시트 ID 입력 후 실행
TARGET_SS_ID = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"

# ── 헤더 정의 ─────────────────────────────────────────
# Row 1: 대분류
HEADER1 = [
    "반", "번호", "성명", "성별",
    "유형(일반은 체크 안함)", "", "",
    "지원 희망교(구체적 희망고가 있으면 교명, 아니면 O)",
    "", "", "", "", "", "", "", "",
    "특이사항", "", "", "",
    "영재고", "과학고", "자사고",
    "1차", "2차", "최종",
]
# Row 2: 소분류
HEADER2 = [
    "", "", "", "",
    "사회통합\n전형", "특례", "보훈",
    "영재고", "과학고", "예술계고",
    "마이스터고\n특성화고", "특성,마이스터 학과",
    "자사고", "국제고\n외국어고", "일반고", "기타\n(대안)",
    "쌍둥이", "학폭", "교직원\n자녀", "장애",
    "", "", "",
    "", "", "",
]

NUM_COLS = len(HEADER1)  # 26 (A~Z)

# ── 헬퍼 ──────────────────────────────────────────────
def col_letter(idx):
    """0-based index → 열 문자 (A, B, … Z)"""
    return chr(65 + idx)

def a1(row, col):
    """1-based row, 0-based col → A1 표기"""
    return f"{col_letter(col)}{row}"

def range_a1(r1, c1, r2, c2):
    return f"{a1(r1, c1)}:{a1(r2, c2)}"


def build_class_sheet_data(cls_num, students):
    """반 시트에 들어갈 전체 데이터 반환 (header 2행 + 학생 데이터)"""
    rows = [HEADER1[:], HEADER2[:]]
    for s in students:
        row = [s["class"], s["num"], s["name"], s["gender"]] + [""] * (NUM_COLS - 4)
        rows.append(row)
    return rows


def make_merge_requests(sheet_id):
    """반 시트 헤더 셀 병합 batchUpdate requests"""
    # (startRow, endRow, startCol, endCol) — 0-based, endRow/endCol exclusive
    merges = [
        (0, 2, 0, 1),   # A1:A2  반
        (0, 2, 1, 2),   # B1:B2  번호
        (0, 2, 2, 3),   # C1:C2  성명
        (0, 2, 3, 4),   # D1:D2  성별
        (0, 1, 4, 7),   # E1:G1  유형
        (0, 1, 7, 16),  # H1:P1  지원 희망교
        (0, 1, 16, 20), # Q1:T1  특이사항
        (0, 2, 20, 21), # U1:U2  영재고(합불)
        (0, 2, 21, 22), # V1:V2  과학고(합불)
        (0, 2, 22, 23), # W1:W2  자사고(합불)
        (0, 2, 23, 24), # X1:X2  1차
        (0, 2, 24, 25), # Y1:Y2  2차
        (0, 2, 25, 26), # Z1:Z2  최종
    ]
    requests = []
    for sr, er, sc, ec in merges:
        requests.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": sr, "endRowIndex": er,
                    "startColumnIndex": sc, "endColumnIndex": ec,
                },
                "mergeType": "MERGE_ALL",
            }
        })
    return requests


def make_format_requests(sheet_id, num_students):
    """헤더 볼드/배경색, 행 고정, 열 너비 설정"""
    requests = []

    # 헤더 2행 볼드 + 배경 (연한 회색)
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0, "endRowIndex": 2,
                "startColumnIndex": 0, "endColumnIndex": NUM_COLS,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.93, "green": 0.93, "blue": 0.93},
                    "textFormat": {"bold": True},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "wrapStrategy": "WRAP",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)",
        }
    })

    # 1차/2차/최종 컬럼 배경 강조 (연한 노랑)
    for col_idx in [23, 24, 25]:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0, "endRowIndex": 2 + num_students,
                    "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 1.0, "green": 0.98, "blue": 0.8},
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    # 상위 2행 고정
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": 2},
            },
            "fields": "gridProperties.frozenRowCount",
        }
    })

    # 열 너비 조정
    widths = {
        0: 45,  # 반
        1: 50,  # 번호
        2: 75,  # 성명
        3: 50,  # 성별
        7: 120, # 영재고 지원
        8: 120, # 과학고
        9: 130, # 예술계고
        10: 120, # 특성화
        11: 130, # 학과
        12: 120, # 자사고
        13: 130, # 외고
        14: 130, # 일반고
        23: 80, # 1차
        24: 80, # 2차
        25: 80, # 최종
    }
    for col_idx, width in widths.items():
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": col_idx,
                    "endIndex": col_idx + 1,
                },
                "properties": {"pixelSize": width},
                "fields": "pixelSize",
            }
        })

    return requests


def create_tracking_sheet_data(class_students):
    """입시_트래킹 마스터 시트 데이터 구성"""
    header = ["반", "번호", "성명", "성별", "유형", "지원학교", "학과", "1차", "2차", "최종", "비고"]
    rows = [header]
    for cls in sorted(class_students.keys()):
        for s in class_students[cls]:
            rows.append([s["class"], s["num"], s["name"], s["gender"],
                         "", "", "", "", "", "", ""])
    return rows


def create_result_sheet(title, sections):
    """전기고/후기고 최종 합불 양식 헤더 생성"""
    row1, row2, row3 = [], [], []
    for section_name, cols in sections:
        n = len(cols)
        row1 += [section_name] + [""] * (n - 1)
        row2 += [""] * n
        row3 += cols
    return [row1, row2, row3]


# ── 메인 ──────────────────────────────────────────────
def main():
    print("=" * 55)
    print(" 2026학년도 고입 진학현황 스프레드시트 초기화")
    print("=" * 55)

    creds         = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc            = gspread.authorize(creds)
    sheets_svc    = build("sheets", "v4", credentials=creds)
    drive_svc     = build("drive",  "v3", credentials=creds)

    # ── Step 1. 명렬표 데이터 수집 ───────────────────────
    print("\n[1/5] 2026 학생 명렬표 수집 중...")
    roster_sht = gc.open_by_key(ROSTER_ID).worksheet("2026년 전체 학생 명렬표")
    class_students = {}
    for r in roster_sht.get_all_values()[1:]:
        if r[0] != "3":
            continue
        cls = int(r[1])
        class_students.setdefault(cls, []).append(
            {"class": r[1], "num": r[2], "name": r[3], "gender": r[4]}
        )
    total = sum(len(v) for v in class_students.values())
    print(f"  → 3학년 {total}명 / {len(class_students)}개 반")

    # ── Step 2. 스프레드시트 준비 ────────────────────────
    print("\n[2/5] 스프레드시트 준비 중...")
    if not TARGET_SS_ID:
        print("  ❌ TARGET_SS_ID 가 비어 있습니다.")
        print("     init_2026_spreadsheet.py 상단의 TARGET_SS_ID 에 시트 ID를 입력해주세요.")
        return

    ss_id = TARGET_SS_ID

    # 제목 변경
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ss_id,
        body={"requests": [{"updateSpreadsheetProperties": {
            "properties": {"title": "2026학년도 목일중 고입 진학 현황"},
            "fields": "title"
        }}]}
    ).execute()

    # 기존 시트 목록 확인
    meta = sheets_svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    # 필요한 시트 목록
    sheet_titles = (
        [f"진학희망 및 지원유형 조사(3{str(c).zfill(2)})_Sheet1" for c in sorted(class_students.keys())]
        + ["전기고 최종 합불", "후기고 최종 합불", "입시_트래킹"]
    )

    # 없는 시트 추가
    add_requests = [
        {"addSheet": {"properties": {"title": t}}}
        for t in sheet_titles if t not in existing
    ]
    if add_requests:
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=ss_id, body={"requests": add_requests}
        ).execute()

    # 시트 ID 맵 재조회
    meta = sheets_svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    # 기존 기본 "시트1" 또는 "Sheet1" 삭제
    for default_name in ["시트1", "Sheet1"]:
        if default_name in sheet_map:
            sheets_svc.spreadsheets().batchUpdate(
                spreadsheetId=ss_id,
                body={"requests": [{"deleteSheet": {"sheetId": sheet_map[default_name]}}]}
            ).execute()
            print(f"  → 기본 시트 '{default_name}' 삭제")

    # 시트 ID 맵 최종 재조회
    meta = sheets_svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    print(f"  → 준비 완료: https://docs.google.com/spreadsheets/d/{ss_id}")

    # ── Step 3. 반별 시트 데이터 입력 + 서식 ─────────────
    print("\n[3/5] 반별 시트 데이터 입력 중...")
    all_requests = []

    for cls in sorted(class_students.keys()):
        title    = f"진학희망 및 지원유형 조사(3{str(cls).zfill(2)})_Sheet1"
        sid      = sheet_map[title]
        students = class_students[cls]
        data     = build_class_sheet_data(cls, students)

        # 데이터 쓰기
        sht = gc.open_by_key(ss_id).worksheet(title)
        sht.update(values=data, range_name=f"A1:{col_letter(NUM_COLS - 1)}{len(data)}",
                   value_input_option="USER_ENTERED")

        all_requests += make_merge_requests(sid)
        all_requests += make_format_requests(sid, len(students))
        print(f"  → 3-{cls}반 ({len(students)}명) 완료")
        time.sleep(0.3)  # API 속도 제한 방지

    # ── Step 4. 결과 시트 및 트래킹 시트 입력 ────────────
    print("\n[4/5] 결과 시트 / 트래킹 시트 입력 중...")

    # 전기고 최종 합불
    early_data = create_result_sheet("", [
        ("과학고",           ["연번", "반", "이름", "성별", "지원교", "1차", "2차", "최종"]),
        ("예고",             ["연번", "반", "성명", "성별", "예술계고", "1차", "2차", "최종"]),
        ("특성화고 / 마이스터고", ["연번", "반", "이름", "성별", "지원고등학교", "배정학과", "1차", "2차", "최종"]),
    ])
    gc.open_by_key(ss_id).worksheet("전기고 최종 합불").update(values=early_data, range_name="A1")

    # 후기고 최종 합불
    late_data = create_result_sheet("", [
        ("자사고",              ["연번", "반", "이름", "성별", "지원교", "1차", "2차", "최종"]),
        ("외고/국제고",         ["연번", "반", "성명", "성별", "지원교", "1차", "2차", "최종"]),
        ("비평준화고 / 중점고", ["연번", "반", "이름", "성별", "지원고등학교", "1차", "2차", "최종"]),
    ])
    gc.open_by_key(ss_id).worksheet("후기고 최종 합불").update(values=late_data, range_name="A1")

    # 입시_트래킹
    tracking_data = create_tracking_sheet_data(class_students)
    tracking_sht  = gc.open_by_key(ss_id).worksheet("입시_트래킹")
    tracking_sht.update(values=tracking_data, range_name="A1")

    # 트래킹 시트 서식
    t_sid = sheet_map["입시_트래킹"]
    all_requests += [
        {
            "repeatCell": {
                "range": {"sheetId": t_sid, "startRowIndex": 0, "endRowIndex": 1,
                          "startColumnIndex": 0, "endColumnIndex": 11},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": {"red": 0.27, "green": 0.51, "blue": 0.71},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    "horizontalAlignment": "CENTER",
                }},
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        },
        {
            "updateSheetProperties": {
                "properties": {"sheetId": t_sid, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
    ]
    # 1차/2차/최종 컬럼 배경 강조
    for col_idx in [7, 8, 9]:
        all_requests.append({
            "repeatCell": {
                "range": {"sheetId": t_sid, "startRowIndex": 0,
                          "endRowIndex": total + 1,
                          "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": {"red": 1.0, "green": 0.98, "blue": 0.8},
                }},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    print("  → 결과/트래킹 시트 완료")

    # ── Step 5. 서식 일괄 적용 + 공유 ───────────────────
    print("\n[5/5] 서식 적용 및 공유 처리 중...")

    # batchUpdate 50개씩 나눠서 전송 (API 한도)
    chunk = 50
    for i in range(0, len(all_requests), chunk):
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=ss_id,
            body={"requests": all_requests[i:i + chunk]}
        ).execute()
        time.sleep(0.5)

    print(f"  → {USER_EMAIL} 이 직접 생성한 시트이므로 별도 공유 불필요")

    print("\n" + "=" * 55)
    print(" 완료!")
    print(f" URL: https://docs.google.com/spreadsheets/d/{ss_id}")
    print("=" * 55)


if __name__ == "__main__":
    main()
