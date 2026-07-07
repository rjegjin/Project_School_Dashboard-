"""
설문지 응답 시트1 → 301~314 반별 시트 자동 동기화
- 이미 손으로 입력된 셀은 건드리지 않음 (덮어쓰기 없음)
- 폼 체크박스 응답(콤마 구분) 파싱하여 해당 열에 기입
"""

import sys
import re
import gspread
from google.oauth2.service_account import Credentials

# ── 인증 ────────────────────────────────────────────────────────────────────
SERVICE_KEY = "/home/rjegj/projects/.secrets/service_key.json"
SPREADSHEET_ID = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_sheet_client():
    creds = Credentials.from_service_account_file(SERVICE_KEY, scopes=SCOPES)
    return gspread.authorize(creds)

# ── 컬럼 인덱스 (0-based) ───────────────────────────────────────────────────
# 반별 시트 기준
COL_BAN     = 0   # A: 반
COL_NUM     = 1   # B: 번호
COL_NAME    = 2   # C: 성명
COL_GENDER  = 3   # D: 성별
COL_SOCIAL  = 4   # E: 사회통합전형
COL_SPECIAL = 5   # F: 특례
COL_BOHUN   = 6   # G: 보훈
COL_YOUNG   = 7   # H: 영재고
COL_SCI     = 8   # I: 과학고
COL_ART     = 9   # J: 예술계고
COL_MEISTER = 10  # K: 마이스터고/특성화고
COL_DEPT    = 11  # L: 특성,마이스터 학과
COL_JASA   = 12  # M: 자사고
COL_INTL   = 13  # N: 국제고/외국어고
COL_ILBAN  = 14  # O: 일반고
COL_ETC    = 15  # P: 기타(대안)
COL_TWIN   = 16  # Q: 쌍둥이
COL_VIOL   = 17  # R: 학폭
COL_TEACH  = 18  # S: 교직원자녀
COL_DISAB        = 19  # T: 장애(본인,형제자매)
COL_SPECIAL_NOTE = 20  # U: 특이사항 세부내용 (폼 신규 → 시트에 자동 추가)

# 데이터 행 시작 (헤더 2행 이후)
DATA_START_ROW = 2  # 0-based index

# ── 폼 → 시트 매핑 ──────────────────────────────────────────────────────────
# 지원 고등학교 유형 체크박스 값 → 반별 시트 컬럼 인덱스
SCHOOL_TYPE_MAP = {
    "영재고": COL_YOUNG,
    "영재학교": COL_YOUNG,
    "과학고": COL_SCI,
    "예술계고": COL_ART,
    "마이스터고": COL_MEISTER,
    "특성화고": COL_MEISTER,
    "마이스터고/특성화고": COL_MEISTER,
    "자사고": COL_JASA,
    "국제고": COL_INTL,
    "외국어고": COL_INTL,
    "국제고/외국어고": COL_INTL,
    "일반고": COL_ILBAN,
    "기타": COL_ETC,
    "기타(대안)": COL_ETC,
}

# 전형 구분 체크박스 값 → 컬럼 인덱스
JEONHYEONG_MAP = {
    "사회통합전형": COL_SOCIAL,
    "사회통합": COL_SOCIAL,
    "특례": COL_SPECIAL,
    "보훈": COL_BOHUN,
}

# 특별 기재 사항 체크박스 값 → 컬럼 인덱스
SPECIAL_NOTE_MAP = {
    "쌍둥이": COL_TWIN,
    "학폭": COL_VIOL,
    "교직원자녀": COL_TEACH,
    "교직원 자녀": COL_TEACH,
    "장애": COL_DISAB,
}

# ── 폼 응답 컬럼 인덱스 (0-based) ───────────────────────────────────────────
# A=0: 타임스탬프
# B=1: 섹션 헤더 (무시)
FORM_BAN        = 2   # C: 반 선택
FORM_NUM        = 3   # D: 번호
FORM_NAME       = 4   # E: 이름
FORM_TYPE       = 5   # F: 지원 고등학교 유형 (체크박스)
FORM_SCHOOL     = 6   # G: 지원 희망교 명칭
FORM_JEONHYEONG = 7   # H: 전형 구분 (체크박스)
FORM_SPECIAL        = 8   # I: 특별 기재 사항 (체크박스)
FORM_STAGE          = 9   # J: 현재 준비 단계 (반별 시트에 해당 열 없음 → 로그만)
FORM_DEPT           = 10  # K: 마이스터고/특성화고 학과명
FORM_SPECIAL_DETAIL = 11  # L: 특이사항 세부 내용


def parse_checkbox(value: str) -> list[str]:
    """체크박스 응답(콤마 구분)을 리스트로 파싱"""
    if not value or not value.strip():
        return []
    return [v.strip() for v in value.split(",") if v.strip()]


def ban_to_sheet_name(ban_value: str) -> str | None:
    """폼의 반 값 → 시트 이름 (예: '1반' 또는 '1' → '301')"""
    # '1반', '1', '301' 등 다양한 형식 처리
    match = re.search(r"\d+", str(ban_value))
    if not match:
        return None
    num = int(match.group())
    if 1 <= num <= 14:
        return f"3{num:02d}"
    if 301 <= num <= 314:
        return str(num)
    return None


def build_student_index(sheet_data: list[list]) -> dict[str, int]:
    """번호 → 행 인덱스 매핑 (data_start_row 이후 행만)"""
    index = {}
    for i, row in enumerate(sheet_data[DATA_START_ROW:], start=DATA_START_ROW):
        if len(row) > COL_NUM and row[COL_NUM]:
            try:
                num = int(str(row[COL_NUM]).strip())
                index[num] = i
            except ValueError:
                pass
    return index


def safe_get(row: list, col: int) -> str:
    """행에서 안전하게 값 읽기"""
    if col < len(row):
        return str(row[col]).strip()
    return ""


def sync_row_to_sheet(
    sheet_data: list[list],
    row_idx: int,
    updates: dict[int, str],  # {col_idx: value}
    sheet_name: str,
    student_name: str,
) -> list[tuple[int, int, str]]:
    """
    실제 변경사항 계산.
    빈 셀에만 기록 (덮어쓰기 없음).
    반환: [(row_idx, col_idx, value), ...]
    """
    changes = []
    row = sheet_data[row_idx]
    for col_idx, value in updates.items():
        current = safe_get(row, col_idx)
        if current == "":
            changes.append((row_idx, col_idx, value))
        else:
            print(f"  ⏭ [{sheet_name}] {student_name} 열{col_idx+1}: 이미 '{current}' 있음 → 스킵")
    return changes


def col_idx_to_a1(col_idx: int) -> str:
    """0-based 컬럼 인덱스 → A1 표기 (A, B, ..., Z, AA, ...)"""
    result = ""
    col_idx += 1
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def sync_all(dry_run: bool = False):
    """
    설문지 응답 → 301~314 반별 시트 동기화
    dry_run=True: 실제 쓰기 없이 변경 예정 내용만 출력
    """
    client = get_sheet_client()
    wb = client.open_by_key(SPREADSHEET_ID)

    # 설문지 응답 읽기
    form_ws = wb.worksheet("설문지 응답 시트1")
    form_data = form_ws.get_all_values()

    if len(form_data) <= 1:
        print("설문지 응답이 없습니다.")
        return

    responses = form_data[1:]  # 헤더 제외
    print(f"총 {len(responses)}개 응답 처리 시작\n")

    # 반별 시트 캐시
    sheet_cache: dict[str, list[list]] = {}
    ws_cache: dict[str, gspread.Worksheet] = {}

    total_updated = 0
    total_skipped = 0

    for resp_idx, resp in enumerate(responses, start=2):  # 1-based row (헤더=1)
        ban_raw  = safe_get(resp, FORM_BAN)
        num_raw  = safe_get(resp, FORM_NUM)
        name     = safe_get(resp, FORM_NAME)
        type_raw = safe_get(resp, FORM_TYPE)
        school   = safe_get(resp, FORM_SCHOOL)
        jeon_raw = safe_get(resp, FORM_JEONHYEONG)
        spec_raw = safe_get(resp, FORM_SPECIAL)
        stage    = safe_get(resp, FORM_STAGE)

        # 반 → 시트 이름
        sheet_name = ban_to_sheet_name(ban_raw)
        if not sheet_name:
            print(f"⚠ 행{resp_idx}: 반 값 '{ban_raw}' 인식 불가 → 스킵")
            total_skipped += 1
            continue

        # 번호 파싱
        try:
            num = int(re.search(r"\d+", num_raw).group())
        except (AttributeError, ValueError):
            print(f"⚠ 행{resp_idx}: 번호 '{num_raw}' 인식 불가 → 스킵")
            total_skipped += 1
            continue

        # 시트 로드 (캐시)
        if sheet_name not in sheet_cache:
            try:
                ws = wb.worksheet(sheet_name)
                ws_cache[sheet_name] = ws
                sheet_cache[sheet_name] = ws.get_all_values()
            except gspread.WorksheetNotFound:
                print(f"⚠ 시트 '{sheet_name}' 없음 → 스킵")
                total_skipped += 1
                continue

        data = sheet_cache[sheet_name]
        student_index = build_student_index(data)

        if num not in student_index:
            print(f"⚠ [{sheet_name}] 번호 {num}({name}) 명렬표에 없음 → 스킵")
            total_skipped += 1
            continue

        row_idx = student_index[num]
        updates: dict[int, str] = {}

        # ── 지원 고등학교 유형 + 희망교명 ───────────────────────────────────
        for school_type in parse_checkbox(type_raw):
            col = None
            for keyword, c in SCHOOL_TYPE_MAP.items():
                if keyword in school_type:
                    col = c
                    break
            if col is None:
                print(f"  ⚠ [{sheet_name}] {name}: 유형 '{school_type}' 매핑 없음")
                continue
            # 마이스터/특성화는 희망교명을, 일반고는 'O'를 기본값으로
            value = school if school else "O"
            updates[col] = value

        # ── 마이스터고/특성화고 학과 ─────────────────────────────────────────
        dept = safe_get(resp, FORM_DEPT)
        if dept:
            updates[COL_DEPT] = dept

        # ── 특이사항 세부 내용 ───────────────────────────────────────────────
        special_detail = safe_get(resp, FORM_SPECIAL_DETAIL)
        if special_detail:
            updates[COL_SPECIAL_NOTE] = special_detail

        # ── 전형 구분 ────────────────────────────────────────────────────────
        for jeon in parse_checkbox(jeon_raw):
            for keyword, col in JEONHYEONG_MAP.items():
                if keyword in jeon:
                    updates[col] = "O"
                    break

        # ── 특별 기재 사항 ───────────────────────────────────────────────────
        for spec in parse_checkbox(spec_raw):
            for keyword, col in SPECIAL_NOTE_MAP.items():
                if keyword in spec:
                    updates[col] = "O"
                    break

        # ── 준비 단계 (로그만) ───────────────────────────────────────────────
        if stage:
            print(f"  ℹ [{sheet_name}] {name} 준비단계: '{stage}' (반별 시트에 열 없음 → 무시)")

        if not updates:
            print(f"  - [{sheet_name}] {num}번 {name}: 기록할 데이터 없음")
            continue

        # ── 변경 계산 ────────────────────────────────────────────────────────
        changes = sync_row_to_sheet(data, row_idx, updates, sheet_name, name)

        if not changes:
            print(f"  ✓ [{sheet_name}] {num}번 {name}: 변경 없음 (이미 입력됨)")
            continue

        print(f"  ✏ [{sheet_name}] {num}번 {name}: {len(changes)}개 셀 업데이트")
        for r, c, v in changes:
            col_label = col_idx_to_a1(c)
            print(f"      {col_label}{r+1} ← '{v}'")

        if not dry_run:
            ws = ws_cache[sheet_name]
            # 배치 업데이트
            cell_list = []
            for r, c, v in changes:
                cell_list.append(gspread.Cell(r + 1, c + 1, v))  # 1-based
                # 로컬 캐시도 업데이트
                while len(sheet_cache[sheet_name][r]) <= c:
                    sheet_cache[sheet_name][r].append("")
                sheet_cache[sheet_name][r][c] = v
            ws.update_cells(cell_list)
            total_updated += len(changes)

    print(f"\n{'[DRY RUN] ' if dry_run else ''}완료: {total_updated}개 셀 업데이트, {total_skipped}개 응답 스킵")


if __name__ == "__main__":
    dry = "--dry" in sys.argv or "--dry-run" in sys.argv
    if dry:
        print("=== DRY RUN 모드 (실제 변경 없음) ===\n")
    sync_all(dry_run=dry)
