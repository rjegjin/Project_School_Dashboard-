#!/usr/bin/env python3
"""
반별/조기졸업예정자 시트에서 희망 지원 데이터를 읽어 입시_트래킹과 비교/동기화한다.

기본 동작은 기존 호환성을 위해 입시_트래킹의 비어 있는 기존 `유형`, `지원학교`
컬럼만 채운다. 새 희망/접수 분리 컬럼은 `--apply-schema`를 명시했을 때만
추가/갱신한다. `최종유형`, `최종학교`는 결과 입력 또는 관리자 확정 전에는
희망값으로 자동 승격하지 않는다.

실행:
  python sync_type_to_tracking.py 2026 --dry-run
  python sync_type_to_tracking.py 2026 --dry-run --apply-schema
  python sync_type_to_tracking.py 2026 --apply-schema
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from dataclasses import dataclass

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

TYPE_MAP = {
    7: "영재고",
    8: "과학고",
    9: "예술계고",
    10: "특성화고",
    12: "자사고",
    13: "외고/국제고",
    14: "일반고",
    15: "기타/대안",
}

RECEIPT_TYPE_MAP = {
    21: "영재고",
    22: "과학고",
    23: "자사고",
}

PRIORITY = ["영재고", "과학고", "예술계고", "특성화고", "자사고", "외고/국제고", "기타/대안", "일반고"]
EMPTY_MARKERS = {"", "X", "x", "0"}
CHECK_MARKERS = {"O", "o", "○"}
RECEIPT_POSITIVE_MARKERS = {"O", "o", "○", "V", "v", "✓", "1", "y", "Y", "예", "있음"}
RECEIPT_NEGATIVE_MARKERS = {"X", "x", "0", "아니오", "없음", "철회"}

SCHEMA_COLUMNS = [
    "학년",
    "학생ID",
    "희망유형",
    "희망학교",
    "접수유형",
    "접수학교",
    "접수상태",
    "최종유형",
    "최종학교",
    "데이터상태",
    "조기졸업여부",
    "출처",
]


@dataclass(frozen=True)
class SourceStudent:
    grade: str
    cls: str
    num: str
    name: str
    gender: str
    hope_type: str
    hope_school: str
    source: str
    early_grad: bool = False
    receipt_type: str = ""
    receipt_school: str = ""
    receipt_status: str = ""
    receipt_note: str = ""

    def student_id(self, year: str) -> str:
        return f"{year}-{self.grade}-{self.cls}-{self.num}"

    def identity(self) -> tuple[str, str, str]:
        return self.cls, self.num, self.name


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("year", nargs="?", default="2026")
    parser.add_argument("--dry-run", action="store_true", help="시트에 쓰지 않고 변경 후보만 출력")
    parser.add_argument(
        "--apply-schema",
        action="store_true",
        help="희망/접수/최종 분리 컬럼을 입시_트래킹에 추가하고 갱신",
    )
    parser.add_argument(
        "--no-legacy-fill",
        action="store_true",
        help="기존 유형/지원학교 빈칸 채우기도 하지 않음",
    )
    parser.add_argument(
        "--include-early-graduates",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="조기졸업예정자 시트를 소스에 포함",
    )
    return parser.parse_args()


def col_letter(idx: int) -> str:
    """0-based column index를 A1 column letter로 변환한다."""
    idx += 1
    letters = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def norm(value: object) -> str:
    return str(value).strip()


def detect_type(row: list[str]) -> str:
    row = row + [""] * (40 - len(row))
    for priority_type in PRIORITY:
        for col_idx, type_name in TYPE_MAP.items():
            if type_name != priority_type:
                continue
            val = norm(row[col_idx])
            if val and val not in EMPTY_MARKERS:
                return type_name
    return ""


def detect_school(row: list[str], school_type: str) -> str:
    if not school_type:
        return ""
    for col_idx, type_name in TYPE_MAP.items():
        if type_name != school_type or col_idx >= len(row):
            continue
        val = norm(row[col_idx])
        return "" if val in EMPTY_MARKERS or val in CHECK_MARKERS else val
    return ""


def detect_receipt(row: list[str], hope_type: str, hope_school: str) -> tuple[str, str, str, str]:
    """반별 시트 V:X 실제 접수 블록에서 접수유형/학교/상태를 판정한다."""
    receipts: list[tuple[str, str, str]] = []
    negatives: list[str] = []

    for col_idx, type_name in RECEIPT_TYPE_MAP.items():
        val = norm(row[col_idx]) if col_idx < len(row) else ""
        if not val:
            continue

        if val in RECEIPT_NEGATIVE_MARKERS:
            negatives.append(type_name)
            continue

        if val in RECEIPT_POSITIVE_MARKERS:
            school = detect_school(row, type_name)
            if not school and hope_type == type_name:
                school = hope_school
            receipts.append((type_name, school, val))
            continue

        receipts.append((type_name, val, val))

    if len(receipts) > 1:
        chosen_type, chosen_school, _raw = sorted(
            receipts,
            key=lambda item: list(RECEIPT_TYPE_MAP.values()).index(item[0]),
        )[0]
        note = ", ".join(f"{rtype}:{school or raw}" for rtype, school, raw in receipts)
        return chosen_type, chosen_school, "복수접수 확인필요", note

    if len(receipts) == 1:
        chosen_type, chosen_school, _raw = receipts[0]
        return chosen_type, chosen_school, "접수확정", ""

    if negatives:
        return "", "", "접수안함/철회", ", ".join(negatives)

    return "", "", "", ""


def load_class_sources(ss) -> list[SourceStudent]:
    students: list[SourceStudent] = []
    for sht in ss.worksheets():
        title = sht.title.strip()
        if not (title.isdigit() and len(title) == 3 and title.startswith("3")):
            continue

        rows = sht.get_all_values()
        data_start = 2 if len(rows) > 2 and len(rows[1]) > 2 and not norm(rows[1][2]) else 1
        for row_idx, row in enumerate(rows[data_start:], start=data_start + 1):
            if len(row) < 3 or not norm(row[2]):
                continue

            hope_type = detect_type(row)
            hope_school = detect_school(row, hope_type)
            receipt_type, receipt_school, receipt_status, receipt_note = detect_receipt(row, hope_type, hope_school)
            students.append(
                SourceStudent(
                    grade="3",
                    cls=norm(row[0]),
                    num=norm(row[1]),
                    name=norm(row[2]),
                    gender=norm(row[3]) if len(row) > 3 else "",
                    hope_type=hope_type,
                    hope_school=hope_school,
                    source=f"{title}!{row_idx}",
                    receipt_type=receipt_type,
                    receipt_school=receipt_school,
                    receipt_status=receipt_status,
                    receipt_note=receipt_note,
                )
            )
    return students


def load_early_graduate_sources(ss) -> list[SourceStudent]:
    try:
        sht = ss.worksheet("조기졸업예정자")
    except gspread.exceptions.WorksheetNotFound:
        return []

    students: list[SourceStudent] = []
    rows = sht.get_all_values()
    for row_idx, row in enumerate(rows[2:], start=3):
        if len(row) < 4 or not norm(row[3]):
            continue

        hope_school = norm(row[8]) if len(row) > 8 else ""
        gifted_mark = norm(row[14]) if len(row) > 14 else ""
        hope_type = "영재고" if hope_school or gifted_mark in CHECK_MARKERS else ""
        students.append(
            SourceStudent(
                grade=norm(row[0]) or "2",
                cls=norm(row[1]),
                num=norm(row[2]),
                name=norm(row[3]),
                gender=norm(row[4]) if len(row) > 4 else "",
                hope_type=hope_type,
                hope_school=hope_school,
                source=f"조기졸업예정자!{row_idx}",
                early_grad=True,
            )
        )
    return students


def nonempty_header_len(header: list[str]) -> int:
    last = 0
    for idx, value in enumerate(header):
        if norm(value):
            last = idx + 1
    return last


def safe_append_start_col(worksheet, rows: list[list[str]]) -> int:
    """헤더가 비어 있어도 기존 행 데이터/시트 폭을 침범하지 않는 append 시작 컬럼."""
    used_by_rows = max((len(row) for row in rows), default=0)
    return max(nonempty_header_len(rows[0] if rows else []), used_by_rows, worksheet.col_count)


def col_map(header: list[str]) -> dict[str, int]:
    return {norm(name): idx for idx, name in enumerate(header) if norm(name)}


def get_cell(row: list[str], idx: int | None) -> str:
    return norm(row[idx]) if idx is not None and idx < len(row) else ""


def set_if_changed(updates: list[tuple[str, str, str, str]], row_num: int, col_idx: int, old: str, new: str, reason: str) -> None:
    if old != new:
        updates.append((f"{col_letter(col_idx)}{row_num}", old, new, reason))


def build_source_indexes(sources: list[SourceStudent]) -> tuple[dict[tuple[str, str, str], SourceStudent], dict[tuple[str, str], list[SourceStudent]]]:
    by_identity: dict[tuple[str, str, str], SourceStudent] = {}
    by_key: dict[tuple[str, str], list[SourceStudent]] = defaultdict(list)
    for student in sources:
        by_identity[student.identity()] = student
        by_key[(student.cls, student.num)].append(student)
    return by_identity, by_key


def choose_source(
    row: list[str],
    columns: dict[str, int],
    by_identity: dict[tuple[str, str, str], SourceStudent],
    by_key: dict[tuple[str, str], list[SourceStudent]],
) -> tuple[SourceStudent | None, str]:
    c_cls = columns.get("반", 0)
    c_num = columns.get("번호", 1)
    c_name = columns.get("성명", columns.get("이름", 2))
    c_grade = columns.get("학년")

    cls = get_cell(row, c_cls)
    num = get_cell(row, c_num)
    name = get_cell(row, c_name)
    grade = get_cell(row, c_grade)

    if not cls or not num or not name:
        return None, "학생 기본값 없음"

    if grade:
        candidates = [s for s in by_key.get((cls, num), []) if s.grade == grade and s.name == name]
        if len(candidates) == 1:
            return candidates[0], ""

    identity = (cls, num, name)
    if identity in by_identity:
        return by_identity[identity], ""

    candidates = by_key.get((cls, num), [])
    if not candidates:
        return None, "소스 없음"

    names = ", ".join(f"{s.name}/{s.source}" for s in candidates)
    return None, f"반번호 충돌 또는 이름 불일치: {names}"


def data_status(hope_type: str, receipt_type: str, receipt_status: str, final_type: str, legacy_type: str) -> str:
    if receipt_status == "복수접수 확인필요":
        return "복수접수 확인필요"
    if receipt_type:
        return "접수확정" if not hope_type or receipt_type == hope_type else "희망/접수 불일치"
    if legacy_type and hope_type and legacy_type != hope_type:
        return "기존값/희망 불일치"
    if final_type:
        return "최종확정"
    if hope_type:
        return "희망만"
    return "미입력"


def print_updates(title: str, updates: list[tuple[str, str, str, str]], limit: int = 80) -> None:
    print(f"\n[{title}] {len(updates)}건")
    for cell, old, new, reason in updates[:limit]:
        old_text = old if old else "(blank)"
        new_text = new if new else "(blank)"
        print(f"  - {cell}: {old_text!r} -> {new_text!r} ({reason})")
    if len(updates) > limit:
        print(f"  ... {len(updates) - limit}건 더 있음")


def apply_updates(worksheet, updates: list[tuple[str, str, str, str]]) -> None:
    for chunk_start in range(0, len(updates), 100):
        chunk = updates[chunk_start : chunk_start + 100]
        worksheet.batch_update([{"range": cell, "values": [[new]]} for cell, _old, new, _reason in chunk])


def ensure_column_capacity(worksheet, required_cols: int) -> None:
    if worksheet.col_count < required_cols:
        worksheet.add_cols(required_cols - worksheet.col_count)


def main() -> int:
    args = parse_args()
    ss_id = SPREADSHEET_IDS.get(args.year, args.year)

    print("=" * 70)
    print(f" {args.year}학년도 입시_트래킹 희망/접수 분리 동기화")
    print(f" dry_run={args.dry_run}, apply_schema={args.apply_schema}, legacy_fill={not args.no_legacy_fill}")
    print("=" * 70)

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    sources = load_class_sources(ss)
    if args.include_early_graduates:
        sources.extend(load_early_graduate_sources(ss))

    by_identity, by_key = build_source_indexes(sources)
    type_counts = Counter(s.hope_type for s in sources if s.hope_type)
    early_count = sum(1 for s in sources if s.early_grad)
    print(f"\n[1/3] 소스 로드: {len(sources)}명, 조기졸업예정자 {early_count}명")
    print(f"  희망유형 통계: {dict(type_counts)}")

    tracking = ss.worksheet("입시_트래킹")
    rows = tracking.get_all_values()
    if not rows:
        print("입시_트래킹 시트가 비어 있습니다.")
        return 1

    header = rows[0]
    columns = col_map(header)
    header_updates: list[tuple[str, str, str, str]] = []
    if args.apply_schema:
        next_col = safe_append_start_col(tracking, rows)
        for name in SCHEMA_COLUMNS:
            if name in columns:
                continue
            columns[name] = next_col
            header_updates.append((f"{col_letter(next_col)}1", "", name, "schema column"))
            next_col += 1

    legacy_updates: list[tuple[str, str, str, str]] = []
    schema_updates: list[tuple[str, str, str, str]] = []
    conflicts: list[str] = []
    matched_sources: set[tuple[str, str, str, str]] = set()
    matched = 0

    c_cls = columns.get("반", 0)
    c_num = columns.get("번호", 1)
    c_name = columns.get("성명", columns.get("이름", 2))
    c_type = columns.get("유형")
    c_school = columns.get("지원학교")
    c_receipt_type = columns.get("접수유형")
    c_receipt_school = columns.get("접수학교")
    c_receipt_status = columns.get("접수상태")
    c_final_type = columns.get("최종유형")
    c_final_school = columns.get("최종학교")

    for row_num, row in enumerate(rows[1:], start=2):
        if not get_cell(row, c_cls) or not get_cell(row, c_num) or not get_cell(row, c_name):
            continue

        source, reason = choose_source(row, columns, by_identity, by_key)
        if source is None:
            if reason and reason != "소스 없음":
                conflicts.append(f"row {row_num}: {reason}")
            continue

        matched += 1
        matched_sources.add((source.grade, source.cls, source.num, source.name))
        legacy_type = get_cell(row, c_type)
        legacy_school = get_cell(row, c_school)
        receipt_type_existing = get_cell(row, c_receipt_type)
        receipt_school_existing = get_cell(row, c_receipt_school)
        receipt_status_existing = get_cell(row, c_receipt_status)
        final_type_existing = get_cell(row, c_final_type)
        final_school_existing = get_cell(row, c_final_school)

        if c_type is not None and source.hope_type and legacy_type and legacy_type != source.hope_type:
            conflicts.append(
                f"row {row_num}: 기존 유형={legacy_type!r}, 희망유형={source.hope_type!r}, source={source.source}"
            )
        if c_school is not None and source.hope_school and legacy_school and legacy_school != source.hope_school:
            conflicts.append(
                f"row {row_num}: 기존 지원학교={legacy_school!r}, 희망학교={source.hope_school!r}, source={source.source}"
            )
        if source.receipt_status == "복수접수 확인필요":
            conflicts.append(
                f"row {row_num}: 복수 실제접수={source.receipt_note!r}, 대표접수유형={source.receipt_type!r}, source={source.source}"
            )
        if source.receipt_type and source.hope_type and source.receipt_type != source.hope_type:
            conflicts.append(
                f"row {row_num}: 희망유형={source.hope_type!r}, 접수유형={source.receipt_type!r}, source={source.source}"
            )

        if not args.no_legacy_fill:
            if c_type is not None and source.hope_type and not legacy_type:
                set_if_changed(legacy_updates, row_num, c_type, legacy_type, source.hope_type, f"legacy blank fill from {source.source}")

            if c_school is not None and source.hope_school and not legacy_school:
                set_if_changed(legacy_updates, row_num, c_school, legacy_school, source.hope_school, f"legacy blank fill from {source.source}")

        if args.apply_schema:
            schema_values = {
                "학년": source.grade,
                "학생ID": source.student_id(args.year),
                "희망유형": source.hope_type,
                "희망학교": source.hope_school,
                "접수유형": source.receipt_type,
                "접수학교": source.receipt_school,
                "접수상태": source.receipt_status,
                "최종유형": final_type_existing,
                "최종학교": final_school_existing,
                "데이터상태": data_status(
                    source.hope_type,
                    source.receipt_type or receipt_type_existing,
                    source.receipt_status or receipt_status_existing,
                    final_type_existing,
                    legacy_type,
                ),
                "조기졸업여부": "O" if source.early_grad else "",
                "출처": source.source,
            }
            for name, new_value in schema_values.items():
                col_idx = columns.get(name)
                if col_idx is None:
                    continue
                old_value = get_cell(row, col_idx)
                set_if_changed(schema_updates, row_num, col_idx, old_value, new_value, f"schema sync from {source.source}")

    print(f"\n[2/3] 입시_트래킹 매칭: {matched}행")
    print_updates("헤더 추가 후보", header_updates)
    print_updates("기존 유형/지원학교 빈칸 채우기 후보", legacy_updates)
    print_updates("희망/접수/최종 schema 갱신 후보", schema_updates)

    print(f"\n[충돌/확인 필요] {len(conflicts)}건")
    for item in conflicts[:120]:
        print(f"  - {item}")
    if len(conflicts) > 120:
        print(f"  ... {len(conflicts) - 120}건 더 있음")

    unmatched_sources = [
        s for s in sources
        if (s.grade, s.cls, s.num, s.name) not in matched_sources and (s.hope_type or s.hope_school)
    ]
    print(f"\n[소스에는 있으나 입시_트래킹 매칭 없음] {len(unmatched_sources)}건")
    for student in unmatched_sources[:80]:
        print(
            f"  - {student.source}: 학년={student.grade}, 반={student.cls}, 번호={student.num}, "
            f"성명={student.name}, 희망유형={student.hope_type!r}, 희망학교={student.hope_school!r}"
        )
    if len(unmatched_sources) > 80:
        print(f"  ... {len(unmatched_sources) - 80}건 더 있음")

    if args.dry_run:
        print("\n[3/3] dry-run 완료: 시트에 쓰지 않았습니다.")
        return 0

    if args.apply_schema:
        ensure_column_capacity(tracking, safe_append_start_col(tracking, rows) + len(header_updates))

    if header_updates:
        apply_updates(tracking, header_updates)
    if legacy_updates:
        apply_updates(tracking, legacy_updates)
    if schema_updates:
        apply_updates(tracking, schema_updates)

    print("\n[3/3] 업데이트 완료")
    print(f"  헤더 추가: {len(header_updates)}건")
    print(f"  기존 컬럼 업데이트: {len(legacy_updates)}건")
    print(f"  schema 컬럼 업데이트: {len(schema_updates)}건")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
