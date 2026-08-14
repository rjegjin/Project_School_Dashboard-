#!/usr/bin/env python3
"""입시_트래킹을 단일 마스터 표로 재생성한다.

기존 시트는 A:L(레거시)와 AB:AQ(신 스키마)가 M:AA 공백을 사이에 두고 공존했고,
두 블록이 서로 다른 학생을 가리키는 행이 3건 있었다(학생ID/출처 중복). 어느 쪽이
Master인지 모호한 구조 자체가 원인이므로, 진짜 원본인 반별 시트에서 한 벌로
다시 만든다.

  학생 1명 = 1행, 열 = 기본정보 → 희망 → 시기별 실제지원 → 메타

ponytail: 3행을 손으로 고치지 않는다. 오염된 파생 테이블을 원본에서 재생성하면
같은 오염이 다시 생겨도 이 스크립트 한 번으로 복구된다.

실행:
  python consolidate_tracking.py 2026            # dry-run (기본)
  python consolidate_tracking.py 2026 --apply
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

from school_types import classify_school, normalize_school
from sync_type_to_tracking import (
    KEY_FILE,
    SCOPES,
    SPREADSHEET_IDS,
    load_class_sources,
    load_early_graduate_sources,
    norm,
)

TRACKING = "입시_트래킹"

HEADER = [
    "학년", "반", "번호", "성명", "성별", "학생ID",
    "희망유형", "희망학교", "희망학교_표준명",
    "영재고_접수", "영재고_결과",
    "전기_접수학교", "전기_유형", "전기_결과",
    "후기_접수학교", "후기_유형", "후기_결과",
    "최종배정학교",
    "조기졸업여부", "데이터상태", "출처", "비고",
]

# 진학부 전용 — 동기화가 절대 덮어쓰지 않고 기존 값을 그대로 이월한다
ADMIN_ONLY = ("최종배정학교", "비고")

BACKUP_DIR = os.path.join(os.path.dirname(__file__), "..", "backups")


def data_state(slots: dict, final_school: str) -> str:
    if final_school:
        return "최종확정"
    if any(s.get("result") for s in slots.values()):
        return "결과입력"
    if any(s.get("school") for s in slots.values()):
        return "접수완료"
    return "희망만"


def build_row(student, year: str, carried: dict[str, str]) -> list[str]:
    slots = student.slots
    gifted = slots.get("영재고", {})
    early = slots.get("전기", {})
    late = slots.get("후기", {})
    final_school = carried.get("최종배정학교", "")
    return [
        student.grade,
        student.cls,
        student.num,
        student.name,
        student.gender,
        student.student_id(year),
        student.hope_type,
        student.hope_school,
        normalize_school(student.hope_school),
        gifted.get("school", ""),
        gifted.get("result", ""),
        early.get("school", ""),
        classify_school(early.get("school", "")),
        early.get("result", ""),
        late.get("school", ""),
        classify_school(late.get("school", "")),
        late.get("result", ""),
        final_school,
        "O" if student.early_grad else "",
        data_state(slots, final_school),
        student.source,
        carried.get("비고", ""),
    ]


def read_existing(ws) -> tuple[list[list[str]], dict[str, dict[str, str]]]:
    """기존 시트 → (원본 rows, 학생ID별 진학부 전용 값)."""
    rows = ws.get_all_values()
    if not rows:
        return [], {}
    idx = {name: i for i, name in enumerate(rows[0])}
    carried: dict[str, dict[str, str]] = {}
    id_col = idx.get("학생ID")
    if id_col is None:
        return rows, {}
    for row in rows[1:]:
        sid = norm(row[id_col]) if id_col < len(row) else ""
        if not sid:
            continue
        keep = {}
        for field in ADMIN_ONLY:
            i = idx.get(field)
            if i is not None and i < len(row) and norm(row[i]):
                keep[field] = norm(row[i])
        if keep:
            carried[sid] = keep
    return rows, carried


def orphan_rows(existing: list[list[str]], known: set[tuple[str, str]], year: str) -> list[list[str]]:
    """반별 시트에 원본이 없는 행. 전출 가능성이 있어 지우지 않고 플래그만 단다."""
    if not existing:
        return []
    idx = {name: i for i, name in enumerate(existing[0])}

    def cell(row, name, default=""):
        i = idx.get(name)
        return norm(row[i]) if i is not None and i < len(row) else default

    out = []
    for row in existing[1:]:
        name = cell(row, "성명")
        if not name:
            continue
        cls = cell(row, "반").lstrip("0") or cell(row, "반")
        num = cell(row, "번호")
        # 번호가 유실된 행이 있으므로 학생ID가 아니라 (반, 성명)으로 대조한다
        if (cls, name) in known or not cls:
            continue
        sid = f"{year}-{cell(row, '학년') or '3'}-{cls}-{num}"
        blank = [""] * len(HEADER)
        blank[0] = cell(row, "학년") or "3"
        blank[1], blank[2], blank[3], blank[4] = cls, num, name, cell(row, "성별")
        blank[5] = sid
        # 오염된 AB:AQ가 아니라 레거시 A:L(이 행에서는 정확했다)을 살린다
        blank[6] = cell(row, "유형")
        blank[7] = cell(row, "지원학교")
        blank[8] = normalize_school(blank[7])
        blank[19] = "확인필요-반별시트없음"
        blank[20] = "(원본없음)"
        out.append(blank)
    return out


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("year", nargs="?", default="2026")
    p.add_argument("--apply", action="store_true", help="실제로 시트를 덮어쓴다")
    args = p.parse_args()

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    ss = gspread.authorize(creds).open_by_key(SPREADSHEET_IDS[args.year])
    ws = ss.worksheet(TRACKING)

    existing, carried = read_existing(ws)

    sources = load_class_sources(ss) + load_early_graduate_sources(ss)
    # 반/번호 없는 행은 학생이 아니라 시트 잔재다 (예: 309!36 이름만 남은 행)
    strays = [s for s in sources if not s.cls or not s.num]
    sources = [s for s in sources if s.cls and s.num]
    rows = [build_row(s, args.year, carried.get(s.student_id(args.year), {})) for s in sources]

    ids = [r[5] for r in rows]
    dupes = {sid for sid in ids if ids.count(sid) > 1}
    orphans = orphan_rows(existing, {(r[1], r[3]) for r in rows}, args.year)

    print(f"기존 {max(len(existing) - 1, 0)}행 → 신규 {len(rows)}행 + 미확인 {len(orphans)}행")
    print(f"진학부 전용 이월(최종배정학교/비고): {len(carried)}건")
    if strays:
        print("⚠ 반/번호 없는 시트 잔재(마스터에서 제외, 반별 시트는 손대지 않음):")
        for s in strays:
            print(f"    {s.source} {s.name!r}")
    if dupes:
        print(f"⚠ 학생ID 중복 {len(set(dupes))}건: {sorted(set(dupes))}")
    if orphans:
        print("⚠ 반별 시트에 원본이 없어 '확인필요'로 남기는 행:")
        for o in orphans:
            print(f"    {o[1]}반 {o[2]}번 {o[3]}")
    unstd = sorted({r[7] for r in rows if r[7] and not r[8]})
    if unstd:
        print(f"표준화 불가 학교명 {len(unstd)}종(복수기재/메모): {unstd[:8]}")

    payload = [HEADER] + rows + orphans
    if not args.apply:
        print("\n[dry-run] --apply 를 붙이면 실제로 씁니다.")
        print("헤더:", HEADER)
        for r in payload[1:4]:
            print("  ", r)
        return

    os.makedirs(BACKUP_DIR, exist_ok=True)
    path = os.path.join(BACKUP_DIR, f"tracking_{datetime.now():%Y%m%d_%H%M%S}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False)
    print(f"백업 저장: {path}")

    ws.clear()
    ws.update(values=payload, range_name="A1")
    print(f"✅ {TRACKING} 재생성 완료 ({len(payload) - 1}행 × {len(HEADER)}열)")


if __name__ == "__main__":
    main()
