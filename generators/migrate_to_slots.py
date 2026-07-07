#!/usr/bin/env python3
"""구형 V:X(유형별 접수) + Y:AB(1차/2차/3차/최종) → 시기 슬롯 구조 이관.

기본 dry-run: 이관 계획만 출력한다. --apply 시에만 시트에 쓴다.
근거 스펙: docs/superpowers/specs/2026-07-07-admission-slot-redesign-design.md
"""
from __future__ import annotations

import argparse
import csv
from datetime import datetime
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SS_ID = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"

NEW_CLASS_HEADERS = ["영재고_접수학교", "영재고_결과", "전기_접수학교", "전기_결과", "후기_접수학교", "후기_결과", ""]
OLD_CLASS_HEADERS = ["실제접수_영재고", "실제접수_과학고", "실제접수_자사고", "1차", "2차", "3차", "최종"]
POSITIVE = {"O", "o", "○", "V", "v", "✓", "1", "y", "Y", "예", "있음"}
# 구형 유형별 접수 → (새 슬롯 school col offset(V=0기준), 희망학교 fallback col 0-based)
OLD_RECEIPT_TO_SLOT = {0: (0, 7), 1: (2, 8), 2: (4, 12)}  # 영재고->영재고슬롯/H, 과학고->전기/I, 자사고->후기/M
TRACKING_RENAMES = {
    "접수유형": "영재고_접수",
    "접수학교": "영재고_결과",
    "접수상태": "전기_접수학교",
    "최종유형": "전기_유형",
    "최종학교": "전기_결과",
}


def norm(v) -> str:
    return str(v).strip()


def normalize_result(rounds: list[str]) -> tuple[str, str]:
    """구형 1차/2차/3차/최종 값 → (결과코드, 비고). 마지막 비어있지 않은 라운드 기준."""
    labels = ["1차", "2차", "3차", "최종"]
    code = ""
    note = ""
    for label, raw_value in zip(labels, rounds):
        val = norm(raw_value)
        if not val:
            continue
        positive = val in {"합", "합격", "O", "o", "○", "pass", "PASS"}
        negative = val in {"불", "불합", "불합격", "X", "x", "탈락"}
        if label == "최종":
            code = "최종합" if positive else ("최종불" if negative else val)
        elif label == "3차":
            code = "2차합" if positive else ("2차불" if negative else val)  # 새 코드에 3차 없음
            note = f"구3차={val}"
        else:
            code = f"{label}합" if positive else (f"{label}불" if negative else val)
        if code and code not in {"1차합", "1차불", "2차합", "2차불", "최종합", "최종불", "포기"}:
            note = (note + " " if note else "") + f"비표준값 원본={val}"
    return code, note


def backup(ss, titles: list[str]) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = Path(__file__).resolve().parents[1] / "backups" / f"slots_{stamp}"
    out.mkdir(parents=True)
    for title in titles:
        rows = ss.worksheet(title).get_all_values()
        with open(out / f"{title}.csv", "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rows)
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    ss = gspread.authorize(creds).open_by_key(SS_ID)

    class_titles = [w.title for w in ss.worksheets() if w.title.isdigit() and w.title.startswith("3") and len(w.title) == 3]

    if args.apply:
        path = backup(ss, class_titles + ["입시_트래킹", "입시 진행 현황"])
        print(f"백업 완료: {path}")

    # ── 반별 시트: V:AB 이관 계획 수립 ──
    response = ss.values_batch_get(ranges=[f"'{t}'!A1:AB80" for t in class_titles])
    plans = []  # (sheet, row_num, old V:AB 7칸, new V:AB 7칸)
    for title, vr in zip(class_titles, response["valueRanges"]):
        rows = vr.get("values", [])
        header = (rows[0] + [""] * 28)[21:28] if rows else []
        if [norm(h) for h in header] != OLD_CLASS_HEADERS:
            print(f"⚠ {title}: V:AB 헤더가 구형과 다름 — 건너뜀: {header}")
            continue
        for row_num, row in enumerate(rows[1:], start=2):
            row = row + [""] * (28 - len(row))
            old_block = [norm(c) for c in row[21:28]]
            if not any(old_block):
                continue
            new_block = [""] * 7
            # 접수 3칸: 값이 긍정마커면 희망학교 fallback, 아니면 학교명 그대로
            active_slot = None
            for old_idx, (new_offset, hope_col) in OLD_RECEIPT_TO_SLOT.items():
                val = old_block[old_idx]
                if not val or val in {"X", "x", "0"}:
                    continue
                school = norm(row[hope_col]) if val in POSITIVE else val
                new_block[new_offset] = school
                active_slot = new_offset
            # 결과 4칸 → 접수가 있는 슬롯의 결과 칸으로 (없으면 영재고 슬롯 + 확인 표시)
            code, note = normalize_result(old_block[3:7])
            if code:
                target = (active_slot + 1) if active_slot is not None else 1
                new_block[target] = code
                if active_slot is None:
                    note = (note + " " if note else "") + "접수칸없이결과만존재-확인필요"
            if any(new_block):
                plans.append((title, row_num, old_block, new_block, note))

    print(f"\n=== 반별 시트 이관 계획: {len(plans)}행 ===")
    for title, row_num, old, new, note in plans:
        print(f"  {title}!{row_num}: {old} -> {new}" + (f"  [{note}]" if note else ""))

    print(f"\n=== 반별 시트 헤더 교체: {len(class_titles)}개 시트 V1:AB1 -> {NEW_CLASS_HEADERS} ===")
    print(f"=== 트래킹 헤더 rename + 데이터 클리어: {TRACKING_RENAMES} ===")

    if not args.apply:
        print("\ndry-run 완료. 위 계획을 육안 확인 후 --apply로 실행하세요.")
        return 0

    for title in class_titles:
        ws = ss.worksheet(title)
        ws.update(values=[NEW_CLASS_HEADERS], range_name="V1:AB1")
        row_plans = [(r, new) for t, r, _o, new, _n in plans if t == title]
        if row_plans:
            ws.batch_update([{"range": f"V{r}:AB{r}", "values": [new]} for r, new in row_plans])
        print(f"  ✅ {title} 적용")

    tracking = ss.worksheet("입시_트래킹")
    header = tracking.row_values(1)
    updates = []
    for idx, name in enumerate(header):
        if norm(name) in TRACKING_RENAMES:
            col = gspread.utils.rowcol_to_a1(1, idx + 1)
            updates.append({"range": col, "values": [[TRACKING_RENAMES[norm(name)]]]})
            col_letter = gspread.utils.rowcol_to_a1(1, idx + 1).rstrip("1")
            updates.append({"range": f"{col_letter}2:{col_letter}{tracking.row_count}", "values": [[""]] * (tracking.row_count - 1)})
    if updates:
        tracking.batch_update(updates)
    print("  ✅ 입시_트래킹 헤더 rename + 구 데이터 클리어 완료")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
