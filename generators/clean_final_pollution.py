#!/usr/bin/env python3
"""
입시_트래킹의 `최종유형`/`최종학교` 정렬오류 스냅샷을 1회성으로 제거한다.

배경 (2026-06-30 진단):
  `최종유형`/`최종학교`(col 34/35) 값은 같은 행 학생의 값이 아니라
  희망 데이터가 약 2행 밀려 들어간 stale 스냅샷이다.
    - 검증: 최종(n) == 희망(n+2) 가 385건 중 268건(70%) 일치
            (offset 0/1/3은 23/26/14%로 유의하게 낮음)
    - 맨 위 손정인/황준서 행이 출처 중복(301!3, 301!4)인 stale 샘플행이라
      그만큼 최종 컬럼 전체가 밀렸다.
  진짜 희망 데이터는 `희망유형`/`희망학교`(col 29/30) 및 legacy `유형`/`지원학교`
  (col 5/6)에 정상 정렬되어 보존돼 있다. (희망+legacy 모두 빈 학생은 28명뿐)

  결과 발표 전(2026-06-30)이므로 실제 '최종' 결과는 존재할 수 없다.

조치 (pure clear, migrate 금지):
  최종유형/최종학교가 채워진 행의 해당 두 셀만 빈칸으로 클리어한다.
  ※ 최종값을 희망으로 복사하지 않는다 — 밀린 값이라 남의 학생 데이터다.
  ※ '접수확정' 또는 실제 결과가 있는 행은 보존하고 수동검토 대상으로 보고한다.
  실제 변경 전 클리어 대상 전건을 CSV로 backups/ 에 백업한다.

실행:
  python generators/clean_final_pollution.py 2026            # dry-run (기본, 변경 없음)
  python generators/clean_final_pollution.py 2026 --apply    # 백업 후 실제 클리어
"""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from datetime import datetime
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_IDS = {
    "2025": "1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI",
    "2026": "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0",
}
BACKUP_DIR = Path(__file__).resolve().parent.parent / "backups"

# 결과 블록에서 '실제 결과 있음'으로 인정하는 값 (소문자 비교)
RESULT_TOKENS = {
    "합격", "불합격", "pass", "fail", "o", "yes", "v", "x",
    "대기", "추가합격", "추합", "예비", "1차합격", "2차합격", "응시", "포기",
}


def norm(value: object) -> str:
    return str(value).strip()


def col_letter(idx: int) -> str:
    idx += 1
    letters = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("year", nargs="?", default="2026")
    p.add_argument("--apply", action="store_true", help="백업 후 실제로 최종 컬럼을 클리어한다 (미지정 시 dry-run)")
    return p.parse_args()


def has_real_result(row: list[str], cols: dict[str, int]) -> bool:
    """결과 블록(1차/2차/최종)에 의미있는 결과값이 하나라도 있으면 True."""
    for name in ("1차", "2차", "최종"):
        idx = cols.get(name)
        if idx is None or idx >= len(row):
            continue
        if norm(row[idx]):  # 사람이 무언가 기록 → 보수적으로 결과 있음
            return True
    return False


def main() -> int:
    args = parse_args()
    ss_id = SPREADSHEET_IDS.get(args.year, args.year)

    print("=" * 70)
    print(f" {args.year}학년도 입시_트래킹 최종유형/최종학교 정렬오류 제거")
    print(f" mode = {'APPLY (실제 클리어)' if args.apply else 'DRY-RUN (변경 없음)'}")
    print("=" * 70)

    gc = gspread.authorize(Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES))
    ws = gc.open_by_key(ss_id).worksheet("입시_트래킹")
    rows = ws.get_all_values()
    if not rows:
        print("입시_트래킹 시트가 비어 있습니다.")
        return 1

    header = rows[0]
    cols = {norm(h): i for i, h in enumerate(header) if norm(h)}
    for required in ("최종유형", "최종학교", "접수상태"):
        if required not in cols:
            print(f"❌ 필수 컬럼 '{required}' 을(를) 찾을 수 없습니다. 헤더: {header}")
            return 1

    c_final_type = cols["최종유형"]
    c_final_school = cols["최종학교"]
    c_receipt_status = cols["접수상태"]
    c_cls = cols.get("반", 1)
    c_num = cols.get("번호", 2)
    c_name = cols.get("성명", 3)

    to_clear: list[dict] = []
    keep_receipt: list[int] = []
    keep_result: list[int] = []

    for row_num, row in enumerate(rows[1:], start=2):
        ft = norm(row[c_final_type]) if c_final_type < len(row) else ""
        fs = norm(row[c_final_school]) if c_final_school < len(row) else ""
        if not ft and not fs:
            continue

        rs = norm(row[c_receipt_status]) if c_receipt_status < len(row) else ""
        if rs == "접수확정":
            keep_receipt.append(row_num)
            continue
        if has_real_result(row, cols):
            keep_result.append(row_num)
            continue

        to_clear.append({
            "row": row_num,
            "반": norm(row[c_cls]) if c_cls < len(row) else "",
            "번호": norm(row[c_num]) if c_num < len(row) else "",
            "성명": norm(row[c_name]) if c_name < len(row) else "",
            "최종유형": ft,
            "최종학교": fs,
        })

    total = len(to_clear) + len(keep_receipt) + len(keep_result)
    print(f"\n최종유형/최종학교가 채워진 행: {total}건")
    print(f"  ├─ 보존(접수확정, 수동검토): {len(keep_receipt)}건  rows={keep_receipt}")
    print(f"  ├─ 보존(실제 결과 있음):     {len(keep_result)}건  rows={keep_result}")
    print(f"  └─ 클리어 대상(정렬오류):     {len(to_clear)}건")
    print("  ※ 최종값은 희망으로 복사하지 않습니다 (밀린 값=남의 데이터). 진짜 희망은 희망/legacy 컬럼에 보존됨.")

    if to_clear:
        dist = Counter(p["최종유형"] for p in to_clear if p["최종유형"])
        print(f"\n클리어 대상 최종유형 분포(밀린 값 기준): {dict(dist)}")
        print("\n클리어 대상 (최대 30건):")
        for p in to_clear[:30]:
            print(f"  row {p['row']:>3}: {p['반']}반 {p['번호']}번 {p['성명']:<6} "
                  f"최종유형={p['최종유형']!r} 최종학교={p['최종학교']!r} → (빈칸)")
        if len(to_clear) > 30:
            print(f"  ... {len(to_clear) - 30}건 더 있음")

    if not args.apply:
        print("\n[DRY-RUN] 시트에 쓰지 않았습니다. 실제 클리어하려면 --apply 를 붙이세요.")
        return 0

    if not to_clear:
        print("\n클리어할 행이 없습니다.")
        return 0

    # ── 백업 ──
    BACKUP_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"final_misalign_{args.year}_{stamp}.csv"
    with backup_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["row", "반", "번호", "성명", "최종유형", "최종학교"])
        w.writeheader()
        w.writerows(to_clear)
    print(f"\n백업 저장: {backup_path}")

    # ── 클리어 (최종유형/최종학교 두 셀만 빈칸) ──
    requests = []
    for p in to_clear:
        requests.append({"range": f"{col_letter(c_final_type)}{p['row']}", "values": [[""]]})
        requests.append({"range": f"{col_letter(c_final_school)}{p['row']}", "values": [[""]]})
    for i in range(0, len(requests), 100):
        ws.batch_update(requests[i:i + 100])

    print(f"클리어 완료: {len(to_clear)}행 × 2컬럼 = {len(requests)}셀")
    print("→ 다음 auto_sync 실행 시 데이터상태가 '희망만' 등으로 재계산됩니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
