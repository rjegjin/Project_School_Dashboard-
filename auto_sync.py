#!/usr/bin/env python3
"""
고입 진학현황 자동 동기화 & 알림 시스템

cron (매일 07:30) 또는 수동 실행:
  /home/rjegj/projects/unified_venv/bin/python auto_sync.py [--force] [--dry-run]

단계:
  1. 설문지 → 반별 시트 동기화 (sync_form_to_class_sheets)
  2. 반별 → 입시_트래킹 유형 동기화 (sync_type_to_tracking)
  3. 트래킹 → 진행현황 시트 갱신 (build_progress_from_tracking)
  4. 스냅샷 비교 → 변경 감지
  5. Telegram 알림 (변경 있을 때 또는 --force)
"""

import sys
import os
import json
import subprocess
import hashlib
import requests
from datetime import datetime
from pathlib import Path

# ── 경로 설정 ────────────────────────────────────────────────────────────────
PROJECT_DIR = Path("/home/rjegj/projects/Project_HighSchool_apply_Dashboard")
GENERATORS_DIR = PROJECT_DIR / "generators"
SNAPSHOT_FILE = PROJECT_DIR / ".auto_sync_snapshot.json"
LOG_FILE = PROJECT_DIR / "logs" / "auto_sync.log"
VENV_PYTHON = "/home/rjegj/projects/unified_venv/bin/python"

# ── Telegram ─────────────────────────────────────────────────────────────────
TELEGRAM_BOT_TOKEN = "8636989262:AAGOVaEZaTZhRb3TrerxdQYBh9kvkiV4Rgc"
TELEGRAM_CHAT_ID = "5929322817"

# ── Google Sheets ─────────────────────────────────────────────────────────────
SERVICE_KEY = "/home/rjegj/projects/.secrets/service_key.json"
SPREADSHEET_ID_2026 = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"


def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def send_telegram(text: str):
    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
            json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"},
            timeout=10,
        )
        if not resp.ok:
            log(f"  Telegram 실패: {resp.text}")
    except Exception as e:
        log(f"  Telegram 오류: {e}")


def run_script(script_path: Path, args: list[str] = None, dry_run: bool = False) -> tuple[bool, str]:
    """스크립트 실행 → (성공 여부, 출력)"""
    cmd = [VENV_PYTHON, str(script_path)] + (args or [])
    if dry_run:
        log(f"  [DRY-RUN] {' '.join(cmd)}")
        return True, "[dry-run]"
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            cwd=str(GENERATORS_DIR),
        )
        output = result.stdout + result.stderr
        if result.returncode != 0:
            log(f"  FAIL (rc={result.returncode}): {output[-500:]}")
            return False, output
        return True, output
    except subprocess.TimeoutExpired:
        return False, "Timeout (120s)"
    except Exception as e:
        return False, str(e)


def fetch_tracking_snapshot() -> dict:
    """입시_트래킹 시트 현황 스냅샷 (학생 수, 유형별 집계)"""
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_file(
            SERVICE_KEY,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(SPREADSHEET_ID_2026)

        ws = ss.worksheet("입시_트래킹")
        rows = ws.get_all_values()
        if not rows:
            return {}

        header = rows[0]
        data_rows = [r for r in rows[1:] if any(c.strip() for c in r)]

        # 유형 컬럼 찾기
        type_col = next((i for i, h in enumerate(header) if "유형" in h or "type" in h.lower()), None)
        school_col = next((i for i, h in enumerate(header) if "배정" in h or "합격" in h), None)

        type_counts: dict[str, int] = {}
        decided = 0
        for row in data_rows:
            if type_col is not None and type_col < len(row):
                t = row[type_col].strip()
                if t:
                    type_counts[t] = type_counts.get(t, 0) + 1
            if school_col is not None and school_col < len(row):
                if row[school_col].strip():
                    decided += 1

        return {
            "total": len(data_rows),
            "decided": decided,
            "types": type_counts,
            "hash": hashlib.md5(str(rows).encode()).hexdigest(),
            "timestamp": datetime.now().isoformat(),
        }
    except Exception as e:
        log(f"  스냅샷 조회 실패: {e}")
        return {}


def load_snapshot() -> dict:
    if SNAPSHOT_FILE.exists():
        try:
            return json.loads(SNAPSHOT_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_snapshot(snap: dict):
    SNAPSHOT_FILE.write_text(json.dumps(snap, ensure_ascii=False, indent=2), encoding="utf-8")


def detect_changes(old: dict, new: dict) -> list[str]:
    """이전 스냅샷 대비 변경 사항 요약"""
    changes = []
    if not old:
        changes.append("첫 번째 스냅샷 기록")
        return changes

    if old.get("hash") == new.get("hash"):
        return []  # 변경 없음

    old_total = old.get("total", 0)
    new_total = new.get("total", 0)
    if new_total != old_total:
        changes.append(f"전체 학생 수: {old_total} → {new_total} ({new_total - old_total:+d}명)")

    old_decided = old.get("decided", 0)
    new_decided = new.get("decided", 0)
    if new_decided != old_decided:
        changes.append(f"배정 확정: {old_decided} → {new_decided} ({new_decided - old_decided:+d}명)")

    old_types = old.get("types", {})
    new_types = new.get("types", {})
    all_types = set(old_types) | set(new_types)
    for t in sorted(all_types):
        ov, nv = old_types.get(t, 0), new_types.get(t, 0)
        if ov != nv:
            changes.append(f"  · {t}: {ov} → {nv} ({nv - ov:+d}명)")

    return changes


def format_summary(snap: dict, changes: list[str]) -> str:
    now = datetime.now().strftime("%m/%d %H:%M")
    total = snap.get("total", "?")
    decided = snap.get("decided", "?")
    types = snap.get("types", {})

    type_lines = "\n".join(
        f"  · {t}: {c}명"
        for t, c in sorted(types.items(), key=lambda x: -x[1])
        if c > 0
    )

    change_block = ""
    if changes:
        change_block = "\n\n📌 <b>변경 사항</b>\n" + "\n".join(f"  {c}" for c in changes)

    return f"""🎓 <b>고입 진학현황 자동 동기화</b> ({now})

👥 전체: {total}명 | 확정: {decided}명
📊 유형별:
{type_lines or "  (데이터 없음)"}
{change_block}

✅ 동기화 완료"""


def main():
    dry_run = "--dry-run" in sys.argv
    force = "--force" in sys.argv

    log("=" * 60)
    log(f"고입 자동 동기화 시작 (dry={dry_run}, force={force})")
    log("=" * 60)

    errors = []

    # 1단계: 설문지 → 반별 시트
    log("[1/5] 설문지 → 반별 시트 동기화...")
    ok, out = run_script(GENERATORS_DIR / "sync_form_to_class_sheets.py", ["2026"], dry_run)
    if not ok:
        errors.append(f"설문지 동기화 실패: {out[-200:]}")
        log(f"  ⚠ 실패 (계속 진행)")
    else:
        log("  ✅ 완료")

    # 2단계: 반별 → 입시_트래킹
    log("[2/5] 반별 → 입시_트래킹 유형 동기화...")
    ok, out = run_script(GENERATORS_DIR / "sync_type_to_tracking.py", ["2026"], dry_run)
    if not ok:
        errors.append(f"유형 동기화 실패: {out[-200:]}")
        log(f"  ⚠ 실패 (계속 진행)")
    else:
        log("  ✅ 완료")

    # 2.5단계: 반별 → 특별전형_트래킹
    log("[3/5] 반별 → 특별전형_트래킹 동기화...")
    ok, out = run_script(GENERATORS_DIR / "sync_special_to_tracking.py", ["2026"], dry_run)
    if not ok:
        errors.append(f"특별전형 동기화 실패: {out[-200:]}")
        log(f"  ⚠ 실패 (계속 진행)")
    else:
        log("  ✅ 완료")

    # 3단계: 진행현황 갱신
    log("[4/5] 트래킹 → 진행현황 시트 생성...")
    ok, out = run_script(GENERATORS_DIR / "build_progress_from_tracking.py", ["2026"], dry_run)
    if not ok:
        errors.append(f"진행현황 생성 실패: {out[-200:]}")
        log(f"  ⚠ 실패")
    else:
        log("  ✅ 완료")

    # 4단계: HTML 대시보드 생성
    log("[5/5] 정적 HTML 대시보드 생성...")
    ok, out = run_script(GENERATORS_DIR / "generate_dashboard.py", ["2026"], dry_run)
    if not ok:
        log(f"  ⚠ HTML 생성 실패 (계속 진행)")
    else:
        log("  ✅ reports/ HTML 3개 갱신 완료")

    # 4.5단계: 특별전형 HTML 생성
    ok, out = run_script(GENERATORS_DIR / "generate_special_report.py", ["2026"], dry_run)
    if not ok:
        log(f"  ⚠ 특별전형 HTML 생성 실패 (계속 진행)")
    else:
        log("  ✅ 특별전형_현황.html 갱신 완료")

    # 6단계: 변경 감지
    log("스냅샷 비교 중...")
    old_snap = load_snapshot()
    new_snap = fetch_tracking_snapshot() if not dry_run else {"total": 0, "hash": "dry", "timestamp": datetime.now().isoformat()}
    changes = detect_changes(old_snap, new_snap)

    if changes:
        log(f"  변경 감지: {len(changes)}건")
        for c in changes:
            log(f"    {c}")
    else:
        log("  변경 없음")

    # 7단계: Telegram 알림
    should_notify = bool(changes) or force or bool(errors)
    if should_notify and not dry_run:
        log("Telegram 알림 전송...")
        msg = format_summary(new_snap, changes)
        if errors:
            msg += "\n\n⚠️ <b>오류 발생</b>\n" + "\n".join(f"  · {e}" for e in errors)
        send_telegram(msg)
        log("  ✅ 전송 완료")
    elif not should_notify:
        log("변경 없음 — Telegram 생략")

    # 스냅샷 저장
    if new_snap and not dry_run:
        save_snapshot(new_snap)

    log("=" * 60)
    log(f"완료. 오류={len(errors)}건, 변경={len(changes)}건")
    log("=" * 60)
    return 0 if not errors else 1


if __name__ == "__main__":
    sys.exit(main())
