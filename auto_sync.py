#!/usr/bin/env python3
"""
고입 진학현황 자동 동기화 & 알림 시스템

cron (매일 07:30) 또는 수동 실행:
  $WORKSPACE_DIR/unified_venv/bin/python auto_sync.py [--force] [--dry-run]

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
import time
from datetime import datetime
from pathlib import Path

# ── 경로 설정 ────────────────────────────────────────────────────────────────
WORKSPACE_ROOT = Path(os.getenv("WORKSPACE_DIR", os.path.expanduser("~/projects")))
PROJECT_DIR = WORKSPACE_ROOT / "Project_HighSchool_apply_Dashboard"
GENERATORS_DIR = PROJECT_DIR / "generators"
SNAPSHOT_FILE = PROJECT_DIR / ".auto_sync_snapshot.json"
LOG_FILE = PROJECT_DIR / "logs" / "auto_sync.log"
VENV_PYTHON = str(WORKSPACE_ROOT / "unified_venv/bin/python")
ENV_FILE = WORKSPACE_ROOT / ".secrets" / ".env"


def load_env_file(path: Path):
    if not path.exists():
        return

    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"'")
        os.environ.setdefault(key, value)


load_env_file(ENV_FILE)

# ── Telegram ─────────────────────────────────────────────────────────────────
TELEGRAM_BOT_TOKEN = os.getenv("ATTENDANCE_TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("ATTENDANCE_TELEGRAM_CHAT_ID")

# ── Google Sheets ─────────────────────────────────────────────────────────────
SERVICE_KEY = str(WORKSPACE_ROOT / ".secrets/service_key.json")
SPREADSHEET_ID_2026 = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"


def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def send_telegram(text: str):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        log("  Telegram 자격증명 누락: ATTENDANCE_TELEGRAM_TOKEN / ATTENDANCE_TELEGRAM_CHAT_ID 확인")
        return

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


def is_quota_error(text: str) -> bool:
    return "Quota exceeded" in text or "APIError: [429]" in text


def run_script(script_path: Path, args: list[str] = None, dry_run: bool = False) -> tuple[bool, str]:
    """스크립트 실행 → (성공 여부, 출력)"""
    cmd = [VENV_PYTHON, str(script_path)] + (args or [])
    if dry_run:
        log(f"  [DRY-RUN] {' '.join(cmd)}")
        return True, "[dry-run]"

    last_output = ""
    for attempt in range(1, 3):
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
                cwd=str(GENERATORS_DIR),
            )
            output = result.stdout + result.stderr
            last_output = output
            if result.returncode == 0:
                return True, output

            if attempt == 1 and is_quota_error(output):
                log("  Google Sheets 읽기 쿼터 초과 — 65초 대기 후 1회 재시도")
                time.sleep(65)
                continue

            log(f"  FAIL (rc={result.returncode}): {output[-500:]}")
            return False, output
        except subprocess.TimeoutExpired:
            last_output = "Timeout (120s)"
            return False, last_output
        except Exception as e:
            last_output = str(e)
            if attempt == 1 and is_quota_error(last_output):
                log("  Google Sheets 읽기 쿼터 초과 — 65초 대기 후 1회 재시도")
                time.sleep(65)
                continue
            return False, last_output

    return False, last_output


SPECIAL_NAMES = ["사회통합전형", "특례", "보훈", "쌍둥이", "학폭", "교직원자녀", "장애", "다자녀(3인+)"]


def _col(header: list[str], *names: str, default: int | None = None) -> int | None:
    for name in names:
        if name in header:
            return header.index(name)
    return default


def _get(row: list[str], idx: int | None) -> str:
    return row[idx].strip() if idx is not None and idx < len(row) else ""


def _preferred_col(header: list[str], primary: str, fallback: str) -> int | None:
    return _col(header, primary, fallback)


def fetch_sync_data() -> tuple[object, dict[str, object], list[list[str]], list[list[str]]]:
    """후반 단계에서 공유할 시트 데이터를 한 번만 읽는다."""
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_file(
        SERVICE_KEY,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SPREADSHEET_ID_2026)
    worksheets = {ws.title: ws for ws in ss.worksheets()}

    tracking = worksheets["입시_트래킹"]
    tracking_rows = tracking.get_all_values()

    special_rows: list[list[str]] = []
    special = worksheets.get("특별전형_트래킹")
    if special is not None:
        special_rows = special.get_all_values()

    return ss, worksheets, tracking_rows, special_rows


def fetch_sync_data_with_retry() -> tuple[object, dict[str, object], list[list[str]], list[list[str]]]:
    last_error: Exception | None = None
    for attempt in range(1, 3):
        try:
            return fetch_sync_data()
        except Exception as e:
            last_error = e
            if attempt == 1 and is_quota_error(str(e)):
                log("  공유 데이터 읽기 쿼터 초과 — 65초 대기 후 1회 재시도")
                time.sleep(65)
                continue
            raise

    raise last_error or RuntimeError("공유 데이터 로드 실패")


def build_progress_rows(tracking_rows: list[list[str]]) -> list[list[str]]:
    if not tracking_rows:
        return [["반", "번호", "성명", "성별", "지원유형", "1차", "2차", "최종", "비고"]]

    header = tracking_rows[0]
    c_cls = _col(header, "반", default=0)
    c_num = _col(header, "번호", default=1)
    c_name = _col(header, "성명", "이름", default=2)
    c_gender = _col(header, "성별", default=3)
    c_type = _preferred_col(header, "최종유형", "유형")
    c_first = _col(header, "1차")
    c_second = _col(header, "2차")
    c_final = _col(header, "최종")
    c_school = _preferred_col(header, "최종학교", "지원학교")
    c_major = _col(header, "학과")
    c_grade = _col(header, "학년")
    c_early = _col(header, "조기졸업여부")

    required = {
        "최종유형/유형": c_type,
        "1차": c_first,
        "2차": c_second,
        "최종": c_final,
    }
    missing = [name for name, idx in required.items() if idx is None]
    if missing:
        raise ValueError(f"입시_트래킹 필수 컬럼 없음: {', '.join(missing)}")

    progress_rows = [["반", "번호", "성명", "성별", "지원유형", "1차", "2차", "최종", "비고"]]
    for row in tracking_rows[1:]:
        if not _get(row, c_name):
            continue

        school_type = _get(row, c_type)
        if not school_type:
            continue

        school = _get(row, c_school)
        major = _get(row, c_major)
        grade = _get(row, c_grade)
        early_grad = _get(row, c_early).upper() == "O"
        remark = f"조기졸업({grade}학년)" if early_grad and grade else ("조기졸업" if early_grad else "")
        if school:
            remark += f" / 지원: {school}" if remark else f"지원: {school}"
        if major:
            remark += f" / {major}" if remark else f"학과: {major}"

        progress_rows.append([
            _get(row, c_cls),
            _get(row, c_num),
            _get(row, c_name),
            _get(row, c_gender),
            school_type,
            _get(row, c_first),
            _get(row, c_second),
            _get(row, c_final),
            remark,
        ])

    return progress_rows


def update_progress_sheet(ss, worksheets: dict[str, object], progress_rows: list[list[str]]) -> None:
    import gspread

    try:
        progress = worksheets.get("입시 진행 현황") or ss.worksheet("입시 진행 현황")
        progress.clear()
    except gspread.exceptions.WorksheetNotFound:
        progress = ss.add_worksheet("입시 진행 현황", rows=5000, cols=10)

    progress.update(values=progress_rows, range_name="A1")


def dashboard_students_from_rows(tracking_rows: list[list[str]]) -> list[dict[str, str]]:
    if not tracking_rows:
        return []

    header = tracking_rows[0]
    c_cls = _col(header, "반", default=0)
    c_num = _col(header, "번호", default=1)
    c_name = _col(header, "이름", "성명", default=2)
    c_gender = _col(header, "성별", default=3)
    c_type = _preferred_col(header, "최종유형", "유형")
    c_school = _preferred_col(header, "최종학교", "지원학교")
    c_final = _col(header, "최종")
    c_grade = _col(header, "학년")
    c_early = _col(header, "조기졸업여부")
    if c_type is None:
        raise ValueError("입시_트래킹 '최종유형' 또는 '유형' 컬럼 없음")

    students = []
    for row in tracking_rows[1:]:
        if not _get(row, c_name):
            continue

        type_val = _get(row, c_type)
        if not type_val:
            continue

        final_val = _get(row, c_final)
        result = ""
        if final_val.lower() in ["합격", "pass", "o", "○", "yes", "v"]:
            result = "합격"
        elif final_val.lower() in ["불합격", "fail", "x", "no"]:
            result = "불합격"

        students.append({
            "class": _get(row, c_cls),
            "num": _get(row, c_num),
            "name": _get(row, c_name),
            "gender": _get(row, c_gender),
            "type": type_val,
            "school": _get(row, c_school) or type_val,
            "result": result,
            "grade": _get(row, c_grade),
            "early_grad": _get(row, c_early).upper() == "O",
        })

    return students


def generate_dashboard_reports(tracking_rows: list[list[str]], year: str = "2026") -> int:
    sys.path.insert(0, str(GENERATORS_DIR))
    import generate_dashboard

    students = dashboard_students_from_rows(tracking_rows)
    if not students:
        return 0

    early = [s for s in students if s["type"] in generate_dashboard.EARLY_TYPES]
    late = [s for s in students if s["type"] in generate_dashboard.LATE_TYPES]
    out = lambda name: str(Path(generate_dashboard.OUTPUT_DIR) / name)

    count = 0
    if early:
        generate_dashboard.generate_html(early, f"{year} 전기고 지원 현황", out("전기고_현황.html"), year)
        count += 1
    if late:
        generate_dashboard.generate_html(late, f"{year} 후기고 지원 현황", out("후기고_현황.html"), year)
        count += 1
    generate_dashboard.generate_html(students, f"{year} 전체 진학 현황", out("전체_현황.html"), year)
    count += 1
    return count


def open_report_page() -> bool:
    """전체 현황 HTML을 기본 보고서로 연다."""
    import webbrowser

    report_path = (PROJECT_DIR / "reports" / "전체_현황.html").resolve()
    if not report_path.exists():
        return False

    url = f"file://{report_path}"
    if webbrowser.open(url):
        return True

    openers = [
        ["xdg-open", str(report_path)],
        ["gio", "open", str(report_path)],
        ["firefox", str(report_path)],
    ]
    for cmd in openers:
        try:
            subprocess.Popen(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
            return True
        except FileNotFoundError:
            continue
        except Exception:
            continue

    return False


def special_report_students_from_rows(
    special_rows: list[list[str]],
    tracking_rows: list[list[str]],
) -> list[dict[str, object]]:
    sys.path.insert(0, str(GENERATORS_DIR))
    import generate_special_report

    if not special_rows:
        return []

    special_header = special_rows[0]
    c_cls = _col(special_header, "반", default=0)
    c_num = _col(special_header, "번호", default=1)
    c_name = _col(special_header, "이름", "성명", default=2)
    c_gender = _col(special_header, "성별", default=3)

    special_data = {}
    for row in special_rows[1:]:
        if not _get(row, c_name):
            continue

        cats = []
        for cat in SPECIAL_NAMES:
            ci = _col(special_header, cat)
            if ci is not None and _get(row, ci).upper() == "O":
                cats.append(cat)

        key = (_get(row, c_cls), _get(row, c_num))
        special_data[key] = {
            "cls": _get(row, c_cls),
            "num": _get(row, c_num),
            "name": _get(row, c_name),
            "gender": _get(row, c_gender),
            "cats": cats,
        }

    tracking_map = {}
    if tracking_rows:
        header = tracking_rows[0]
        tc_cls = _col(header, "반", default=0)
        tc_num = _col(header, "번호", default=1)
        tc_type = _preferred_col(header, "최종유형", "유형")
        tc_school = _preferred_col(header, "최종학교", "지원학교")
        tc_final = _col(header, "최종")
        for row in tracking_rows[1:]:
            key = (_get(row, tc_cls), _get(row, tc_num))
            if not key[1]:
                continue
            tracking_map[key] = {
                "type": _get(row, tc_type),
                "school": _get(row, tc_school),
                "status": _get(row, tc_final),
            }

    students = []
    for key, sp in special_data.items():
        tr = tracking_map.get(key, {})
        type_val = str(tr.get("type", ""))
        students.append({
            **sp,
            "type": type_val,
            "school": tr.get("school", ""),
            "status": tr.get("status", ""),
            "group": generate_special_report.school_group(type_val),
        })

    return students


def generate_special_report_from_rows(
    special_rows: list[list[str]],
    tracking_rows: list[list[str]],
    year: str = "2026",
) -> bool:
    sys.path.insert(0, str(GENERATORS_DIR))
    import generate_special_report

    students = special_report_students_from_rows(special_rows, tracking_rows)
    if not students:
        return False

    out = Path(generate_special_report.OUTPUT_DIR) / "특별전형_현황.html"
    generate_special_report.generate_html(students, year, str(out))
    return True


def snapshot_from_rows(
    tracking_rows: list[list[str]],
    special_rows: list[list[str]] | None = None,
) -> dict:
    if not tracking_rows:
        return {}

    header = tracking_rows[0]
    data_rows = [r for r in tracking_rows[1:] if any(c.strip() for c in r)]
    type_col = _preferred_col(header, "희망유형", "유형")
    # 배정확정은 진학부가 입력하는 최종배정학교만 인정 — 희망값 fallback 금지
    school_col = _col(header, "최종배정학교")

    type_counts: dict[str, int] = {}
    decided = 0
    for row in data_rows:
        if type_col is not None:
            t = _get(row, type_col)
            if t:
                type_counts[t] = type_counts.get(t, 0) + 1
        if school_col is not None and _get(row, school_col):
            decided += 1

    special_counts: dict[str, int] = {}
    if special_rows:
        sheader = special_rows[0]
        for name in SPECIAL_NAMES:
            col = _col(sheader, name)
            if col is not None:
                cnt = sum(1 for row in special_rows[1:] if _get(row, col) == "O")
                if cnt:
                    special_counts[name] = cnt

    return {
        "total": len(data_rows),
        "decided": decided,
        "types": type_counts,
        "special": special_counts,
        "hash": hashlib.md5(str(tracking_rows).encode()).hexdigest(),
        "timestamp": datetime.now().isoformat(),
    }


def fetch_tracking_snapshot() -> dict:
    """입시_트래킹 + 특별전형_트래킹 시트 현황 스냅샷"""
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_file(
            SERVICE_KEY,
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(SPREADSHEET_ID_2026)

        # ── 입시_트래킹 ──────────────────────────────────────────
        ws = ss.worksheet("입시_트래킹")
        rows = ws.get_all_values()
        if not rows:
            return {}

        header = rows[0]
        data_rows = [r for r in rows[1:] if any(c.strip() for c in r)]

        type_col = _preferred_col(header, "희망유형", "유형")
        school_col = _col(header, "최종배정학교")

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

        # ── 특별전형_트래킹 ──────────────────────────────────────
        special_counts: dict[str, int] = {}
        try:
            sws = ss.worksheet("특별전형_트래킹")
            srows = sws.get_all_values()
            if srows:
                sheader = srows[0]
                for name in SPECIAL_NAMES:
                    col = next((i for i, h in enumerate(sheader) if h.strip() == name), None)
                    if col is not None:
                        cnt = sum(1 for r in srows[1:] if col < len(r) and r[col].strip() == "O")
                        if cnt:
                            special_counts[name] = cnt
        except Exception:
            pass  # 특별전형_트래킹 없으면 생략

        return {
            "total": len(data_rows),
            "decided": decided,
            "types": type_counts,
            "special": special_counts,
            "hash": hashlib.md5(str(rows).encode()).hexdigest(),
            "timestamp": datetime.now().isoformat(),
        }
    except Exception as e:
        log(f"  스냅샷 조회 실패: {e}")
        return None  # None = 조회 실패, {} = 데이터 없음 구분


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
    special = snap.get("special", {})

    type_lines = "\n".join(
        f"  · {t}: {c}명"
        for t, c in sorted(types.items(), key=lambda x: -x[1])
        if c > 0
    )

    special_block = ""
    if special:
        special_lines = "  " + " | ".join(f"{k}: {v}명" for k, v in special.items())
        special_block = f"\n🏷 특별전형:\n{special_lines}"

    change_block = ""
    if changes:
        change_block = "\n\n📌 <b>변경 사항</b>\n" + "\n".join(f"  {c}" for c in changes)

    return f"""🎓 <b>고입 진학현황 자동 동기화</b> ({now})

👥 전체: {total}명 | 확정: {decided}명
📊 유형별:
{type_lines or "  (데이터 없음)"}
{special_block}{change_block}

✅ 동기화 완료"""


def main():
    dry_run = "--dry-run" in sys.argv
    force = "--force" in sys.argv

    log("=" * 60)
    log(f"고입 자동 동기화 시작 (dry={dry_run}, force={force})")
    log("=" * 60)

    errors = []
    report_opened = False

    # 1단계: 설문지 → 반별 시트
    log("[1/5] 설문지 → 반별 시트 동기화...")
    ok, out = run_script(GENERATORS_DIR / "sync_form_to_class_sheets.py", ["2026"], dry_run)
    if not ok:
        errors.append(f"설문지 동기화 실패: {out[-200:]}")
        log(f"  ⚠ 실패 (계속 진행)")
    else:
        log("  ✅ 완료")

    # 2단계: 반별 → 입시_트래킹
    log("[2/5] 반별/조기졸업 → 입시_트래킹 희망/최종 schema 동기화...")
    ok, out = run_script(
        GENERATORS_DIR / "sync_type_to_tracking.py",
        ["2026", "--apply-schema", "--no-legacy-fill"],
        dry_run,
    )
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

    ss = None
    worksheets = {}
    tracking_rows: list[list[str]] = []
    special_rows: list[list[str]] = []
    shared_data_failed = False

    if not dry_run:
        log("공유 시트 데이터 읽는 중... (입시_트래킹/특별전형_트래킹)")
        try:
            ss, worksheets, tracking_rows, special_rows = fetch_sync_data_with_retry()
            log(f"  ✅ 공유 데이터 로드 완료: 입시_트래킹 {max(len(tracking_rows)-1, 0)}명, 특별전형_트래킹 {max(len(special_rows)-1, 0)}명")
        except Exception as e:
            shared_data_failed = True
            errors.append(f"공유 데이터 로드 실패: {e}")
            log(f"  ⚠ 공유 데이터 로드 실패: {e}")

    # 3단계: 진행현황 갱신
    log("[4/5] 트래킹 → 진행현황 시트 생성...")
    if dry_run:
        log("  [DRY-RUN] 공유 데이터 기반 진행현황 생성 생략")
    elif ss is None:
        log("  ⚠ 공유 데이터 없음 — 진행현황 생성 생략")
    else:
        try:
            progress_rows = build_progress_rows(tracking_rows)
            update_progress_sheet(ss, worksheets, progress_rows)
            log(f"  ✅ 완료 ({len(progress_rows)-1}행)")
        except Exception as e:
            errors.append(f"진행현황 생성 실패: {e}")
            log(f"  ⚠ 실패: {e}")

    # 4단계: HTML 대시보드 생성
    log("[5/5] 정적 HTML 대시보드 생성...")
    if dry_run:
        log("  [DRY-RUN] 공유 데이터 기반 HTML 생성 생략")
    elif not tracking_rows:
        log("  ⚠ 입시_트래킹 데이터 없음 — HTML 생성 생략")
    else:
        try:
            html_count = generate_dashboard_reports(tracking_rows, "2026")
            log(f"  ✅ reports/ HTML {html_count}개 갱신 완료")
            if open_report_page():
                report_opened = True
                log("  ✅ 전체_현황.html 브라우저 열기 요청 완료")
            else:
                log("  ⚠ 전체_현황.html 브라우저 열기 실패")
        except Exception as e:
            log(f"  ⚠ HTML 생성 실패 (계속 진행): {e}")

    # 4.5단계: 특별전형 HTML 생성
    if dry_run:
        log("  [DRY-RUN] 공유 데이터 기반 특별전형 HTML 생성 생략")
    elif not special_rows:
        log("  ⚠ 특별전형_트래킹 데이터 없음 — 특별전형 HTML 생성 생략")
    else:
        try:
            if generate_special_report_from_rows(special_rows, tracking_rows, "2026"):
                log("  ✅ 특별전형_현황.html 갱신 완료")
            else:
                log("  ⚠ 특별전형 HTML 생성 대상 없음")
        except Exception as e:
            log(f"  ⚠ 특별전형 HTML 생성 실패 (계속 진행): {e}")

    if not dry_run and not report_opened:
        if open_report_page():
            report_opened = True
            log("  ✅ 기존 전체_현황.html 브라우저 열기 요청 완료")
        else:
            log("  ⚠ 기존 전체_현황.html 브라우저 열기 실패")

    # 6단계: 변경 감지
    log("스냅샷 비교 중...")
    old_snap = load_snapshot()
    if dry_run:
        new_snap = {"total": 0, "hash": "dry", "timestamp": datetime.now().isoformat()}
        changes = []
    elif tracking_rows:
        new_snap = snapshot_from_rows(tracking_rows, special_rows)
        changes = detect_changes(old_snap, new_snap)
    elif shared_data_failed:
        new_snap = None
        log("  ⚠ 공유 데이터 로드 실패 — 변경 감지 건너뜀 (추가 조회 방지)")
        changes = []
    else:
        new_snap = fetch_tracking_snapshot()
        if new_snap is None:
            # Quota 초과 등 조회 실패 — diff 건너뜀 (오보 방지)
            log("  ⚠ 스냅샷 조회 실패 — 변경 감지 건너뜀 (오보 방지)")
            changes = []
        else:
            changes = detect_changes(old_snap, new_snap)

    if changes:
        log(f"  변경 감지: {len(changes)}건")
        for c in changes:
            log(f"    {c}")
    else:
        log("  변경 없음")

    # 7단계: Telegram 알림
    should_notify = bool(changes) or force or bool(errors)
    if should_notify and not dry_run and new_snap is not None:
        log("Telegram 알림 전송...")
        msg = format_summary(new_snap, changes)
        if errors:
            msg += "\n\n⚠️ <b>오류 발생</b>\n" + "\n".join(f"  · {e}" for e in errors)
        send_telegram(msg)
        log("  ✅ 전송 완료")
    elif force and new_snap is None and not dry_run:
        # --force인데 스냅샷 실패 → 오류 알림만
        log("Telegram 오류 알림 전송...")
        send_telegram("⚠️ 고입 자동 동기화: 스냅샷 조회 실패 (API Quota 초과)\n동기화 자체는 완료됨")
        log("  ✅ 전송 완료")
    elif not should_notify:
        log("변경 없음 — Telegram 생략")

    # 스냅샷 저장 (조회 성공 시만)
    if new_snap is not None and not dry_run:
        save_snapshot(new_snap)

    log("=" * 60)
    log(f"완료. 오류={len(errors)}건, 변경={len(changes)}건")
    log("=" * 60)
    return 0 if not errors else 1


if __name__ == "__main__":
    sys.exit(main())
