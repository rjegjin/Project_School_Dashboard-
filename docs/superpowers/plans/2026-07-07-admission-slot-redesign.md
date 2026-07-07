# 고입 시기 슬롯 구조 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 반별 시트의 접수·결과 기록을 유형 축(영재·과학·자사 3종)에서 시기 축(영재고/전기/후기 슬롯 × 접수학교+결과)으로 재편하고, 오염된 최종 데이터·진행현황 과소집계·Sheets 쿼터 초과를 같은 마이그레이션으로 해소한다.

**Architecture:** 반별 시트 V:AA를 3슬롯×2칸(접수학교, 결과)으로 재정의. `sync_type_to_tracking.py`의 접수 파서를 슬롯 파서로 교체하고, 학교명→유형 분류는 신규 공용 모듈 `generators/school_types.py`가 담당. 입시_트래킹의 기존 schema 컬럼 5개(AF:AJ)는 헤더 rename으로 재사용하고 나머지는 sync의 apply-schema가 자동 append. `최종배정학교`는 진학부 전용 — sync가 절대 쓰지 않는다.

**Tech Stack:** Python 3 (`/home/rjegj/projects/unified_venv/bin/python`), gspread(`values_batch_get`), pytest 9.1.1, Google Sheets (2026 스프레드시트 `14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0`).

## Global Constraints

- 모든 Python 실행: `/home/rjegj/projects/unified_venv/bin/python` (로컬 venv 생성 금지)
- Sheets 인증: `sheets_client.get_client()` (mh-common) 또는 기존 파일의 `Credentials.from_service_account_file("/home/rjegj/projects/.secrets/service_key.json")` 패턴 유지 — oauth2client 금지
- 작업 디렉터리(서브 repo): `/home/rjegj/projects/Project_HighSchool_apply_Dashboard`
- 시트에 쓰는 스크립트는 전부 dry-run 기본, `--apply` 명시 시에만 쓰기
- 커밋 prefix: `feat:`/`fix:`/`chore:`, 커밋 메시지 끝에 `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`
- **결과 코드 표준값**: `1차합` `1차불` `2차합` `2차불` `최종합` `최종불` `포기` (빈칸=대기)
- **유형 표준값**: `영재고` `과학고` `예술계고` `특성화고` `자사고` `외고/국제고` `일반고` `기타/대안`
- **반별 시트 새 레이아웃 (0-based col idx)**: V=21 `영재고_접수학교`, W=22 `영재고_결과`, X=23 `전기_접수학교`, Y=24 `전기_결과`, Z=25 `후기_접수학교`, AA=26 `후기_결과`, AB=27 폐기(빈 헤더)
- **입시_트래킹 새 schema 컬럼**: 학년, 학생ID, 희망유형, 희망학교, 영재고_접수, 영재고_결과, 전기_접수학교, 전기_유형, 전기_결과, 후기_접수학교, 후기_유형, 후기_결과, 최종배정학교, 데이터상태, 조기졸업여부, 출처 — 이 중 `최종배정학교`는 sync가 헤더만 보장하고 값은 절대 쓰지 않음
- 스펙 문서: `docs/superpowers/specs/2026-07-07-admission-slot-redesign-design.md`

---

### Task 1: 학교명→유형 분류기 `school_types.py`

**Files:**
- Create: `generators/school_types.py`
- Test: `tests/test_school_types.py`

**Interfaces:**
- Produces: `classify_school(name: str) -> str` — 유형 표준값 또는 `""`(미분류). Task 2, 6, 7이 import.
- Produces: `SCHOOL_TYPE_OVERRIDES: dict[str, str]` — 키워드로 못 잡는 학교명(자사고 등) 수동 등록 사전.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_school_types.py`:

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "generators"))

from school_types import classify_school


def test_keyword_classification():
    assert classify_school("한성과학고") == "과학고"
    assert classify_school("서울예술고등학교") == "예술계고"
    assert classify_school("수도전기공업고(마이스터)") == "특성화고"
    assert classify_school("대원외고") == "외고/국제고"
    assert classify_school("동탄국제고") == "외고/국제고"
    assert classify_school("경기북과학고") == "과학고"
    assert classify_school("서울과학영재학교") == "영재고"


def test_override_wins_over_keyword():
    assert classify_school("하나고") == "자사고"
    assert classify_school("민족사관고") == "자사고"


def test_unknown_returns_empty():
    assert classify_school("알수없는고등학교") == ""
    assert classify_school("") == ""
    assert classify_school("  ") == ""


def test_whitespace_and_suffix_normalization():
    assert classify_school(" 한성과학고등학교 ") == "과학고"
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `cd /home/rjegj/projects/Project_HighSchool_apply_Dashboard && /home/rjegj/projects/unified_venv/bin/python -m pytest tests/test_school_types.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'school_types'`

- [ ] **Step 3: 최소 구현**

`generators/school_types.py`:

```python
"""학교명 → 고입 유형 분류기 (시기 슬롯 구조 공용).

키워드로 못 잡는 학교(자사고는 이름에 '자사'가 없다)는
SCHOOL_TYPE_OVERRIDES에 등록한다. 미분류("")는 sync가 '확인필요'로
플래그하므로, 플래그가 뜰 때마다 여기에 추가하면 된다.
"""

# ponytail: 수동 사전 = 미분류 플래그가 뜰 때 채우는 calibration knob
SCHOOL_TYPE_OVERRIDES = {
    "하나고": "자사고",
    "민족사관고": "자사고",
    "민사고": "자사고",
    "상산고": "자사고",
    "외대부고": "자사고",
    "북일고": "자사고",
    "인천하늘고": "자사고",
}

# 순서 중요: 먼저 매칭되는 키워드가 이긴다 (영재 > 과학, 국제중학교 같은 오탐 없음)
_KEYWORD_RULES = [
    ("영재", "영재고"),
    ("과학고", "과학고"),
    ("과고", "과학고"),
    ("예술", "예술계고"),
    ("예고", "예술계고"),
    ("체육", "예술계고"),
    ("체고", "예술계고"),
    ("마이스터", "특성화고"),
    ("외국어", "외고/국제고"),
    ("외고", "외고/국제고"),
    ("국제고", "외고/국제고"),
]


def classify_school(name: str) -> str:
    """학교명에서 유형 표준값을 판정한다. 미분류는 ""."""
    cleaned = str(name).strip()
    if not cleaned:
        return ""
    base = cleaned.replace("등학교", "").rstrip()  # "고등학교" -> "고"
    for key, type_name in SCHOOL_TYPE_OVERRIDES.items():
        if key in base:
            return type_name
    for keyword, type_name in _KEYWORD_RULES:
        if keyword in base:
            return type_name
    return ""
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `/home/rjegj/projects/unified_venv/bin/python -m pytest tests/test_school_types.py -v`
Expected: 4 passed

- [ ] **Step 5: 커밋**

```bash
git add generators/school_types.py tests/test_school_types.py
git commit -m "feat: 학교명→유형 분류기 school_types 추가 (시기 슬롯 구조 공용)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 2: `sync_type_to_tracking.py` 슬롯 파서 전환

**Files:**
- Modify: `generators/sync_type_to_tracking.py` (RECEIPT_TYPE_MAP 파서 → 슬롯 파서, SCHEMA_COLUMNS 교체, data_status 재정의)
- Test: `tests/test_slot_parser.py`

**Interfaces:**
- Consumes: `school_types.classify_school(name) -> str` (Task 1)
- Produces: `detect_slots(row: list[str]) -> dict[str, dict]` — 키 `"영재고" | "전기" | "후기"`, 값 `{"school": str, "result": str, "flags": list[str]}`. flags 값: `"유형미분류"`, `"결과코드비표준"`.
- Produces: `data_status(hope_type, slots, final_assigned) -> str` — `미입력`/`희망만`/`영재고진행`/`전기진행`/`후기진행`/`배정완료`/`확인필요`
- Produces: 트래킹 새 schema 컬럼 헤더 (Global Constraints 참조) — Task 4, 5, 6, 7이 이 컬럼명으로 읽는다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_slot_parser.py`:

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "generators"))

from sync_type_to_tracking import detect_slots, data_status


def _row(v="", w="", x="", y="", z="", aa=""):
    """반별 시트 한 행: A:U(21칸) 더미 + V:AA 슬롯값."""
    return [""] * 21 + [v, w, x, y, z, aa]


def test_empty_row_has_no_slots():
    slots = detect_slots(_row())
    assert all(s["school"] == "" and s["result"] == "" for s in slots.values())


def test_basic_slot_parsing():
    slots = detect_slots(_row(v="서울과학영재학교", w="1차합", x="한성과학고", y="", z="하나고", aa="최종불"))
    assert slots["영재고"] == {"school": "서울과학영재학교", "result": "1차합", "flags": []}
    assert slots["전기"] == {"school": "한성과학고", "result": "", "flags": []}
    assert slots["후기"] == {"school": "하나고", "result": "최종불", "flags": []}


def test_nonstandard_result_code_flagged():
    slots = detect_slots(_row(x="한성과학고", y="합격했음"))
    assert "결과코드비표준" in slots["전기"]["flags"]
    assert slots["전기"]["result"] == "합격했음"  # 값은 보존


def test_unclassified_school_flagged_in_data_status():
    slots = detect_slots(_row(z="알수없는고"))
    status = data_status("일반고", slots, final_assigned="")
    assert status == "확인필요"


def test_short_row_padding():
    slots = detect_slots([""] * 23 + ["한성과학고"])  # X열까지만 존재
    assert slots["전기"]["school"] == "한성과학고"


def test_data_status_progression():
    empty = detect_slots(_row())
    assert data_status("", empty, "") == "미입력"
    assert data_status("과학고", empty, "") == "희망만"
    assert data_status("과학고", detect_slots(_row(v="경기과학영재학교")), "") == "영재고진행"
    assert data_status("과학고", detect_slots(_row(x="한성과학고", y="1차합")), "") == "전기진행"
    assert data_status("과학고", detect_slots(_row(z="대원외고")), "") == "후기진행"
    assert data_status("과학고", detect_slots(_row(x="한성과학고")), "서울고") == "배정완료"


def test_later_slot_wins_for_status():
    slots = detect_slots(_row(v="경기과학영재학교", w="최종불", x="한성과학고"))
    assert data_status("과학고", slots, "") == "전기진행"
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `/home/rjegj/projects/unified_venv/bin/python -m pytest tests/test_slot_parser.py -v`
Expected: FAIL — `ImportError: cannot import name 'detect_slots'`

- [ ] **Step 3: 슬롯 파서 구현**

`generators/sync_type_to_tracking.py` 수정 내용:

(a) 상수 교체 — `RECEIPT_TYPE_MAP`, `RECEIPT_POSITIVE_MARKERS`, `RECEIPT_NEGATIVE_MARKERS` 삭제 후:

```python
from school_types import classify_school

# 반별 시트 시기 슬롯: (접수학교 col, 결과 col) 0-based
SLOT_COLS = {
    "영재고": (21, 22),  # V, W
    "전기": (23, 24),    # X, Y
    "후기": (25, 26),    # Z, AA
}
SLOT_ORDER = ["영재고", "전기", "후기"]
RESULT_CODES = {"1차합", "1차불", "2차합", "2차불", "최종합", "최종불", "포기"}
```

(b) `SCHEMA_COLUMNS` 교체:

```python
SCHEMA_COLUMNS = [
    "학년",
    "학생ID",
    "희망유형",
    "희망학교",
    "영재고_접수",
    "영재고_결과",
    "전기_접수학교",
    "전기_유형",
    "전기_결과",
    "후기_접수학교",
    "후기_유형",
    "후기_결과",
    "최종배정학교",
    "데이터상태",
    "조기졸업여부",
    "출처",
]
ADMIN_ONLY_COLUMNS = {"최종배정학교"}  # sync가 헤더만 만들고 값은 절대 쓰지 않음
```

(c) `detect_receipt()` 삭제, `detect_slots()` 추가:

```python
def detect_slots(row: list[str]) -> dict[str, dict]:
    """반별 시트 V:AA 시기 슬롯 블록을 판정한다."""
    row = row + [""] * (28 - len(row))
    slots: dict[str, dict] = {}
    for slot, (school_col, result_col) in SLOT_COLS.items():
        school = norm(row[school_col])
        result = norm(row[result_col])
        flags: list[str] = []
        if result and result not in RESULT_CODES:
            flags.append("결과코드비표준")
        if school:
            # 복수지원(쉼표)은 첫 학교 기준으로 유형 분류
            first = school.split(",")[0].strip()
            if slot != "영재고" and not classify_school(first):
                flags.append("유형미분류")
        slots[slot] = {"school": school, "result": result, "flags": flags}
    return slots
```

(d) `data_status()` 교체:

```python
def data_status(hope_type: str, slots: dict[str, dict], final_assigned: str) -> str:
    if final_assigned:
        return "배정완료"
    if any(s["flags"] for s in slots.values()):
        return "확인필요"
    for slot in reversed(SLOT_ORDER):  # 가장 늦은 시기 슬롯이 현재 진행 단계
        if slots[slot]["school"]:
            return f"{slot}진행"
    if hope_type:
        return "희망만"
    return "미입력"
```

(e) `SourceStudent`의 receipt_* 4개 필드를 `slots: tuple = ()`로 교체 (frozen dataclass이므로 dict 대신 `tuple(sorted(...))` 불가 — 간단히 `slots_json: str = ""`에 `json.dumps(slots, ensure_ascii=False, sort_keys=True)`로 저장하고 사용처에서 `json.loads`. `import json` 추가):

```python
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
    slots_json: str = "{}"

    @property
    def slots(self) -> dict:
        return json.loads(self.slots_json)
```

(f) `load_class_sources()`에서 `detect_receipt(...)` 호출부를 다음으로 교체:

```python
slots = detect_slots(row)
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
        slots_json=json.dumps(slots, ensure_ascii=False, sort_keys=True),
    )
)
```

(g) `main()`의 `c_receipt_*`/`c_final_*` 변수와 receipt 관련 conflicts 4블록 중 receipt 2블록 삭제. schema_values 교체:

```python
slots = source.slots
final_assigned = get_cell(row, columns.get("최종배정학교"))
gifted = slots.get("영재고", {})
early = slots.get("전기", {})
late = slots.get("후기", {})
schema_values = {
    "학년": source.grade,
    "학생ID": source.student_id(args.year),
    "희망유형": source.hope_type,
    "희망학교": source.hope_school,
    "영재고_접수": gifted.get("school", ""),
    "영재고_결과": gifted.get("result", ""),
    "전기_접수학교": early.get("school", ""),
    "전기_유형": classify_school(early.get("school", "").split(",")[0]),
    "전기_결과": early.get("result", ""),
    "후기_접수학교": late.get("school", ""),
    "후기_유형": classify_school(late.get("school", "").split(",")[0]),
    "후기_결과": late.get("result", ""),
    "데이터상태": data_status(source.hope_type, slots, final_assigned),
    "조기졸업여부": "O" if source.early_grad else "",
    "출처": source.source,
}
```

(주의: `최종배정학교`는 schema_values에 **넣지 않는다** — ADMIN_ONLY. 헤더 생성 루프는 SCHEMA_COLUMNS 전체를 돌므로 헤더는 보장된다.)

(h) 슬롯 플래그를 conflicts로 노출 — schema_values 계산 직전에:

```python
for slot_name, slot in slots.items():
    for flag in slot["flags"]:
        conflicts.append(
            f"row {row_num}: [{slot_name}] {flag}: 학교={slot['school']!r}, 결과={slot['result']!r}, source={source.source}"
        )
```

(i) dry-run 출력 분리 — `print_updates("희망/접수/최종 schema 갱신 후보", schema_updates)`를 슬롯별로:

```python
def _updates_for(prefix: str, updates):
    return [u for u in updates if prefix in u[3] or prefix in u[0]]

print_updates("schema 갱신 후보 (전체)", schema_updates)
```

reason 문자열에 컬럼명을 포함시키면 충분하므로 `set_if_changed` 호출 시 reason을 `f"{name} <- {source.source}"`로 변경한다 (기존 `f"schema sync from {source.source}"` 대체).

- [ ] **Step 4: 테스트 통과 확인**

Run: `/home/rjegj/projects/unified_venv/bin/python -m pytest tests/ -v`
Expected: test_slot_parser 7개 + test_school_types 4개 모두 PASS

- [ ] **Step 5: 문법·임포트 검증 (시트 접근 없이)**

Run: `/home/rjegj/projects/unified_venv/bin/python -c "import sys; sys.path.insert(0, 'generators'); import sync_type_to_tracking; print('OK')"`
Expected: `OK`

- [ ] **Step 6: 커밋**

```bash
git add generators/sync_type_to_tracking.py tests/test_slot_parser.py
git commit -m "feat: sync 접수 파서를 유형 3종에서 시기 슬롯(영재고/전기/후기)으로 교체

최종배정학교는 ADMIN_ONLY — sync가 절대 쓰지 않는다.
데이터상태를 슬롯 진행 단계 기반으로 재정의.

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 3: 반별 시트 batch 읽기 (쿼터 초과 해소)

**Files:**
- Modify: `generators/sync_type_to_tracking.py:199-231` (`load_class_sources`)
- Modify: `generators/sync_special_to_tracking.py:74-79` (반별 시트 루프)

**Interfaces:**
- Consumes: gspread `Spreadsheet.values_batch_get(ranges) -> dict` (unified_venv gspread에 존재 확인됨)
- Produces: 동작 동일, API 읽기 호출 수 반별 14회 → 1회

- [ ] **Step 1: `load_class_sources` batch 전환**

```python
def load_class_sources(ss) -> list[SourceStudent]:
    class_titles = [
        sht.title.strip()
        for sht in ss.worksheets()
        if sht.title.strip().isdigit() and len(sht.title.strip()) == 3 and sht.title.strip().startswith("3")
    ]
    if not class_titles:
        return []

    # ponytail: 14개 시트 개별 read가 분당 쿼터를 치던 원인 — 1회 batch로 통합
    response = ss.values_batch_get(ranges=[f"'{t}'!A1:AB80" for t in class_titles])
    students: list[SourceStudent] = []
    for title, value_range in zip(class_titles, response["valueRanges"]):
        rows = value_range.get("values", [])
        data_start = 2 if len(rows) > 2 and len(rows[1]) > 2 and not norm(rows[1][2]) else 1
        for row_idx, row in enumerate(rows[data_start:], start=data_start + 1):
            if len(row) < 3 or not norm(row[2]):
                continue
            hope_type = detect_type(row)
            hope_school = detect_school(row, hope_type)
            slots = detect_slots(row)
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
                    slots_json=json.dumps(slots, ensure_ascii=False, sort_keys=True),
                )
            )
    return students
```

- [ ] **Step 2: `sync_special_to_tracking.py` 동일 패턴 적용**

기존 74~79행 루프(시트별 `get_all_values`)를 위와 같은 `values_batch_get` 1회 호출로 교체한다. 기존 파싱 로직(컬럼 인덱스 E,F,G,Q,R,S,T,U)은 그대로 유지 — 읽기 방식만 바꾼다.

- [ ] **Step 3: 검증 (읽기 전용)**

Run: `/home/rjegj/projects/unified_venv/bin/python generators/sync_type_to_tracking.py 2026 --dry-run --apply-schema --no-legacy-fill 2>&1 | head -30`
Expected: `[1/3] 소스 로드: 4xx명` — 기존과 같은 인원수, 쿼터 오류 없음. (아직 반별 시트 헤더가 구형이므로 슬롯 값은 대부분 빈칸 — 정상)

- [ ] **Step 4: 커밋**

```bash
git add generators/sync_type_to_tracking.py generators/sync_special_to_tracking.py
git commit -m "fix: 반별 시트 읽기를 values_batch_get 1회로 통합 — 쿼터 초과 65초 재시도 제거

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 4: 마이그레이션 스크립트

**Files:**
- Create: `generators/migrate_to_slots.py`

**Interfaces:**
- Consumes: `school_types.classify_school` (Task 1)
- Produces: 반별 시트 V1:AB1 새 헤더 + 기존 24행(V:X 14, Y:AB 10) 이관 + 트래킹 AF1:AJ1 헤더 rename(`접수유형→영재고_접수, 접수학교→영재고_결과, 접수상태→전기_접수학교, 최종유형→전기_유형, 최종학교→전기_결과`) + 해당 5컬럼 데이터 클리어. dry-run 기본.

- [ ] **Step 1: 스크립트 작성**

`generators/migrate_to_slots.py`:

```python
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

from school_types import classify_school

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
            ws.batch_update([{"range": f"V{r}:AB{r}", "values": [new + [""]]} for r, new in row_plans])
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
```

- [ ] **Step 2: dry-run 실행 및 이관 계획 육안 확인**

Run: `/home/rjegj/projects/unified_venv/bin/python generators/migrate_to_slots.py`
Expected: 이관 계획 약 24행 미만 출력 (V:X 14행 + Y:AB 10행, 중복 행 있으면 그 이하), 오류 없음. **출력을 사용자에게 보여주고 확인받은 뒤** 다음 단계 진행.

- [ ] **Step 3: 적용**

Run: `/home/rjegj/projects/unified_venv/bin/python generators/migrate_to_slots.py --apply`
Expected: 백업 경로 출력 → 시트별 `✅` → 트래킹 클리어 완료

- [ ] **Step 4: 적용 결과 검증**

Run: `/home/rjegj/projects/unified_venv/bin/python -c "
import sys; sys.path.insert(0, 'generators')
from sheets_client import get_client
ss = get_client().open_by_key('14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0')
print(ss.worksheet('301').get('V1:AB1'))
hdr = ss.worksheet('입시_트래킹').row_values(1)
print([h for h in hdr if '접수' in h or '결과' in h or '배정' in h])"`
Expected: 301 헤더 = 새 6칸+빈칸, 트래킹에 `영재고_접수`, `전기_접수학교` 등 rename된 헤더 존재

- [ ] **Step 5: 커밋**

```bash
git add generators/migrate_to_slots.py
git commit -m "feat: 시기 슬롯 구조 마이그레이션 스크립트 (헤더 교체 + 24행 이관 + 트래킹 클리어)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 5: `build_progress_from_tracking.py` 기준 컬럼 교체 (8행 과소집계 해소)

**Files:**
- Modify: `generators/build_progress_from_tracking.py:56-99` (컬럼 선택·행 생성), `:80` (출력 헤더)

**Interfaces:**
- Consumes: 트래킹 새 schema 컬럼 (Task 2)
- Produces: `입시 진행 현황` 시트 헤더 `["반", "번호", "성명", "성별", "희망유형", "영재고", "전기", "후기", "최종배정", "비고"]`. Task 7(대시보드 TAB 5/8)이 이 헤더를 읽는다.

- [ ] **Step 1: 집계 기준 교체**

기존 68행 `type_col = col_map.get("최종유형", col_map.get("유형"))` → 집계 대상을 "희망유형 또는 슬롯 접수가 하나라도 있는 학생"으로 변경:

```python
def get(row, name):
    idx = col_map.get(name)
    return str(row[idx]).strip() if idx is not None and idx < len(row) else ""

def slot_cell(row, school_col, result_col):
    school, result = get(row, school_col), get(row, result_col)
    return f"{school} {result}".strip()

progress_rows = []
for row in tracking_rows[1:]:
    hope = get(row, "희망유형")
    gifted = slot_cell(row, "영재고_접수", "영재고_결과")
    early = slot_cell(row, "전기_접수학교", "전기_결과")
    late = slot_cell(row, "후기_접수학교", "후기_결과")
    assigned = get(row, "최종배정학교")
    if not (hope or gifted or early or late or assigned):
        continue
    progress_rows.append([
        get(row, "반"), get(row, "번호"), get(row, "성명"), get(row, "성별"),
        hope, gifted, early, late, assigned, get(row, "데이터상태"),
    ])
```

출력 헤더(기존 80행)를 `["반", "번호", "성명", "성별", "희망유형", "영재고", "전기", "후기", "최종배정", "비고"]`로 교체. 기존 "최종 유형이 없으면 스킵" 로직(99행 부근)은 삭제. 유형별 분포 출력(165행)은 `희망유형` Counter로 교체.

- [ ] **Step 2: 실행 검증**

Run: `/home/rjegj/projects/unified_venv/bin/python generators/build_progress_from_tracking.py 2026`
Expected: 행 수가 희망 입력 학생 규모(150명 이상)로 복원. 8행이 아님.

- [ ] **Step 3: 커밋**

```bash
git add generators/build_progress_from_tracking.py
git commit -m "fix: 진행현황 집계 기준을 최종유형에서 희망+슬롯으로 교체 — 8행 과소집계 해소

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 6: `generate_final_sheets.py` 합격 판정 교체

**Files:**
- Modify: `generators/generate_final_sheets.py:66-118` (합격자 추출 로직)

**Interfaces:**
- Consumes: 트래킹 슬롯 컬럼 + `school_types.classify_school` (Task 1)
- Produces: `전기고_최종`/`후기고_최종` 시트 — 기존 EARLY_SECTIONS/LATE_SECTIONS 구조 유지

- [ ] **Step 1: 판정 로직 교체**

기존 77-94행(최종유형/최종학교 + `최종==합격` 판정)을 슬롯 순회로 교체:

```python
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from school_types import classify_school

SLOT_FIELDS = [
    ("영재고_접수", "영재고_결과", "영재고"),   # 유형 고정
    ("전기_접수학교", "전기_결과", None),      # 유형은 classify_school
    ("후기_접수학교", "후기_결과", None),
]

for r in tracking_rows[1:]:
    for school_col, result_col, fixed_type in SLOT_FIELDS:
        school = get(r, school_col)
        result = get(r, result_col)
        if not school or result != "최종합":
            continue
        school_type = fixed_type or classify_school(school.split(",")[0])
        # 이하 기존 passed["early"/"late"] 분류 로직에 school_type/school 사용
```

`유형→섹션` 매핑: 전기(EARLY) = 영재고·과학고·예술계고·특성화고, 후기(LATE) = 자사고·외고/국제고·일반고(비평준화 섹션). 섹션명 매칭은 기존 `section_name.split(" / ")[0]` 방식 유지. 1차/2차 셀은 결과 코드 이력이 한 칸뿐이므로 빈칸으로 두고 최종만 "합격" 기록.

- [ ] **Step 2: 실행 검증 (합격자 0명 시즌이므로 구조만)**

Run: `/home/rjegj/projects/unified_venv/bin/python generators/generate_final_sheets.py 2026`
Expected: 오류 없이 완료, `전기고 합격자: 0명` 안팎(7월 기준 최종합 없음) — 크래시 없음이 검증 목표

- [ ] **Step 3: 커밋**

```bash
git add generators/generate_final_sheets.py
git commit -m "fix: 최종 합불 시트 판정을 슬롯 결과(최종합) 기반으로 교체

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 7: 잔존 참조 정리 + 대시보드 기준 표기

**Files:**
- Modify: `integrated_dashboard.py:52-57` (`load_2026_data` 정규화), 각 탭 제목 문자열
- Modify: `auto_sync.py` (트래킹 요약이 구 컬럼 참조 시)
- Delete: `generators/create_progress_tracker.py`, `generators/cloud_function_main.py`

**Interfaces:**
- Consumes: 트래킹 새 schema 컬럼 (Task 2), 진행현황 새 헤더 (Task 5)

- [ ] **Step 1: 구 컬럼 참조 전수 조사**

Run: `grep -rn "접수유형\|접수상태\|최종유형\|최종학교\|접수학교" --include="*.py" . | grep -v test_ | grep -v migrate_to_slots`
Expected: 각 참조 지점 목록. 아래 Step 2~3에서 전부 처리하고, 처리 못 한 항목은 커밋 메시지에 명시.

- [ ] **Step 2: `load_2026_data` 정규화 교체 (`integrated_dashboard.py:52-57`)**

```python
        # 표시용 유형/지원학교: 최종배정 > 후기 > 전기 > 영재고 > 희망 (늦은 단계 우선)
        def _first_nonempty(*series_list):
            out = pd.Series([""] * len(df), index=df.index)
            for s in series_list:
                if s is None:
                    continue
                s = s.astype(str).str.strip()
                out = out.where(out != "", s)
            return out

        def _col(name):
            return df[name] if name in df.columns else None

        df["지원학교"] = _first_nonempty(
            _col("최종배정학교"), _col("후기_접수학교"), _col("전기_접수학교"), _col("영재고_접수"), _col("희망학교"), _col("지원학교")
        )
        df["유형"] = _first_nonempty(
            _col("후기_유형"), _col("전기_유형"), _col("희망유형"), _col("유형")
        )
```

(영재고_접수가 지원학교로 잡히는 행의 유형은 희망유형 fallback으로 충분 — 영재고 지원자는 희망유형도 영재고인 경우가 대부분. 미스매치는 데이터상태 `확인필요`로 이미 잡힌다.)

- [ ] **Step 3: 탭 제목에 기준 명시**

`integrated_dashboard.py`에서 `st.header(` / `st.subheader(` 호출 중 전체 현황·전기고·후기고·심층 분석 탭 제목 문자열 뒤에 기준 표기를 덧붙인다:
- 전체 현황: `"📈 전체 현황 (기준: 최종배정 > 접수 > 희망)"`
- 전기고/후기고: `"(기준: 접수 우선, 미접수는 희망)"`
- 진행 현황 계열(TAB 5/8): `"(기준: 희망 + 시기 슬롯)"`

문자열만 수정 — 로직 변경 없음.

- [ ] **Step 4: `auto_sync.py` 요약부 확인·수정**

`auto_sync.py:517-543` 부근의 트래킹 요약이 `유형`/`최종유형` 등 구 컬럼을 집계하면 `희망유형`과 `데이터상태` 기반으로 교체한다. 참조가 없으면 수정하지 않는다.

- [ ] **Step 5: 구버전 스크립트 삭제**

```bash
git rm generators/create_progress_tracker.py generators/cloud_function_main.py
```

(SHEET_DEPENDENCY.md 기준: 전자는 build_progress_from_tracking으로 대체된 구버전, 후자는 미사용.)

- [ ] **Step 6: 대시보드 기동 검증**

Run: `cd /home/rjegj/projects/Project_HighSchool_apply_Dashboard && timeout 30 /home/rjegj/projects/unified_venv/bin/python -m streamlit run integrated_dashboard.py --server.headless true --server.port 8599 & sleep 15 && curl -s http://localhost:8599 | head -5; kill %1 2>/dev/null`
Expected: HTML 응답 (Streamlit 기동 성공)

- [ ] **Step 7: 커밋**

```bash
git add integrated_dashboard.py auto_sync.py
git commit -m "feat: 대시보드 표시 기준을 슬롯 우선순위로 교체 + 탭별 기준 명시, 구버전 스크립트 삭제

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

### Task 8: 전체 파이프라인 검증 (301 → 전체)

**Files:** 없음 (실행·검증만)

- [ ] **Step 1: 301 테스트 입력**

301 시트에 테스트 행 1건 입력 (스크립트로):

```python
# 301 시트 3행(첫 학생)에: X열(전기_접수학교)="한성과학고", Y열(전기_결과)="1차합"
from sheets_client import get_client
ws = get_client().open_by_key("14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0").worksheet("301")
ws.update(values=[["한성과학고", "1차합"]], range_name="X3:Y3")
```

- [ ] **Step 2: dry-run으로 301 반영 확인**

Run: `/home/rjegj/projects/unified_venv/bin/python generators/sync_type_to_tracking.py 2026 --dry-run --apply-schema --no-legacy-fill 2>&1 | grep -A2 "301!3\|전기"`
Expected: 해당 학생의 `전기_접수학교 <- 301!3`, `전기_유형` = `과학고`, `데이터상태` = `전기진행` 갱신 후보 표시

- [ ] **Step 3: 실제 적용 + 테스트 입력 원복**

```bash
/home/rjegj/projects/unified_venv/bin/python generators/sync_type_to_tracking.py 2026 --apply-schema --no-legacy-fill
```

트래킹에서 해당 학생 행 확인 후, 301 X3:Y3와 트래킹 해당 셀을 빈칸으로 원복하고 sync 1회 재실행.

- [ ] **Step 4: auto_sync 전체 실행**

Run: `/home/rjegj/projects/unified_venv/bin/python auto_sync.py --force 2>&1 | tail -20`
Expected: `오류=0건`, 쿼터 재시도 메시지 없음, 진행현황 행 수 150+ 규모

- [ ] **Step 5: DEV_LOG 기록**

```bash
cd /home/rjegj/projects && unified_venv/bin/python System_Master/auto_logger.py --path Project_HighSchool_apply_Dashboard --task "시기 슬롯 구조 마이그레이션 적용 + 전체 파이프라인 검증 완료" --done
```

---

### Task 9: 문서 갱신

**Files:**
- Modify: `ADMISSION_STAGE_REDESIGN.md` (진행 기록 append)
- Modify: `SHEET_DEPENDENCY.md` (컬럼 표·파이프라인 갱신)

- [ ] **Step 1: ADMISSION_STAGE_REDESIGN.md 진행 기록 append**

`## 진행 기록` 섹션에 추가:

```markdown
### 2026-07-07

- 본 문서의 유형 기반 V:X 설계를 시기 슬롯 구조로 개정했다.
- 기준 문서: `docs/superpowers/specs/2026-07-07-admission-slot-redesign-design.md`
- 반별 시트 V:AA = 영재고/전기/후기 슬롯(접수학교+결과), AB 폐기.
- 결과 코드: 1차합/1차불/2차합/2차불/최종합/최종불/포기 (한 칸 덮어쓰기).
- `최종배정학교`는 진학부 전용 — sync가 쓰지 않는다.
```

- [ ] **Step 2: SHEET_DEPENDENCY.md 갱신**

`입시_트래킹` 표의 sync_type_to_tracking 행 설명을 "반별 시트(301~314) 희망 H:P + 시기 슬롯 V:AA → 희망/슬롯/데이터상태 갱신"으로, 특별전형 읽는 컬럼 표는 그대로, 파이프라인 순서도의 2단계 설명을 갱신. `create_progress_tracker.py` 행 삭제.

- [ ] **Step 3: 커밋 + 푸시**

```bash
git add ADMISSION_STAGE_REDESIGN.md SHEET_DEPENDENCY.md
git commit -m "docs: 시기 슬롯 구조 개정 반영 (ADMISSION_STAGE_REDESIGN, SHEET_DEPENDENCY)

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
git push
```

---

## 계획 외 확인사항 (실행 중 발견 시 처리)

- 담임 안내: 새 V:AA 입력 규칙(학교명 + 결과 코드)은 시트 상단 메모/데이터 검증(dropdown)으로 안내하면 좋으나 본 계획 범위 밖 — 결과 칸에 Sheets 데이터 검증 dropdown(7개 코드) 추가는 Task 4 적용 시 여유 있으면 포함 가능.
- legacy `유형`/`지원학교` 컬럼 삭제는 이번에 하지 않는다(대시보드 하위 호환·2025 비교용). 숨김 처리만 수동으로.
