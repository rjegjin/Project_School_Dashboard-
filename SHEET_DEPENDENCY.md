# 시트 의존성 맵 (Sheet Dependency Map)

> 작성: 2026-04-12 / 업데이트: 2026-04-13  
> 목적: 어떤 .py 파일이 어떤 Google Sheets 시트를 읽거나 쓰는지 명확히 기록  
> 스프레드시트 ID (2026): `14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0`

---

## 1. 시트별 의존 파일 목록

### `입시_트래킹`
| 파일 | 읽기/쓰기 | 역할 |
|------|-----------|------|
| `generators/sync_type_to_tracking.py` | **쓰기** | 반별 시트(301~314) → 유형, 지원학교 자동 감지 후 입력 (빈 셀만) |
| `generators/generate_final_sheets.py` | 읽기 | 최종=합격인 학생 추출 → 전기고_최종/후기고_최종 생성 |
| `generators/build_progress_from_tracking.py` | 읽기 | 유형 있는 학생 전원 → 입시 진행 현황 생성 |
| `generators/create_progress_tracker.py` | 읽기 | 입시 진행 현황 생성 (build_progress_from_tracking의 구버전) |
| `generators/init_2026_spreadsheet.py` | **쓰기** | 시트 최초 생성 시 헤더 + 샘플 데이터 입력 (아래 주의 참조) |
| `generators/cloud_function_main.py` | 읽기 | Cloud Function 트리거용 (미사용) |
| `integrated_dashboard.py` | 읽기 | 대시보드 전체 현황/전기고/후기고/반별/필터링 탭 |
| `auto_sync.py` | 읽기 | 스냅샷 비교용 집계 (쓰기 없음) |

---

### `전기고_최종`
| 파일 | 읽기/쓰기 | 역할 |
|------|-----------|------|
| `generators/generate_final_sheets.py` | **쓰기 (전체 초기화)** | 입시_트래킹 최종=합격 중 전기고 유형 → 시트 재생성 |
| `generators/init_2026_spreadsheet.py` | **쓰기** | 최초 시트 생성 (샘플 데이터 포함, 아래 주의 참조) |
| `integrated_dashboard.py` | 읽기 (간접) | `load_2025_final_results()`에서 2025 시트를 읽는 방식과 유사하게 참조 |
| `refine_main_sheet.py` | 읽기 | 시트 존재 여부 확인용 |

---

### `후기고_최종`
| 파일 | 읽기/쓰기 | 역할 |
|------|-----------|------|
| `generators/generate_final_sheets.py` | **쓰기 (전체 초기화)** | 입시_트래킹 최종=합격 중 후기고 유형 → 시트 재생성 |
| `generators/init_2026_spreadsheet.py` | **쓰기** | 최초 시트 생성 (샘플 데이터 포함, 아래 주의 참조) |
| `refine_main_sheet.py` | 읽기 | 시트 존재 여부 확인용 |

---

### `입시 진행 현황`
| 파일 | 읽기/쓰기 | 역할 |
|------|-----------|------|
| `generators/build_progress_from_tracking.py` | **쓰기 (전체 초기화)** | 입시_트래킹에서 유형 있는 전체 학생 → 진행현황 재생성 |
| `generators/create_progress_tracker.py` | **쓰기 (전체 초기화)** | 구버전 (build_progress_from_tracking으로 대체됨) |
| `integrated_dashboard.py` | 읽기 | TAB 5, TAB 8에서 시각화 |
| `auto_sync.py` | 실행 (간접) | `build_progress_from_tracking.py` 호출 |

---

---

### `특별전형_트래킹`
| 파일 | 읽기/쓰기 | 역할 |
|------|-----------|------|
| `generators/sync_special_to_tracking.py` | **쓰기 (전체 초기화)** | 반별 시트(301~314) 특별전형 컬럼 → 전체 학생 O/공백 기록 |
| `generators/generate_special_report.py` | 읽기 | 특별전형_트래킹 + 입시_트래킹 → `reports/특별전형_현황.html` |
| `integrated_dashboard.py` | 읽기 (간접) | 관리 탭에서 동기화/HTML 생성 버튼 실행 |
| `auto_sync.py` | 실행 (간접) | `sync_special_to_tracking.py` 호출 (3단계) |

**읽는 컬럼 (반별 시트 0-based):**
| col | 컬럼명 |
|-----|--------|
| 4 (E) | 사회통합전형 |
| 5 (F) | 특례 |
| 6 (G) | 보훈 |
| 16 (Q) | 쌍둥이 |
| 17 (R) | 학폭 |
| 18 (S) | 교직원자녀 |
| 19 (T) | 장애 |
| 20 (U) | 다자녀(3인+) ← `add_dajanyeo_column.py`로 추가 (2026-04-13) |

---

## 2. 자동화 파이프라인 실행 순서

```
[cron 07:30 평일]
auto_sync.py
  ├── 1단계: sync_form_to_class_sheets.py    (설문지 → 반별 시트 301~314)
  ├── 2단계: sync_type_to_tracking.py         (반별 시트 → 입시_트래킹 유형/지원학교)
  ├── 3단계: sync_special_to_tracking.py      (반별 시트 → 특별전형_트래킹)  ← NEW
  ├── 4단계: build_progress_from_tracking.py  (입시_트래킹 → 입시 진행 현황)
  └── 5단계: generate_dashboard.py            (입시_트래킹 → reports/ HTML 일반 3개)
           + generate_special_report.py       (특별전형_트래킹 → reports/특별전형_현황.html)  ← NEW

[수동, 최초 1회]
add_dajanyeo_column.py 2026                  (반별 시트 전체에 다자녀(3인+) col 20 삽입 — 멱등)

[수동, 합불 발표 후]
generate_final_sheets.py                     (입시_트래킹 최종 → 전기고_최종 / 후기고_최종)
```

---

### `reports/` (정적 HTML 출력)
| 파일 | 생성 스크립트 | 내용 |
|------|--------------|------|
| `전기고_현황.html` | `generators/generate_dashboard.py` | 영재고/과학고/예술계고/특성화고 카드 |
| `후기고_현황.html` | `generators/generate_dashboard.py` | 자사고/외고/비평준화고/일반고 카드 |
| `전체_현황.html`  | `generators/generate_dashboard.py` | 전체 (유형 필터 버튼 포함) |
| `특별전형_현황.html` | `generators/generate_special_report.py` | 특별전형 8종 요약카드 + 전형×계열 매트릭스 + 학생 필터 |

> `reports/_legacy/` — 구형 `목일중_*` 파일 보관 (Streamlit 뷰어에서 자동 제외)

---

## 3. ⚠️ 샘플 데이터 주의 (2026-04-12 발견)

### 문제 경위

`init_2026_spreadsheet.py`가 스프레드시트를 최초 생성할 때  
`입시_트래킹`의 **앞 10행(반=1, 번호=1~10)** 에 샘플 학생 데이터를 삽입했음.

이 샘플 데이터가 `유형`과 `최종=합격` 컬럼을 채운 상태로 남아 있어  
→ `generate_final_sheets.py` 실행 시 `전기고_최종`, `후기고_최종` 시트가  
→ `build_progress_from_tracking.py` 실행 시 `입시 진행 현황`이  
**실제 데이터 없이 샘플만으로 채워진 상태**가 됨.

### 샘플 vs 실제 판별 기준

| 구분 | 반 컬럼 값 | 유형 근거 | 최종 컬럼 |
|------|-----------|-----------|-----------|
| **샘플** | 1 (번호 1~10, 이름: 강은서~손정인) | 반별 시트(301)에 해당 유형 없음 | `합격` (가짜) |
| **실제** | 1~14 (sync_type 실행 결과) | 반별 시트(301~314)에서 감지된 값 | 비어있음 (미발표) |

> 반 번호 체계: 입시_트래킹은 `1`~`14` (301반 → 1, 314반 → 14)  
> 실제 데이터: 유형 180명, 지원학교 98명 (2026-04-12 기준)

### 조치 (2026-04-12)

**입시_트래킹** 앞 10행(강은서~손정인)의 `유형` + `최종` 컬럼 → **빈칸으로 클리어**  
**전기고_최종**, **후기고_최종** → 헤더/섹션 구조 유지, 데이터 행 클리어  
**입시 진행 현황** → 전체 클리어 후 실제 데이터로 재생성

---

## 4. 향후 실행 가이드

```bash
# 유형 자동 동기화 (반별 시트 → 입시_트래킹)
python generators/sync_type_to_tracking.py 2026

# 특별전형 동기화 (반별 시트 → 특별전형_트래킹)
python generators/sync_special_to_tracking.py 2026

# 특별전형 HTML 생성 (특별전형_트래킹 → reports/특별전형_현황.html)
python generators/generate_special_report.py 2026

# 다자녀 컬럼 삽입 (최초 1회, 멱등)
python generators/add_dajanyeo_column.py 2026

# 진행현황 재생성 (유형 입력 후)
python generators/build_progress_from_tracking.py 2026

# 합불 발표 후 → 최종 합불 시트 생성
python generators/generate_final_sheets.py 2026

# 전체 파이프라인 (cron 동일)
python auto_sync.py --force
```
