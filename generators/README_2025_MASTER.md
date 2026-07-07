# 2025_입시결과_마스터 시트 — 사용 가이드

## 📌 개요

2025학년도 입시 데이터를 "결과 우선(results-first)" 구조로 전처리하여 Google Sheets에 마스터 시트를 생성합니다.

**마스터 시트**: `2025_입시결과_마스터` (SS2 Spreadsheet)

---

## 🚀 빠른 시작

### 1. 마스터 시트 생성 (처음 1회)
```bash
cd /home/rjegj/projects/Project_HighSchool_apply_Dashboard

# 마스터 시트 생성
python generators/build_2025_master_results.py

# 검증
python generators/validate_2025_master.py

# 대시보드 테스트
python generators/test_2025_dashboard_load.py
```

### 2. 생성 후 결과 확인
- Google Sheets 열기: https://docs.google.com/spreadsheets/d/1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI
- `2025_입시결과_마스터` 시트에서 415명 데이터 확인

---

## 📊 마스터 시트 구조

### 컬럼 (8개)

| # | 컬럼명 | 타입 | 설명 |
|---|---|---|---|
| 1 | 반 | 문자 | 학년-반 (예: 1, 2, 3 등) |
| 2 | 번호 | 문자 | 학번 |
| 3 | 성명 | 문자 | 학생명 |
| 4 | 성별 | 문자 | 남/여 |
| 5 | 최종유형 | 문자 | 과학고/예고/특성화고/자사고/외고·국제고/비평준화고/일반고 |
| 6 | 배정학교 | 문자 | 학교명 |
| 7 | 합격여부 | 문자 | 합격/불합격/미결 |
| 8 | 데이터출처 | 문자 | 전기고합불시트/후기고합불시트/일반고xlsx/기본정보만 |

### 데이터 특성

- **총 학생**: 415명
- **결과 있는 학생**: 409명 (98.6%)
- **결과 없는 학생**: 6명 (기본정보만)

### 데이터 분포

**유형별**:
```
일반고:           289명 (69.6%)  ← 가장 많음
자사고:            55명 (13.2%)
외고·국제고:      23명 (5.5%)
특성화고:         23명 (5.5%)
예고:             11명 (2.6%)
과학고:            6명 (1.4%)
비평준화고:        2명 (0.5%)
미분류:            6명 (1.4%)
```

**합격 상황**:
```
합격:    392명 (94.5%)
불합격:  16명 (3.9%)
미결:     7명 (1.7%)
```

---

## 🔧 사용 시나리오

### 시나리오 1: 처음 생성
```bash
# 모든 원본 데이터에서 마스터 시트 생성
python generators/build_2025_master_results.py
```

**결과 샘플**:
```
================================================================================
✓ 마스터 시트 생성 완료
================================================================================
총 415명
  - 전기고 합격: 40명
  - 후기고 합격: 80명
  - 일반고 배정: 289명
  - 기본정보만: 6명

합격자 총: 392명
불합격: 16명
미결정: 7명
```

### 시나리오 2: 데이터 재검증
```bash
# 마스터 시트 검증
python generators/validate_2025_master.py
```

**출력**:
- 시트 로드 확인
- 필수 컬럼 완결성
- 성명 결측치 확인
- 유형/합격/출처 분포

### 시나리오 3: 대시보드 호환성 확인
```bash
# 대시보드에서 로드 가능한지 테스트
python generators/test_2025_dashboard_load.py
```

**검증 항목**:
- 데이터 로드 성공
- 모든 필수 컬럼 존재
- 데이터 샘플 출력
- 통계 계산

---

## 📝 원본 데이터 소스

### Stage 1: 최종 결과 파싱

| 원본 | 위치 | 파싱 로직 | 추출 내용 |
|---|---|---|---|
| **전기고 합불** | SS2 시트 | `parse_early_high_final()` | 과학고/예고/특성화고 3섹션 |
| **후기고 합불** | SS2 시트 | `parse_late_high_final()` | 자사고/외고/비평준화고 3섹션 |
| **일반고 배정** | xlsx 파일 | `parse_general_high_xlsx()` | 배정학교/합격여부 |

### Stage 2: 학생 기본정보 로드

| 데이터 | 위치 | 로드 함수 |
|---|---|---|
| 반/번호/성명/성별 | SS2 "진학희망_" 시트들 | `load_student_base_info()` |

### 병합

성명 기준으로 정확히 매칭하여 마스터 시트 생성

---

## ⚠️ 주의사항

### 1. 성명 정규화
- 파싱 전: `.strip()` 적용 (좌우 공백 제거)
- 병합 시: 정확한 문자열 매칭

### 2. 컬럼 오프셋
- 전기고 시트: 3개 섹션이 **병렬** 배치
  - 과학고: cols 0-4
  - 예고: cols 5-9
  - 특성화고: cols 10-15
- 후기고 시트: 3개 섹션이 **병렬** 배치
  - 자사고: cols 0-5
  - 외고: cols 6-11
  - 비평준화고: cols 12-17

### 3. 데이터 타입
- 합격여부: '합격' (O), '합' (X) → 정규화됨
- 불합격: '불합격' (O), '불' (X) → 정규화됨

---

## 🛠️ 트러블슈팅

### 문제: "2025_입시결과_마스터" 시트 없음
```bash
# 마스터 시트 재생성
python generators/build_2025_master_results.py
```

### 문제: 일부 학생 누락
```bash
# 데이터 검증 (결측 확인)
python generators/validate_2025_master.py

# 샘플 확인
python generators/test_2025_dashboard_load.py
```

### 문제: 대시보드에서 로드 실패
```bash
# integrated_dashboard.py에서 load_2025_data() 호환성 확인
python generators/test_2025_dashboard_load.py

# 문제 없으면 Streamlit 캐시 초기화
streamlit run integrated_dashboard.py --logger.level=debug
```

---

## 📚 관련 파일

```
generators/
├── inspect_2025_sources.py         # 원본 데이터 분석
├── build_2025_master_results.py    # ⭐ 마스터 시트 생성 (핵심)
├── validate_2025_master.py         # 검증
├── test_2025_dashboard_load.py     # 호환성 테스트
└── README_2025_MASTER.md           # 이 문서

integrated_dashboard.py             # 대시보드 (load_2025_data() 개선)
IMPLEMENTATION_SUMMARY.md           # 상세 구현 리포트
```

---

## 📞 더 알아보기

- **구현 상세**: `IMPLEMENTATION_SUMMARY.md` 참고
- **원본 검사**: `generators/inspect_2025_sources.py` 실행
- **코드 주석**: 각 파일의 docstring 참고

---

**마지막 업데이트**: 2026-03-22
**상태**: ✅ 완료 및 검증 완료
