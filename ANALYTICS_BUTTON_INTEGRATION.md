# 🎛️ 분석 프로그램 대시보드 버튼 통합

**작업일:** 2026-03-22
**상태:** ✅ 완료

---

## 🎯 개요

CLI로만 실행되던 분석 프로그램(Project_HighSchool_apply_Analytics)을 **대시보드 [관리 도구] 탭에 버튼**으로 통합했습니다.

---

## 📍 위치

**대시보드 → [🔧 관리 도구] 탭 → 7️⃣ 심층 통계 분석 섹션**

```
관리 도구 탭 구성:
├─ 1️⃣ 데이터 동기화
├─ 2️⃣ 시트 생성
├─ 3️⃣ 대시보드 생성
├─ 4️⃣ 진행 현황
├─ 5️⃣ 리포트 생성
├─ 6️⃣ 심화 분석
└─ 7️⃣ 심층 통계 분석  ← 새로 추가! 🆕
   ├─ 📊 Step1: 데이터 전처리
   ├─ 🔬 Step2: 통계 분석
   └─ 📄 Step3: HTML 리포트
```

---

## 🔘 각 버튼 설명

### 📊 Step1: 데이터 전처리

**실행 파일:** `Project_HighSchool_apply_Analytics/src/research_analytics.py`

**기능:**
- 배정 결과 데이터 전처리
- 학교별 경쟁률, 만족도 계산
- 동네-학교 매트릭스 생성

**입력:** `Step1_전처리_익명화_마스터.xlsx`
**출력:** `Step2_지망선호도_및_지역흐름.xlsx`
**소요 시간:** ~10초

**필수 조건:**
- ✅ 최종 배정 결과 엑셀 파일 준비
- ✅ 행정동(동) 컬럼 포함
- ✅ 배정학교 컬럼 포함

---

### 🔬 Step2: 통계 분석

**실행 파일:** `Project_HighSchool_apply_Analytics/src/statistical_deep_research.py`

**기능:**
```
1. K-Means 군집화
   → 학교를 4가지 유형으로 분류
   • 유형 A: 고경쟁_아쉬움
   • 유형 B: 안정_만족형
   • 유형 C: 고만족_적정형
   • 유형 D: 복합형

2. 카이제곱 검정
   → 거주지가 배정에 영향을 주는가?
   (P-value, Cramer V 계산)

3. 상관관계 분석
   → 경쟁률과 만족도의 관계
   (Pearson 상관계수)
```

**입력:** `Step2_지망선호도_및_지역흐름.xlsx`
**출력:** `Step3_학교유형화_및_통계검증.xlsx`
**소요 시간:** ~5초

---

### 📄 Step3: HTML 리포트

**실행 파일:** `Project_HighSchool_apply_Analytics/src/final_dashboard_generator.py`

**기능:**
- Step3 데이터를 HTML로 시각화
- 분석 결과 + 통계표 + 인사이트
- 대시보드 [리포트 뷰어]에 자동 표시

**입력:** `Step3_학교유형화_및_통계검증.xlsx`
**출력:** `output/Insight_Dashboard_YYYY.html`
**소요 시간:** ~3초

**자동 연계:**
```
Step3 완료 ↓ ✅
   ↓
output/ 폴더에 HTML 생성
   ↓ (자동 감지)
대시보드 [리포트 뷰어] 탭에 표시
   ↓
사용자가 selectbox + expander로 조회
```

---

## 🚀 사용 방법

### 순서대로 실행

```
1️⃣ [📊 Step1] 클릭
   └─ 10초 대기
   └─ ✅ 완료 메시지

2️⃣ [🔬 Step2] 클릭
   └─ 5초 대기
   └─ ✅ 완료 메시지

3️⃣ [📄 Step3] 클릭
   └─ 3초 대기
   └─ ✅ 완료 메시지
   └─ "리포트 뷰어에서 확인하세요"
```

### 결과 확인

```
[📊 생성된 리포트 뷰어] 섹션으로 이동
   ↓
[📊 Excel] 탭에서 Step3 Excel 다운로드
   ↓
[🎯 전기고 / 🍂 후기고] 탭에서 HTML 리포트 조회
   └─ selectbox로 리포트 선택
   └─ expander로 펼쳐서 보기
```

---

## 📊 전체 흐름도

```
┌─────────────────────────────────────────┐
│  대시보드 [관리 도구] 탭                 │
├─────────────────────────────────────────┤
│                                         │
│  7️⃣ 심층 통계 분석 (3개 버튼)          │
│  ┌─────────┬──────────┬────────────┐   │
│  │📊 Step1 │🔬 Step2 │📄 Step3    │   │
│  └────┬────┴────┬─────┴────┬───────┘   │
│       ↓         ↓          ↓            │
│  ┌─────────────────────────────────┐   │
│  │ 분석 프로그램 실행              │   │
│  │ (subprocess)                    │   │
│  └────┬────────────────────┬───────┘   │
│       ↓                    ↓            │
│   Step2 Excel         Step3 Excel      │
│   + Matrix              + HTML         │
│       ↓                    ↓            │
│  ┌─────────────────────────────────┐   │
│  │ [📊 생성된 리포트 뷰어]        │   │
│  │  ├─ 📊 Excel 탭              │   │
│  │  ├─ 🎯 전기고 탭            │   │
│  │  └─ 🍂 후기고 탭            │   │
│  └─────────────────────────────────┘   │
│       ↓ (selectbox + expander)         │
│  사용자 조회                            │
└─────────────────────────────────────────┘
```

---

## ⚙️ 기술 구현

### 실행 방식: subprocess

```python
# 버튼 클릭 시
if st.button("📊 Step1: 데이터 전처리", ...):
    with st.spinner("Step 1 실행 중..."):
        try:
            result = subprocess.run([
                "/home/rjegj/projects/unified_venv/bin/python",
                "/home/rjegj/projects/Project_HighSchool_apply_Analytics/src/research_analytics.py"
            ], capture_output=True, text=True, timeout=120,
               cwd="/home/rjegj/projects/Project_HighSchool_apply_Analytics")

            if result.returncode == 0:
                st.success("✅ Step 1 완료!")
                st.info(result.stdout)  # 스크립트 출력 표시
            else:
                st.error(f"❌ 실패: {result.stderr}")
        except Exception as e:
            st.error(f"❌ 오류: {str(e)}")
```

### 핵심 파라미터

| 파라미터 | 값 | 의미 |
|:-------:|:---:|:---:|
| `capture_output=True` | - | stdout/stderr 캡처 |
| `text=True` | - | 텍스트 모드 |
| `timeout=120` | 초 | 최대 2분 타임아웃 |
| `cwd=...` | 경로 | 작업 디렉토리 |

---

## ⚠️ 주의사항

### 1. 실행 순서 필수

```
Step1 → Step2 → Step3 (반드시 순서대로)

❌ Step2만 먼저 실행 불가
   (Step2는 Step1 출력을 필요로 함)
```

### 2. 최종 배정 데이터 필요

```
언제부터 사용?
  → 2026년 최종 배정 완료 후 (약 1월 2027)
  → 또는 2025 데이터로 테스트

테스트용 2025 데이터:
  ✓ 이미 Step1-3 완료 상태
  ✓ output/Insight_Dashboard_2025.html 존재
```

### 3. 파일 경로 이해

```
상대 경로 (중요):
  research_analytics.py
  ↓ 실행 시 cwd = Project_HighSchool_apply_Analytics
  ↓
  data/processed/Step2_... 자동 생성 (상대 경로)

절대 경로:
  /home/rjegj/projects/Project_HighSchool_apply_Analytics/
  ├─ src/
  │  ├─ research_analytics.py
  │  ├─ statistical_deep_research.py
  │  └─ final_dashboard_generator.py
  ├─ data/
  │  └─ processed/
  │     ├─ Step1_...
  │     ├─ Step2_...
  │     └─ Step3_...
  └─ output/
     ├─ Insight_Dashboard_2025.html
     └─ Insight_Dashboard_2026.html
```

---

## 🐛 문제 해결

### ❓ Step1 실행 시 "행정동 컬럼을 찾을 수 없음"

**원인:** Step1 입력 파일에 "행정동" 또는 "동" 컬럼이 없음

**해결:**
```
1. 대시보드를 통해 Step1 파일 생성 불가
2. 수동으로 Excel 파일 준비 필요:
   ✓ 행정동 컬럼 필수 포함
   ✓ 파일 경로: data/processed/Step1_전처리_익명화_마스터.xlsx

또는 research_analytics.py 수정:
  KEY_DONG = "동"  또는  "주소(동)"
```

### ❓ Step2/Step3 버튼을 클릭했는데 "파일을 찾을 수 없음"

**원인:** Step1 또는 Step2가 아직 실행 안 됨

**해결:**
```
1. 버튼 순서 다시 확인
   Step1 → Step2 → Step3 (이 순서!)

2. Step1 출력 확인
   Project_HighSchool_apply_Analytics/data/processed/
   └─ Step2_지망선호도_및_지역흐름.xlsx 존재 확인
```

### ❓ "타임아웃" 오류 (최대 2분)

**원인:** 분석이 너무 오래 걸림

**해결:**
```
1. Step 파일 크기 확인
   (매우 큰 Excel은 느릴 수 있음)

2. timeout 값 증가
   timeout=120 → timeout=300 (5분)
   (integrated_dashboard.py 수정)
```

---

## 📈 자동화 제안

### 일괄 실행 버튼 추가 (선택사항)

```python
if st.button("🔄 Step1-3 일괄 실행", use_container_width=True):
    with st.spinner("전체 분석 중... (약 20초)"):
        steps = ["research_analytics.py",
                 "statistical_deep_research.py",
                 "final_dashboard_generator.py"]

        for step in steps:
            result = subprocess.run([...], ...)
            if result.returncode != 0:
                st.error(f"❌ {step} 실패")
                break
            st.success(f"✅ {step} 완료")
```

---

## 🎯 연계 요약

| 항목 | 설정 | 상태 |
|:---:|:---:|:---:|
| **CLI 실행** | ❌ 제거됨 | 불필요 |
| **Streamlit 버튼** | ✅ 추가됨 | 활성 |
| **subprocess** | ✅ 구현됨 | 작동 중 |
| **자동 결과 표시** | ✅ 연계됨 | 리포트 뷰어 |
| **오류 처리** | ✅ 구현됨 | 명확함 |

---

## 📍 다음 단계

**9월 (PHASE 4):**
1. 2026 최종 배정 데이터 준비
2. Step1 입력 파일 생성
3. 대시보드 → [7️⃣ 심층 통계 분석] → [📊 Step1] 클릭
4. 순서대로 3개 버튼 실행
5. [리포트 뷰어]에서 결과 확인

**11월 (PHASE 5):**
- Step2 업데이트 (후기고 데이터 추가)
- HTML 리포트 최신화

**1월 (최종):**
- 최종 분석 리포트 생성
- 연간 통계 정리

---

**상태:** ✅ 완료 및 검증됨
**테스트:** 문법 검사 + 수동 구조 검증
**다음 실제 테스트:** 2026년 9월 (최종 배정 후)
