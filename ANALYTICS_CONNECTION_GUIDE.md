# 🔗 통계 분석 프로그램 연계 완벽 가이드

**최종 업데이트:** 2026-03-22
**대상:** 대시보드 팀 + 분석 팀

---

## 🎯 개요

**Project_HighSchool_apply_Analytics**는 고입 배정 결과의 심층 통계 분석 도구입니다. 대시보드와 다음과 같이 연계됩니다:

```
대시보드 (실시간 추적)
    ↓ [최종 배정 완료]
분석 프로그램 (사후 분석)
    ↓ [분석 결과 생성]
대시보드 [심층분석] 탭 (결과 표시)
```

---

## 📊 분석 파이프라인 상세

### Step 1️⃣: 데이터 전처리 (research_analytics.py)

**목적:** 배정 결과를 분석 가능한 형태로 변환

**입력:**
```
data/processed/Step1_전처리_익명화_마스터.xlsx
├─ 행정동 (학생 거주 지역)
├─ 배정학교 (최종 배정된 학교)
├─ 1지망 (1순위 희망학교)
├─ 2지망 (2순위 희망학교)
└─ ... 기타 정보
```

**처리 과정:**
```python
1. 데이터 로드 및 정제 (공백 제거)

2. 학교별 분석
   • 실제배정인원 (배정된 학생 수)
   • 일지망_배정된_사람 (1순위로 배정된 학생)
   • 총_1지망_지원자수 (1순위 지원자 총수)
   • 실질경쟁률 = 지원자 / 배정자
   • 배정만족도(%) = 1순위배정 / 배정자

3. 동네(행정동)별 분석
   • 거주학생수
   • 1지망_성공률(%) (1순위로 배정된 비율)

4. 동네-학교 매트릭스 생성
   • 행: 행정동 (거주 지역)
   • 열: 배정학교
   • 값: 배정된 학생 수
```

**출력:**
```
data/processed/Step2_지망선호도_및_지역흐름.xlsx
├─ Sheet: 연구1_학교별_인기도
│  ├─ 배정고등학교
│  ├─ 실제배정인원
│  ├─ 실질경쟁률
│  └─ 배정만족도(%)
├─ Sheet: 동네별_배정_현황
│  └─ 1지망 성공률 등
└─ Sheet: 부록_동네_학교_전체매트릭스
   └─ 숫자 매트릭스 (시각화용)
```

**실행:**
```bash
cd Project_HighSchool_apply_Analytics
python src/research_analytics.py
```

---

### Step 2️⃣: 심층 통계 분석 (statistical_deep_research.py)

**목적:** Step2 데이터에 고급 통계 적용

**입력:**
```
data/processed/Step2_지망선호도_및_지역흐름.xlsx
├─ 연구1_학교별_인기도 (경쟁률, 만족도)
└─ 부록_동네_학교_전체매트릭스 (흐름)
```

**분석 수행:**

#### 1. 학교 유형화 (K-Means Clustering)
```
입력 지표:
  • 실질경쟁률
  • 배정만족도(%)

분류 결과:
  유형 A: 고경쟁_아쉬움 (경쟁률 높음, 만족도 낮음)
         → 인기 학교인데 1순위 탈락자가 많음 (2순위/임의배정)

  유형 B: 안정_만족형 (경쟁률 낮음, 만족도 높음)
         → 지역 기반 학교 (안정적, 선호)

  유형 C: 고만족_적정형 (만족도 90% 이상)
         → 선호도 높음, 안정적

  유형 D: 복합형 (평균적)
         → 중간 수준

출력: 각 학교에 유형 라벨 부여
```

#### 2. 거주지 영향도 검정 (Chi-Square Test)
```
귀무가설: 거주지(동)와 배정학교는 무관
대립가설: 거주지가 배정학교에 영향을 줌

지표:
  • Chi-square 값
  • P-value (< 0.05면 통계적 유의)
  • Cramer's V (연관성 강도)

해석:
  P < 0.05  → "거주지가 배정에 영향을 줌" (유의미)
  P >= 0.05 → "거주지와 무관" (우연일 가능성)
```

#### 3. 경쟁률-만족도 상관관계 (Pearson Correlation)
```
상관계수:
  r < -0.5  → 강한 음의 상관
            (경쟁률 높으면 만족도 낮음)

  -0.5 ~ 0  → 약한 상관

해석:
  인기 학교는 경쟁이 심해서
  1순위 탈락자가 많아짐 → 만족도 감소
```

**출력:**
```
data/processed/Step3_학교유형화_및_통계검증.xlsx
├─ Sheet: 1_유형화(신뢰데이터)
│  ├─ 배정고등학교
│  ├─ 실질경쟁률
│  ├─ 배정만족도(%)
│  ├─ 군집_Label (클러스터 번호)
│  └─ 분석_학교유형 (유형 A/B/C/D)
├─ Sheet: 1_군집요약
│  └─ 유형별 평균 경쟁률/만족도
├─ Sheet: 2_종속성검정_결과
│  └─ P-value, Cramer V 등
└─ Sheet: 3_상관관계_결과
   └─ 상관계수, P-value 등
```

**실행:**
```bash
python src/statistical_deep_research.py
```

**소요 시간:** ~5초 (분석은 가벼움)

---

### Step 3️⃣: HTML 리포트 생성 (final_dashboard_generator.py)

**목적:** Step3 데이터를 시각화된 HTML로 변환

**입력:**
```
data/processed/Step3_학교유형화_및_통계검증.xlsx
```

**처리:**
```
1. 각 시트를 pandas DataFrame으로 읽기
2. 통계표 → HTML 테이블
3. 스타일 추가 (CSS)
4. 인사이트 메시지 삽입
5. HTML 파일 저장
```

**출력:**
```
output/Insight_Dashboard_2025.html (약 50KB)
├─ 유형화 분석 테이블
├─ 통계 검정 결과
├─ 상관관계 분석
└─ 핵심 인사이트 박스
```

**실행:**
```bash
python src/final_dashboard_generator.py
```

---

## 🔄 현재 상황 (2026-03-22)

### ✅ 완료됨
```
2025 데이터 기준으로 전체 파이프라인 완성
├─ Step1 ✅ (research_analytics.py)
├─ Step2 ✅ (statistical_deep_research.py)
├─ Step3 ✅ (final_dashboard_generator.py)
└─ HTML ✅ (output/Insight_Dashboard_2025.html 생성 완료)
```

### 📊 2025 분석 결과 예시
```
분석 대상: 후기고 배정 (약 200~300명)

주요 발견:
  • 학교 유형화: 3~4개 그룹으로 분류
  • 거주지 영향: 유의미함 (P < 0.05)
    → 동네에 따라 배정학교 분포가 다름
  • 경쟁률-만족도: 약한 음의 상관
    → 인기 학교는 경쟁이 심함
```

### ⏳ 2026 데이터로 확장 가능?

**현재 상황:**
```
2026 데이터
├─ 대시보드: 실시간 입력 중 (4월부터 본격)
├─ 배정 결과: 아직 없음 (최종 배정 = 1월)
└─ Step1 Input 파일: 아직 준비 안 됨
```

**확장 가능성:**
```
✅ Step1 파일 준비 완료 후
   (= 최종 배정 완료 후, 약 1월 2027)

├─ research_analytics.py 실행
│  └─ 2026 Step2 생성
├─ statistical_deep_research.py 실행
│  └─ 2026 Step3 생성
├─ final_dashboard_generator.py 실행
│  └─ 2026 HTML 리포트 생성
└─ 대시보드 [심층분석] 탭에 임베딩
```

---

## 📁 데이터 준비 체크리스트

### Step1 입력 파일 준비 (최종 배정 후)

```bash
# 필요한 데이터
데이터 원본: /다운로드/2026학년도 후기고 최종결과.xlsx
             또는 Google Sheets에서 export

필수 컬럼:
  ✓ 행정동 (학생 거주 지역 - 필수!)
  ✓ 배정학교 (최종 배정된 학교)
  ✓ 1지망 (1순위 희망학교)
  ✓ 2지망 (2순위 희망학교)
  ✓ 학생명 (비식별화 필요 - 번호로 변환)
  ✓ 성별 (선택)

Step1 파일 생성:
  1. 원본 데이터 export
  2. 학생명 → 학번으로 비식별화
  3. 컬럼명 정규화 (행정동, 배정학교 등)
  4. data/processed/Step1_전처리_익명화_마스터.xlsx로 저장
```

### 필수 정보: 행정동 (동)

⚠️ **매우 중요!** research_analytics.py는 "행정동" 컬럼을 반드시 필요로 합니다.

```
예시:
  학생 A: 서울시 강남구 역삼동
  학생 B: 서울시 강남구 삼성동
  학생 C: 서울시 서초구 방배동

→ 행정동: "역삼동", "삼성동", "방배동" (3가지)
→ Step2 분석 시: 동네별 배정 현황 분석
```

---

## 🔗 대시보드 통합 방식

### 옵션 A: 경량 HTML 임베딩 ✨ (권장)

**시기:** 2026년 9월 (PHASE 4)

**구현:**
```python
# 대시보드 새 탭 추가 (또는 관리 도구 확장)
with tab_analytics:
    st.header("📊 심층 분석 보고서")

    # HTML 파일 자동 감지
    output_dir = "../Project_HighSchool_apply_Analytics/output/"
    html_files = [f for f in os.listdir(output_dir)
                  if f.endswith('.html')]

    # 사용자가 선택
    selected = st.selectbox("보고서 선택", html_files)

    # 임베딩 (접기/펼치기)
    with st.expander(f"📄 {selected} (클릭 시 보기)", expanded=False):
        with open(os.path.join(output_dir, selected), 'r') as f:
            html_content = f.read()
        st.components.v1.html(html_content, height=800)
```

**작동 순서:**
```
1. 분석 프로그램 실행 (독립적)
   python src/research_analytics.py
   python src/statistical_deep_research.py
   python src/final_dashboard_generator.py

2. output/ 디렉토리에 HTML 생성됨

3. 대시보드가 자동으로 감지 & 표시
   (새로고침시 최신 버전)

4. 사용자: selectbox + expander로 조회
```

**장점:**
- ⚡ 구현 간단
- 🔄 독립 실행 가능
- 📊 실시간 업데이트 (새로고침)
- 🎯 단기 완성 가능

---

### 옵션 B: 데이터 연동 ✨ (중기)

**시기:** 2026년 11월 (PHASE 5)

**구현:**
```python
@st.cache_data(ttl=3600)
def load_analysis_excel():
    """Step3 Excel 로드"""
    path = "../Project_HighSchool_apply_Analytics/data/processed/"
    file = "Step3_학교유형화_및_통계검증.xlsx"

    dfs = {}
    for sheet in ['1_유형화(신뢰데이터)', '1_군집요약', '2_종속성검정_결과']:
        try:
            df = pd.read_excel(os.path.join(path, file), sheet_name=sheet)
            dfs[sheet] = df
        except:
            pass
    return dfs

# 탭에서 사용
with tab_analytics:
    results = load_analysis_excel()

    if '1_유형화(신뢰데이터)' in results:
        df = results['1_유형화(신뢰데이터)']

        # 필터링
        col1, col2 = st.columns(2)
        with col1:
            selected_type = st.multiselect("학교 유형",
                                          df['분석_학교유형'].unique())

        # 표시
        if selected_type:
            filtered = df[df['분석_학교유형'].isin(selected_type)]
            st.dataframe(filtered, use_container_width=True)

            # 시각화
            st.bar_chart(filtered.set_index('배정고등학교')['배정만족도(%)'])
```

**장점:**
- 🎯 동적 필터링
- 📊 대시보드 내 시각화
- ⚡ 캐싱으로 성능 최적화

**단점:**
- 복잡도 증가
- 추가 개발 필요

---

### 옵션 C: 완전 통합 (장기, 비추천)

```
분석 로직을 Streamlit 내에 포함
(PCA, GMM, 통계 검정 실시간 계산)

❌ 너무 무거움, 복잡함
→ 옵션 A/B 충분
```

---

## 📋 연계 일정

| 시기 | 단계 | 작업 | 상태 |
|:---:|:---:|:---:|:---:|
| **3월** | — | 대시보드 기본 완성 | ✅ |
| **4~6월** | — | 담임 데이터 입력 | 🔄 |
| **7~8월** | — | 전기고 결과 입력 | 📅 |
| **9월** | PHASE 4 | 2026 최종 배정 완료 | 📅 |
| **9월 중** | — | Step1 파일 생성 | 📅 |
| **9월 말** | 옵션 A | Step1-3 실행 | 📅 |
| **10월** | 옵션 A | 대시보드에 임베딩 | 📅 |
| **11월** | PHASE 5 | Step2 업데이트 (후기고) | 📅 |
| **11월 중** | 옵션 B | 프로토타입 개발 | 📅 |
| **12월** | 옵션 B | 배포 & 개선 | 📅 |
| **1월** | — | 최종 분석 | 📅 |

---

## 🎯 실행 명령어

### Quick Start

```bash
cd /home/rjegj/projects/Project_HighSchool_apply_Analytics

# Step 1: 데이터 전처리
python src/research_analytics.py

# Step 2: 통계 분석
python src/statistical_deep_research.py

# Step 3: HTML 리포트 생성
python src/final_dashboard_generator.py

# 결과 확인
ls -lah output/*.html
```

### 자동화 스크립트 (제안)

```bash
#!/bin/bash
# run_analytics.sh

set -e

echo "🔬 분석 프로그램 실행 시작..."

cd Project_HighSchool_apply_Analytics

echo "[1/3] Step 1: 데이터 전처리 중..."
python src/research_analytics.py

echo "[2/3] Step 2: 통계 분석 중..."
python src/statistical_deep_research.py

echo "[3/3] Step 3: HTML 리포트 생성 중..."
python src/final_dashboard_generator.py

echo "✅ 완료! output/ 디렉토리를 확인하세요."
echo "📊 대시보드를 새로고침하면 최신 리포트가 표시됩니다."
```

---

## 🔧 문제 해결

### ❓ "행정동 컬럼을 찾을 수 없음"

**원인:** research_analytics.py가 "행정동" 또는 "동" 컬럼을 찾을 수 없음

**해결:**
```python
# src/research_analytics.py 수정
KEY_DONG = "동"  # 또는 "주소(동)"

# 또는 데이터 파일의 컬럼 이름 변경
# "Address_Dong" → "행정동"
```

### ❓ "배정학교 컬럼이 없음"

**원인:** 컬럼명이 다름 (예: "배정고등학교", "할당학교" 등)

**해결:**
```python
# src/research_analytics.py 수정
KEY_ASSIGNED = "배정고등학교"  # 실제 컬럼명
```

### ❓ "데이터가 너무 적음 (분석 불가)"

**원인:** 유효 데이터 부족

**확인:**
```python
# statistical_deep_research.py 설정
MIN_SAMPLE_SCHOOL = 10  # 학교별 최소 인원
MIN_SAMPLE_DONG = 10    # 동네별 최소 인원

# 기준을 낮춰야 할 수도 있음:
MIN_SAMPLE_SCHOOL = 5
MIN_SAMPLE_DONG = 5
```

### ❓ HTML이 대시보드에 안 보임

**확인:**
```bash
# 1. 파일 생성 확인
ls -la Project_HighSchool_apply_Analytics/output/

# 2. 경로 확인
cat integrated_dashboard.py | grep output_dir

# 3. 파일명 확인 (예: Insight_Dashboard_2025.html)
```

---

## 💡 참고 자료

| 문서 | 내용 |
|:---:|:---:|
| `ANALYTICS_INTEGRATION.md` | 옵션 A/B/C 개요 |
| `DESIGN_IMPROVEMENTS.md` | 대시보드 UI 개선 |
| `ANNUAL_WORKFLOW.md` | 연간 일정 및 역할 |
| `src/research_analytics.py` | Step 1 상세 코드 |
| `src/statistical_deep_research.py` | Step 2-3 상세 코드 |

---

## ✨ 최종 흐름도

```
┌─ 2025 (완료) ────────────────────────────┐
│                                          │
│  Step1: 전처리 ✅                        │
│    ↓                                     │
│  Step2: 통계 분석 ✅                     │
│    ↓                                     │
│  Step3: HTML 생성 ✅                     │
│    ↓                                     │
│  output/Insight_Dashboard_2025.html     │
│    ↓ (현재 별도 조회)                    │
│  2025년 분석 결과 보유                   │
└──────────────────────────────────────────┘
                   ↓
┌─ 2026 (진행 중) ──────────────────────────┐
│                                          │
│  PHASE 4: 최종 배정 완료 (9월)          │
│    ↓                                     │
│  Step1 파일 생성 (행정동 데이터)        │
│    ↓                                     │
│  Step1-3 파이프라인 실행                 │
│    ↓                                     │
│  2026 HTML 리포트 생성                   │
│    ↓ (옵션 A)                            │
│  대시보드 [심층분석] 탭 임베딩          │
│    ↓                                     │
│  사용자: selectbox 선택 + expander 펼침 │
│    ↓                                     │
│  📊 2026 분석 결과 조회 (대시보드 내)   │
└──────────────────────────────────────────┘
```

---

**상태:** 📋 계획 단계 (Step 1 입력 대기)
**다음 리뷰:** 2026년 9월 (PHASE 4 시작 시)
**담당:** 분석팀 ↔ 대시보드팀 협력
