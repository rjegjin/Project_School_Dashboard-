import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import re
import os
from datetime import datetime
import plotly.express as px
from io import BytesIO
import paramiko

# ==========================================
# 설정
# ==========================================
WORKSPACE_ROOT = os.getenv("WORKSPACE_DIR", os.path.expanduser("~/projects"))
PROJECT_ROOT = os.path.join(WORKSPACE_ROOT, "Project_HighSchool_apply_Dashboard")
ANALYTICS_ROOT = os.path.join(WORKSPACE_ROOT, "Project_HighSchool_apply_Analytics")
VENV_PYTHON = os.path.join(WORKSPACE_ROOT, "unified_venv/bin/python")
KEY_FILE_PATH = os.path.join(WORKSPACE_ROOT, ".secrets/service_key.json")

st.set_page_config(
    page_title="📊 고입 진학현황 통합 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Google Sheets 인증
@st.cache_resource
def get_sheets_client():
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(KEY_FILE_PATH, scopes=SCOPES)
    return gspread.authorize(creds)

# ==========================================
# 데이터 로드
# ==========================================
@st.cache_data(ttl=300)
def load_2026_data():
    """2026년 데이터 (Google Sheets - 입시_트래킹)"""
    try:
        client = get_sheets_client()
        ss = client.open_by_key("14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0")

        tracking_sht = ss.worksheet("입시_트래킹")
        tracking_data = tracking_sht.get_all_values()

        df = pd.DataFrame(tracking_data[1:], columns=tracking_data[0])

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
        return df
    except Exception as e:
        st.error(f"2026 데이터 로드 실패: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def load_2026_progress_data():
    """2026년 입시 진행 현황 데이터 (Google Sheets - 입시 진행 현황)"""
    try:
        client = get_sheets_client()
        ss = client.open_by_key("14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0")

        progress_sht = ss.worksheet("입시 진행 현황")
        progress_data = progress_sht.get_all_values()

        if len(progress_data) < 2:
            return pd.DataFrame()

        df = pd.DataFrame(progress_data[1:], columns=progress_data[0])
        return df
    except Exception as e:
        return pd.DataFrame()

def load_2025_final_results():
    """2025 최종 합불 시트에서 합격자 정보 추출"""
    try:
        client = get_sheets_client()
        sheet_url = 'https://docs.google.com/spreadsheets/d/1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI/'
        doc = client.open_by_url(sheet_url)

        final_passed = {}  # {성명: (유형, 학교)}

        # 전기고 최종 합불
        try:
            early_sht = doc.worksheet("전기고 최종 합불")
            rows = early_sht.get_all_values()

            for row in rows:
                if len(row) >= 3 and row[2]:  # 이름 컬럼
                    name = row[2].strip()
                    # 과학고 섹션
                    if row[0] and row[0].strip().isdigit():
                        school = row[4].strip() if len(row) > 4 and row[4] else ""
                        if name and school:
                            final_passed[name] = ('전기고', school)
                    # 예고 섹션
                    if len(row) > 7 and row[7]:
                        school = row[9].strip() if len(row) > 9 and row[9] else ""
                        if row[7].strip().isdigit() and name and school:
                            final_passed[name] = ('전기고', school)
        except:
            pass

        # 후기고 최종 합불
        try:
            late_sht = doc.worksheet("후기고 최종 합불")
            rows = late_sht.get_all_values()

            for row in rows:
                if len(row) >= 3:
                    # 자사고 섹션
                    if row[0] and row[0].strip().isdigit():
                        name = row[2].strip() if row[2] else ""
                        result = row[5].strip() if len(row) > 5 and row[5] else ""
                        school = row[4].strip() if len(row) > 4 and row[4] else ""
                        if name and result == '합':
                            final_passed[name] = ('후기고', school)
                    # 외고 섹션
                    if len(row) > 8 and row[7] and row[7].strip().isdigit():
                        name = row[8].strip() if row[8] else ""
                        result = row[11].strip() if len(row) > 11 and row[11] else ""
                        school = row[10].strip() if len(row) > 10 and row[10] else ""
                        if name and result == '합':
                            final_passed[name] = ('후기고', school)
        except:
            pass

        return final_passed
    except Exception as e:
        return {}

@st.cache_data(ttl=300)
def load_2025_data():
    """
    2025년 데이터 (결과-우선 구조)
    generators/build_2025_master_results.py로 생성된 마스터 결과 시트 로드

    마스터 시트 컬럼:
      반 | 번호 | 성명 | 성별 | 최종유형 | 배정학교 | 합격여부 | 데이터출처
    """
    try:
        client = get_sheets_client()
        doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')

        # 2025_입시결과_마스터 시트 로드
        master_sht = doc.worksheet("2025_입시결과_마스터")
        rows = master_sht.get_all_values()

        if len(rows) < 2:
            st.warning("2025 마스터 결과 시트가 비어있습니다. generators/build_2025_master_results.py를 실행하세요.")
            return pd.DataFrame()

        # DataFrame으로 변환
        df = pd.DataFrame(rows[1:], columns=rows[0])

        # 컬럼 이름 통일 (대시보드 호환성)
        column_mapping = {
            '최종유형': '유형',
            '합격여부': '최종결과',
            '배정학교': '배정학교'
        }
        df = df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns})

        # 배정학교타입 열 추가 (대시보드 호환성)
        if '배정학교' not in df.columns:
            df['배정학교'] = ''
        if '배정학교타입' not in df.columns:
            df['배정학교타입'] = ''
            # 유형에서 배정학교타입 자동 설정
            if '유형' in df.columns:
                for idx, row in df.iterrows():
                    utype = str(row['유형']).strip()
                    if utype in ['과학고', '예고', '특성화고']:
                        df.at[idx, '배정학교타입'] = '전기고'
                    elif utype in ['자사고', '외고/국제고', '비평준화고']:
                        df.at[idx, '배정학교타입'] = '후기고'
                    elif utype == '일반고':
                        df.at[idx, '배정학교타입'] = '후기고'

        return df

    except Exception as e:
        st.warning(f"2025 마스터 시트 로드 실패: {e}\n\ngenerators/build_2025_master_results.py를 실행하세요.")
        return pd.DataFrame()

# ==========================================
# [레거시 코드 - 롤백용] (2025-03-22 주석처리)
# ==========================================
# @st.cache_data(ttl=300)
# def load_2025_data_legacy():
#     """2025년 데이터 (레거시 - 반별 시트) + 최종 결과 자동 병합"""
#     # [이전 복잡한 로직 400줄 - 불필요하므로 생략]
#     pass

@st.cache_data(ttl=300)
def load_local_final_results(file_path):
    """로컬 최종 결과 파일 읽기"""
    try:
        # file:// URL 처리
        if file_path.startswith('file://'):
            file_path = file_path[7:]

        if not os.path.exists(file_path):
            return None

        # 파일 읽기 (맨 첫 행부터)
        df = pd.read_excel(file_path)

        # 첫 행이 헤더처럼 보이면 스킵
        if df.iloc[0].isna().sum() > len(df.columns) * 0.5:
            df = df.iloc[1:].reset_index(drop=True)

        # 컬럼명 정규화 (줄바꿈 제거)
        df.columns = df.columns.str.replace('\r\n', ' ').str.strip()

        # 성명 컬럼 기준으로 유효한 행만 유지
        if '성명' in df.columns:
            df = df.dropna(subset=['성명'])

        return df
    except Exception as e:
        st.error(f"❌ 파일 읽기 실패: {e}")
        return None

# ==========================================
# UI 구성
# ==========================================
st.title("📊 고입 진학현황 통합 대시보드")
st.markdown("---")

# ==========================================
# 관리 도구 (우측 상단 버튼)
# ==========================================
with st.sidebar:
    st.divider()
    st.header("🔧 관리 도구")

    manage_col1, manage_col2 = st.columns(2)

    with manage_col1:
        if st.button("📌 메인대시보드", use_container_width=True, key="main_dashboard"):
            with st.spinner("메인 대시보드 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/create_main_dashboard.py"
                    ], capture_output=True, text=True, timeout=30)
                    if result.returncode == 0:
                        st.success("✅ 메인 대시보드 생성 완료!")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with manage_col2:
        if st.button("🔄 유형 동기화", use_container_width=True, key="sync_type"):
            with st.spinner("유형 자동 동기화 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/sync_type_to_tracking.py",
                        "2026",
                        "--apply-schema",
                        "--no-legacy-fill"
                    ], capture_output=True, text=True, timeout=30)
                    if result.returncode == 0:
                        st.success("✅ 유형 동기화 완료!")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    manage_col3, manage_col4 = st.columns(2)

    with manage_col3:
        if st.button("📋 합불 시트", use_container_width=True, key="final_sheets"):
            with st.spinner("최종 합불 시트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/generate_final_sheets.py",
                        "2026"
                    ], capture_output=True, text=True, timeout=30)
                    if result.returncode == 0:
                        st.success("✅ 합불 시트 생성 완료!")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with manage_col4:
        if st.button("📊 진행현황", use_container_width=True, key="progress_tracker"):
            with st.spinner("진행 현황 시트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/build_progress_from_tracking.py",
                        "2026"
                    ], capture_output=True, text=True, timeout=30)
                    if result.returncode == 0:
                        st.success("✅ 진행 현황 생성 완료!")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    st.divider()

# 좌측 사이드바
with st.sidebar:
    st.header("⚙️ 필터 및 설정")
    
    data_year = st.radio(
        "📅 데이터 선택",
        ["2026 (실시간)", "2025 (레거시)"]
    )
    
    st.divider()
    
    # 데이터 로드
    if data_year == "2026 (실시간)":
        df = load_2026_data()
        data_source = "Google Sheets - 입시_트래킹"
    else:
        df = load_2025_data()

        # 최종 결과 병합 상태 확인
        has_final_from_sheet = '최종결과' in df.columns and (df['최종결과'] == '합격').sum() > 0
        has_final_from_xlsx = '최종' in df.columns and df['최종'].notna().sum() > 0
        has_school = '배정학교' in df.columns and df['배정학교'].notna().sum() > 0

        if has_final_from_sheet or has_final_from_xlsx or has_school:
            data_source = "레거시 + 최종 결과 (자동 병합됨)"

            # 병합 상태 표시
            st.subheader("✅ 최종 결과 (자동 병합)")

            col_info1, col_info2, col_info3 = st.columns(3)
            with col_info1:
                # 최종 합불 시트에서의 합격
                if '최종결과' in df.columns:
                    final_sheet_count = (df['최종결과'] == '합격').sum()
                    pct = f"({final_sheet_count/len(df)*100:.1f}%)" if len(df) > 0 else "(0.0%)"
                    st.metric("📋 최종 합불 시트", f"{final_sheet_count}명", pct)
                else:
                    st.metric("📋 최종 합불 시트", "0명", "0.0%")

            with col_info2:
                # xlsx 배정 결과
                if '최종' in df.columns:
                    final_count = df['최종'].notna().sum()
                    pct = f"({final_count/len(df)*100:.1f}%)" if len(df) > 0 else "(0.0%)"
                    st.metric("📄 배정 결과 (xlsx)", f"{final_count}명", pct)
                else:
                    st.metric("📄 배정 결과 (xlsx)", "0명", "0.0%")

            with col_info3:
                if '배정학교' in df.columns:
                    school_count = df['배정학교'].notna().sum()
                    pct = f"({school_count/len(df)*100:.1f}%)" if len(df) > 0 else "(0.0%)"
                    st.metric("🏫 배정 학교", f"{school_count}명", pct)
                else:
                    st.metric("🏫 배정 학교", "0명", "0.0%")
        else:
            data_source = "레거시 시트 (반별 데이터)"
            st.subheader("⚠️ 최종 결과")
            st.info("최종 결과 정보를 찾을 수 없습니다.")

    st.subheader("🔍 필터")

    # 전체 데이터 백업 (통계용)
    df_full = df.copy()

    # 반 필터 (선택적)
    if not df.empty and '반' in df.columns:
        classes = sorted(df['반'].dropna().unique())

        # Default: 5반 (또는 첫 번째 반)
        default_class = []
        if '5' in classes:
            default_class = ['5']
        elif len(classes) > 0:
            default_class = [classes[0]]

        selected_class = st.multiselect(
            "반 선택 (비워두면 전체)",
            classes,
            default=default_class
        )

        # 필터 적용 (선택된 반이 있을 때만)
        if selected_class:
            df = df[df['반'].isin(selected_class)]
            filter_info = f"{', '.join(str(c) for c in selected_class)}반"
        else:
            filter_info = "전체"
    else:
        filter_info = "필터 없음"

    st.divider()
    st.caption(f"📍 데이터 소스: {data_source}")
    st.caption(f"🎯 현재 보기: {filter_info}")
    st.caption(f"⏰ 마지막 갱신: {datetime.now().strftime('%H:%M:%S')}")

# ==========================================
# 탭 구성
# ==========================================
if data_year == "2026 (실시간)":
    tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11 = st.tabs([  # 12개 탭
        "📅 연간 일정",
        "📈 전체 현황",
        "🎯 전기고",
        "🍂 후기고",
        "📊 심층 분석",
        "🔍 학생 조회",
        "📊 반별 현황",
        "🔄 동기화 상태",
        "📋 데이터 검증",
        "📈 필터링 분석",
        "🎓 학교 분석",
        "🔧 관리 도구"
    ])
    tab_progress = None
else:
    tab0, tab1, tab2, tab3, tab4, tab11 = st.tabs([
        "📅 연간 일정",
        "📈 전체 현황",
        "🎯 전기고",
        "🍂 후기고",
        "📊 심층 분석",
        "🔧 관리 도구"
    ])
    tab5 = tab6 = tab7 = tab8 = tab9 = tab10 = tab_progress = None

# ──────────────────────────────────────────
# TAB 0: 연간 일정
# ──────────────────────────────────────────
with tab0:
    # 마크다운 파일 읽기
    workflow_file = os.path.join(PROJECT_ROOT, "ANNUAL_WORKFLOW.md")
    try:
        with open(workflow_file, 'r', encoding='utf-8') as f:
            workflow_content = f.read()
        st.markdown(workflow_content)
    except FileNotFoundError:
        st.error(f"❌ 파일을 찾을 수 없습니다: {workflow_file}")
    except Exception as e:
        st.error(f"❌ 파일 읽기 오류: {e}")

# ──────────────────────────────────────────
# TAB 1: 전체 현황
# ──────────────────────────────────────────
with tab1:
    st.header("📈 전체 현황 (기준: 최종배정 > 접수 > 희망)")

    if not df_full.empty:
        # 통계 카드 (df_full 기준 - 전체 데이터)
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(
                "🎓 총 학생수",
                len(df_full),
                help="전체 지원자 수 (필터 제외)"
            )

        with col2:
            passed = len(df_full[df_full['최종'].astype(str).str.contains('합격', na=False)]) if '최종' in df_full.columns else 0
            st.metric(
                "🎉 합격",
                passed,
                help="최종 합격 학생 수 (전체)"
            )

        with col3:
            if '유형' in df_full.columns:
                early = len(df_full[df_full['유형'].isin(['과학고', '예술계고', '특성화고', '영재고'])])
                st.metric("🔵 전기고", early, help="전기고 지원자 (전체)")
            else:
                st.metric("🔵 전기고", 0)

        with col4:
            if '유형' in df_full.columns:
                late = len(df_full[df_full['유형'].isin(['자사고', '외고/국제고', '일반고'])])
                st.metric("🟠 후기고", late, help="후기고 지원자 (전체)")
            else:
                st.metric("🟠 후기고", 0)
        
        st.divider()
        
        # 유형별 분포
        if '유형' in df.columns and not df['유형'].isna().all():
            st.subheader("지원 유형 분포")
            
            col1, col2 = st.columns(2)
            
            with col1:
                type_dist = df['유형'].value_counts()
                fig = px.pie(
                    values=type_dist.values,
                    names=type_dist.index,
                    title="지원 유형 분포",
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                fig = px.bar(
                    x=type_dist.index,
                    y=type_dist.values,
                    title="유형별 지원자 수",
                    labels={'x': '유형', 'y': '인원'}
                )
                st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        
        # 데이터 테이블
        st.subheader("📋 전체 데이터")
        
        cols_to_show = [col for col in ['반', '번호', '성명', '성별', '유형', '최종'] if col in df.columns]
        display_df = df[cols_to_show].head(100)
        st.dataframe(display_df, use_container_width=True, height=400)
    else:
        st.warning("불러올 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 2: 전기고
# ──────────────────────────────────────────
with tab2:
    st.header("🎯 전기고 진학현황 (기준: 접수 우선, 미접수는 희망)")
    
    if not df.empty and '유형' in df.columns:
        early_types = ['과학고', '예술계고', '특성화고', '영재고']
        early_df = df[df['유형'].isin(early_types)]

        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📊 전기고 지원자", len(early_df))
        
        with col2:
            if not early_df.empty and '최종' in early_df.columns:
                passed = len(early_df[early_df['최종'].astype(str).str.contains('합격', na=False)])
                st.metric("🎉 합격자", passed)
        
        with col3:
            if not early_df.empty and len(early_df) > 0 and '최종' in early_df.columns:
                passed = len(early_df[early_df['최종'].astype(str).str.contains('합격', na=False)])
                rate = (passed / len(early_df) * 100)
                st.metric("📈 합격률", f"{rate:.1f}%")
        
        st.divider()
        
        # 유형별 현황
        if not early_df.empty and '유형' in early_df.columns:
            st.subheader("유형별 현황")
            
            type_summary = early_df['유형'].value_counts().reset_index()
            type_summary.columns = ['유형', '인원']
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig = px.bar(
                    type_summary,
                    x='유형',
                    y='인원',
                    title="유형별 지원자 수",
                    color='인원',
                    color_continuous_scale='Blues'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.dataframe(type_summary, use_container_width=True)
        
        st.divider()
        
        # 상세 데이터
        st.subheader("📋 전기고 상세 현황")
        
        if not early_df.empty:
            cols_to_show = [col for col in ['반', '성명', '유형', '최종'] if col in early_df.columns]
            display_df = early_df[cols_to_show].head(50)
            st.dataframe(display_df, use_container_width=True, height=400)
    else:
        st.info("전기고 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 3: 후기고
# ──────────────────────────────────────────
with tab3:
    st.header("🍂 후기고 진학현황 (기준: 접수 우선, 미접수는 희망)")
    
    if not df.empty and '유형' in df.columns:
        late_types = ['자사고', '외고/국제고', '일반고', '비평준화고']
        late_df = df[df['유형'].isin(late_types)]
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📊 후기고 지원자", len(late_df))
        
        with col2:
            if not late_df.empty and '최종' in late_df.columns:
                passed = len(late_df[late_df['최종'].astype(str).str.contains('합격', na=False)])
                st.metric("🎉 합격자", passed)
        
        with col3:
            if not late_df.empty and len(late_df) > 0 and '최종' in late_df.columns:
                passed = len(late_df[late_df['최종'].astype(str).str.contains('합격', na=False)])
                rate = (passed / len(late_df) * 100)
                st.metric("📈 합격률", f"{rate:.1f}%")
        
        st.divider()
        
        # 유형별 현황
        if not late_df.empty and '유형' in late_df.columns:
            st.subheader("유형별 현황")
            
            type_summary = late_df['유형'].value_counts().reset_index()
            type_summary.columns = ['유형', '인원']
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig = px.bar(
                    type_summary,
                    x='유형',
                    y='인원',
                    title="유형별 지원자 수",
                    color='인원',
                    color_continuous_scale='Oranges'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.dataframe(type_summary, use_container_width=True)
        
        st.divider()
        
        # 상세 데이터
        st.subheader("📋 후기고 상세 현황")
        
        if not late_df.empty:
            cols_to_show = [col for col in ['반', '성명', '유형', '최종'] if col in late_df.columns]
            display_df = late_df[cols_to_show].head(50)
            st.dataframe(display_df, use_container_width=True, height=400)
    else:
        st.info("후기고 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 4: 심층 분석
# ──────────────────────────────────────────
with tab4:
    st.header("📊 심층 분석")
    
    if not df.empty and '반' in df.columns:
        st.subheader("반별 분석")
        
        class_summary = df.groupby('반').agg({
            '성명': 'count',
        }).rename(columns={'성명': '지원자'})
        
        if '최종' in df.columns:
            class_summary['합격자'] = df[df['최종'].astype(str).str.contains('합격', na=False)].groupby('반').size()
            class_summary['합격자'] = class_summary['합격자'].fillna(0).astype(int)
            class_summary['합격률(%)'] = (class_summary['합격자'] / class_summary['지원자'] * 100).round(1)
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.bar(
                class_summary.reset_index(),
                x='반',
                y='지원자',
                title="반별 지원자 수",
                color='지원자'
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            if '합격률(%)' in class_summary.columns:
                fig = px.bar(
                    class_summary.reset_index(),
                    x='반',
                    y='합격률(%)',
                    title="반별 합격률",
                    color='합격률(%)',
                    color_continuous_scale='RdYlGn'
                )
                st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        st.subheader("반별 상세 통계")
        st.dataframe(class_summary, use_container_width=True)
    else:
        st.info("반별 분석 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 5: 입시 진행 현황 (2026만)
# ──────────────────────────────────────────
# ──────────────────────────────────────────
# TAB 8: 데이터 검증 리포트
# ──────────────────────────────────────────
if tab8 is not None:
    with tab8:
        st.header("📋 데이터 검증 리포트")

        if not df_full.empty:
            # 1. 누락된 정보 현황
            st.subheader("1️⃣ 누락된 정보 현황")

            missing_data = {}
            for col in ['반', '번호', '성명', '성별', '유형', '최종', '배정학교']:
                if col in df_full.columns:
                    missing_count = (df_full[col] == '').sum() + df_full[col].isna().sum()
                    missing_pct = (missing_count / len(df_full) * 100) if len(df_full) > 0 else 0
                    missing_data[col] = {'누락': missing_count, '비율(%)': round(missing_pct, 1)}

            missing_df = pd.DataFrame(missing_data).T
            missing_df = missing_df[missing_df['누락'] > 0]  # 누락이 있는 것만

            if len(missing_df) > 0:
                st.warning(f"⚠️ {len(missing_df)}개 항목에 누락된 데이터가 있습니다.")
                st.dataframe(missing_df, use_container_width=True)

                # 누락된 학생 목록
                st.write("**누락된 학생 목록:**")
                for col in missing_df.index:
                    missing_students = df_full[df_full[col] == '']
                    if len(missing_students) > 0:
                        st.write(f"📌 **{col} 누락 ({len(missing_students)}명)**")
                        st.dataframe(
                            missing_students[['반', '번호', '성명']],
                            use_container_width=True,
                            hide_index=True
                        )
            else:
                st.success("✅ 누락된 데이터가 없습니다!")

            st.divider()

            # 2. 중복 확인
            st.subheader("2️⃣ 중복 데이터 확인")

            if '성명' in df_full.columns:
                duplicate_names = df_full[df_full.duplicated(subset=['성명'], keep=False)].sort_values('성명')
                if len(duplicate_names) > 0:
                    st.warning(f"⚠️ {len(duplicate_names) // 2}명의 중복된 학생이 있습니다.")
                    st.dataframe(
                        duplicate_names[['반', '번호', '성명', '유형']],
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.success("✅ 중복된 학생이 없습니다.")

            st.divider()

            # 3. 데이터 오류 감지
            st.subheader("3️⃣ 데이터 오류 감지")

            error_count = 0

            # 유형 오류
            if '유형' in df_full.columns:
                valid_types = ['영재고', '과학고', '예술고', '예술계고', '특성화고', '자사고', '외고', '국제고', '외고/국제고', '일반고', '비평준화고']
                invalid_types = df_full[~df_full['유형'].isin(valid_types + [''])]
                if len(invalid_types) > 0:
                    st.error(f"❌ 유형 오류: {len(invalid_types)}명")
                    st.dataframe(
                        invalid_types[['반', '번호', '성명', '유형']],
                        use_container_width=True,
                        hide_index=True
                    )
                    error_count += len(invalid_types)

            # 최종 결과 오류
            if '최종' in df_full.columns:
                valid_finals = ['합격', '불합격', '']
                invalid_finals = df_full[~df_full['최종'].isin(valid_finals)]
                if len(invalid_finals) > 0:
                    st.error(f"❌ 최종 결과 오류: {len(invalid_finals)}명")
                    st.dataframe(
                        invalid_finals[['반', '번호', '성명', '최종']],
                        use_container_width=True,
                        hide_index=True
                    )
                    error_count += len(invalid_finals)

            if error_count == 0:
                st.success("✅ 데이터 오류가 없습니다!")

            st.divider()

            # 4. 데이터 품질 점수
            st.subheader("4️⃣ 데이터 품질 점수")

            total_students = len(df_full)
            quality_checks = []

            checks = {
                '반 입력': ('반' in df_full.columns and (df_full['반'] != '').sum()),
                '번호 입력': ('번호' in df_full.columns and (df_full['번호'] != '').sum()),
                '성명 입력': ('성명' in df_full.columns and (df_full['성명'] != '').sum()),
                '유형 입력': ('유형' in df_full.columns and (df_full['유형'] != '').sum()),
                '최종 입력': ('최종' in df_full.columns and (df_full['최종'] != '').sum()),
                '배정학교 입력': ('배정학교' in df_full.columns and (df_full['배정학교'] != '').sum()),
            }

            for check_name, filled_count in checks.items():
                if isinstance(filled_count, bool):
                    filled_count = total_students if filled_count else 0
                pct = (filled_count / total_students * 100) if total_students > 0 else 0
                quality_checks.append({'항목': check_name, '완성도(%)': round(pct, 1)})

            quality_df = pd.DataFrame(quality_checks)
            overall_score = quality_df['완성도(%)'].mean()

            fig = px.bar(
                quality_df,
                x='항목',
                y='완성도(%)',
                title=f"데이터 품질 점수 (평균: {overall_score:.1f}%)",
                range_y=[0, 100],
                color='완성도(%)',
                color_continuous_scale=['#ff6b6b', '#ffd93d', '#6bcf7f']
            )
            st.plotly_chart(fig, use_container_width=True)

            # 품질 등급
            if overall_score >= 90:
                st.success(f"⭐⭐⭐ 매우 우수: {overall_score:.1f}%")
            elif overall_score >= 70:
                st.info(f"⭐⭐ 우수: {overall_score:.1f}%")
            elif overall_score >= 50:
                st.warning(f"⭐ 보통: {overall_score:.1f}%")
            else:
                st.error(f"❌ 부족: {overall_score:.1f}%")

        else:
            st.warning("⚠️ 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 9: 필터링 분석
# ──────────────────────────────────────────
if tab9 is not None:
    with tab9:
        st.header("📈 필터링 분석")

        if not df_full.empty:
            st.subheader("🔍 필터 조건 선택")

            col_f1, col_f2, col_f3 = st.columns(3)

            with col_f1:
                selected_types = st.multiselect(
                    "지원 유형",
                    df_full['유형'].unique(),
                    default=df_full['유형'].unique(),
                    key="filter_type"
                )

            with col_f2:
                selected_finals = st.multiselect(
                    "최종 결과",
                    [x for x in df_full['최종'].unique() if x],
                    default=[x for x in df_full['최종'].unique() if x],
                    key="filter_final"
                )

            with col_f3:
                selected_classes = st.multiselect(
                    "반",
                    sorted(df_full['반'].unique()),
                    default=sorted(df_full['반'].unique()),
                    key="filter_class"
                )

            # 필터 적용
            filtered_df = df_full[
                (df_full['유형'].isin(selected_types)) &
                (df_full['최종'].isin(selected_finals)) &
                (df_full['반'].isin(selected_classes))
            ]

            st.divider()

            # 필터 결과 통계
            col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

            with col_stat1:
                st.metric("필터된 학생", len(filtered_df))

            with col_stat2:
                if '최종' in filtered_df.columns:
                    passed = (filtered_df['최종'] == '합격').sum()
                    st.metric("합격자", passed)
                else:
                    st.metric("합격자", 0)

            with col_stat3:
                if len(filtered_df) > 0 and '최종' in filtered_df.columns:
                    passed = (filtered_df['최종'] == '합격').sum()
                    pct = (passed / len(filtered_df) * 100) if len(filtered_df) > 0 else 0
                    st.metric("합격률", f"{pct:.1f}%")
                else:
                    st.metric("합격률", "0.0%")

            with col_stat4:
                if '배정학교' in filtered_df.columns:
                    school_count = (filtered_df['배정학교'] != '').sum()
                    st.metric("배정학교", school_count)
                else:
                    st.metric("배정학교", 0)

            st.divider()

            # 필터된 데이터 시각화
            col_viz1, col_viz2 = st.columns(2)

            with col_viz1:
                if '유형' in filtered_df.columns:
                    type_dist = filtered_df['유형'].value_counts().reset_index()
                    type_dist.columns = ['유형', '인원']

                    fig = px.pie(
                        type_dist,
                        values='인원',
                        names='유형',
                        title="지원 유형 분포"
                    )
                    st.plotly_chart(fig, use_container_width=True)

            with col_viz2:
                if '최종' in filtered_df.columns:
                    final_dist = filtered_df['최종'].value_counts().reset_index()
                    final_dist.columns = ['결과', '인원']
                    final_dist = final_dist[final_dist['결과'] != '']

                    if len(final_dist) > 0:
                        fig = px.bar(
                            final_dist,
                            x='결과',
                            y='인원',
                            title="최종 결과 분포",
                            color='결과',
                            color_discrete_map={'합격': '#1f77b4', '불합격': '#ff7f0e'}
                        )
                        st.plotly_chart(fig, use_container_width=True)

            st.divider()

            # 필터된 학생 목록
            st.subheader("📋 필터된 학생 목록")

            display_cols = ['반', '번호', '성명', '성별', '유형', '1차', '2차', '최종']
            available_cols = [col for col in display_cols if col in filtered_df.columns]

            st.dataframe(
                filtered_df[available_cols].sort_values(['반', '번호']),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("⚠️ 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 10: 학교 분석
# ──────────────────────────────────────────
if tab10 is not None:
    with tab10:
        st.header("🎓 학교별 합격 분석")

        if not df_full.empty and '배정학교' in df_full.columns:
            # 배정 학교 통계
            school_df = df_full[df_full['배정학교'] != ''].copy()

            if len(school_df) > 0:
                school_stats = school_df.groupby('배정학교').agg({
                    '성명': 'count',
                }).rename(columns={'성명': '인원'})

                # 최종 합격자 포함 분석
                if '최종' in school_df.columns:
                    school_stats['합격'] = school_df[school_df['최종'] == '합격'].groupby('배정학교').size()
                    school_stats['합격'] = school_stats['합격'].fillna(0).astype(int)
                    school_stats['합격률(%)'] = (school_stats['합격'] / school_stats['인원'] * 100).round(1)

                school_stats = school_stats.sort_values('인원', ascending=False)

                st.subheader("📊 학교별 배정 현황")

                col_school1, col_school2 = st.columns(2)

                with col_school1:
                    st.write("**상위 10개 학교**")
                    fig = px.bar(
                        school_stats.head(10).reset_index(),
                        x='배정학교',
                        y='인원',
                        title="상위 10개 배정 학교",
                        color='인원'
                    )
                    st.plotly_chart(fig, use_container_width=True)

                with col_school2:
                    if '합격률(%)' in school_stats.columns:
                        st.write("**학교별 합격률**")
                        fig = px.bar(
                            school_stats[school_stats['합격'] > 0].reset_index(),
                            x='배정학교',
                            y='합격률(%)',
                            title="학교별 합격률 (합격자 있는 학교만)",
                            range_y=[0, 100],
                            color='합격률(%)',
                            color_continuous_scale=['#ff6b6b', '#ffd93d', '#6bcf7f']
                        )
                        st.plotly_chart(fig, use_container_width=True)

                st.divider()

                st.subheader("📋 학교별 상세 통계")
                st.dataframe(school_stats, use_container_width=True)

                st.divider()

                # 유형별 배정 학교
                st.subheader("🎯 지원 유형별 배정 학교")

                if '유형' in school_df.columns:
                    for school_type in sorted(school_df['유형'].unique()):
                        type_schools = school_df[school_df['유형'] == school_type]['배정학교'].value_counts().head(5)
                        if len(type_schools) > 0:
                            st.write(f"**{school_type}** → {', '.join(type_schools.index.tolist())}")

            else:
                st.info("ℹ️ 배정학교 정보가 없습니다.")
        else:
            st.warning("⚠️ 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 11: 관리 도구 (모든 생성/동기화 기능)
# ──────────────────────────────────────────
with tab11:
    st.header("🔧 고입 관리 도구")
    st.markdown("설문지 응답 동기화 → 반별 시트 확인 → 아래 순서대로 실행하세요:")

    # Step 0: 설문지 → 반별 시트
    st.subheader("0️⃣ 설문지 응답 → 반별 시트 동기화")

    form_col1, form_col2 = st.columns([3, 1])
    with form_col1:
        st.caption("학생 설문지 응답을 301~314 반별 시트에 자동으로 반영합니다. 이미 입력된 셀은 덮어쓰지 않습니다.")
    with form_col2:
        form_dry_run = st.checkbox("미리보기 (dry-run)", value=True, key="form_sync_dry")

    if st.button("📥 설문지 → 반별 시트 동기화", use_container_width=True, key="tool_sync_form"):
        with st.spinner("설문지 응답 동기화 중..."):
            try:
                import subprocess
                cmd = [
                    VENV_PYTHON,
                    os.path.join(PROJECT_ROOT, "generators/sync_form_to_class_sheets.py"),
                ]
                if form_dry_run:
                    cmd.append("--dry-run")
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                if result.returncode == 0:
                    if form_dry_run:
                        st.info("🔍 DRY RUN 결과 (실제 변경 없음)")
                    else:
                        st.success("✅ 동기화 완료!")
                    st.code(result.stdout, language="text")
                else:
                    st.error(f"❌ 실패: {result.stderr}")
            except Exception as e:
                st.error(f"❌ 오류: {str(e)}")

    st.divider()

    # Step 1: 동기화
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ 데이터 동기화")

        if st.button("🔄 유형 자동 동기화", use_container_width=True, key="tool_sync_type"):
            st.info("반별 시트 → 입시_트래킹 유형 자동 감지 (1회성)")
            with st.spinner("유형 자동 동기화 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/sync_type_to_tracking.py",
                        "2026",
                        "--apply-schema",
                        "--no-legacy-fill"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 유형 동기화 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with col2:
        st.subheader("2️⃣ 시트 생성")

        if st.button("📋 최종 합불 시트 생성", use_container_width=True, key="tool_final_sheets"):
            st.info("입시_트래킹 → 전기고/후기고 최종 합불 자동 생성")
            with st.spinner("최종 합불 시트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/generate_final_sheets.py",
                        "2026"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 합불 시트 생성 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    st.divider()

    # Step 2: 대시보드 생성
    col3, col4 = st.columns(2)

    with col3:
        st.subheader("3️⃣ 대시보드 생성")

        st.warning("⚠️ 일회성 — 기존 📌 메인 탭(하이퍼링크·응답률) 전체 덮어씀")
        if st.button("📌 메인 대시보드 생성 [일회성]", use_container_width=True, key="tool_main_dashboard"):
            st.info("최초 1회 세팅 전용 — 이미 내용이 있으면 주의!")
            with st.spinner("메인 대시보드 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/create_main_dashboard.py"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 메인 대시보드 생성 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with col4:
        st.subheader("4️⃣ 진행 현황")

        if st.button("📊 입시 진행 현황 생성", use_container_width=True, key="tool_progress"):
            st.info("입시_트래킹 → 입시 진행 현황 시트 생성")
            with st.spinner("진행 현황 시트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/build_progress_from_tracking.py",
                        "2026"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 진행 현황 생성 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    st.divider()

    # Step 3: 리포트 생성
    col5, col6 = st.columns(2)

    with col5:
        st.subheader("5️⃣ 리포트 생성")

        if st.button("📄 진학현황 리포트 (HTML+Excel)", use_container_width=True, key="tool_report"):
            st.info("전기/후기고 진학현황 표형 리포트 생성")
            with st.spinner("리포트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/mokil_high_school_results_gen.py"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 리포트 생성 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with col6:
        st.subheader("6️⃣ 심화 분석")

        if st.button("🎨 컬러리포트 생성 (워터폴)", use_container_width=True, key="tool_color_report"):
            st.info("전기/후기고 진행상황 컬러리포트 (배지/워터폴)")
            with st.spinner("컬러리포트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/generate_table.py",
                        "2026"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 컬러리포트 생성 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    st.divider()

    # Step 4: 심층 통계 분석
    st.subheader("7️⃣ 심층 통계 분석")
    st.markdown("⚠️ 최종 배정 결과 데이터가 필요합니다 (9월 이후)")

    col7, col8, col9 = st.columns(3)

    with col7:
        if st.button("📊 Step1: 데이터 전처리", use_container_width=True, key="analytics_step1"):
            st.info("배정 결과 → 학교별 경쟁률/만족도 분석\n(Step2_지망선호도_및_지역흐름.xlsx 생성)")
            with st.spinner("Step 1 실행 중... (약 10초)"):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        os.path.join(ANALYTICS_ROOT, "src/research_analytics.py")
                    ], capture_output=True, text=True, timeout=120, cwd=ANALYTICS_ROOT)
                    if result.returncode == 0:
                        st.success("✅ Step 1 완료!")
                        st.info(result.stdout)
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with col8:
        if st.button("🔬 Step2: 통계 분석", use_container_width=True, key="analytics_step2"):
            st.info("군집화, 카이제곱 검정, 상관관계\n(Step3_학교유형화_및_통계검증.xlsx 생성)")
            with st.spinner("Step 2 실행 중... (약 5초)"):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        os.path.join(ANALYTICS_ROOT, "src/statistical_deep_research.py")
                    ], capture_output=True, text=True, timeout=120, cwd=ANALYTICS_ROOT)
                    if result.returncode == 0:
                        st.success("✅ Step 2 완료!")
                        st.info(result.stdout)
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with col9:
        if st.button("📄 Step3: HTML 리포트", use_container_width=True, key="analytics_step3"):
            st.info("Step3 → HTML 대시보드 생성\n(output/Insight_Dashboard.html 생성)")
            with st.spinner("Step 3 실행 중... (약 3초)"):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        os.path.join(ANALYTICS_ROOT, "src/final_dashboard_generator.py")
                    ], capture_output=True, text=True, timeout=120, cwd=ANALYTICS_ROOT)
                    if result.returncode == 0:
                        st.success("✅ Step 3 완료!")
                        st.info(result.stdout)
                        st.info("💡 리포트 뷰어에서 생성된 HTML을 확인하세요")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    st.divider()

    # 📊 리포트 뷰어 섹션
    st.subheader("📊 생성된 리포트 뷰어")

    report_dir = "reports"
    if os.path.exists(report_dir):
        html_files = [f for f in os.listdir(report_dir) if f.endswith('.html') and not f.startswith('_')]
        excel_files = [f for f in os.listdir(report_dir) if f.endswith('.xlsx') and not f.startswith('_')]

        if html_files or excel_files:
            # 리포트 분류 (_legacy/ 폴더 파일은 이미 위에서 제외됨)
            early_reports   = sorted([f for f in html_files if '전기고' in f])
            late_reports    = sorted([f for f in html_files if '후기고' in f])
            special_reports = sorted([f for f in html_files if '특별전형' in f])
            other_reports   = sorted([f for f in html_files if f not in early_reports + late_reports + special_reports])

            # 탭 구성
            tab_early, tab_late, tab_special, tab_excel = st.tabs([
                f"🎯 전기고 ({len(early_reports)})",
                f"🍂 후기고 ({len(late_reports)})",
                f"🏷️ 특별전형 ({len(special_reports)})",
                f"📊 Excel ({len(excel_files)})"
            ])

            # ─── 전기고 리포트 탭 ───
            with tab_early:
                if early_reports:
                    # 다운로드 버튼 행
                    st.markdown("**📥 리포트 다운로드**")
                    cols = st.columns(len(early_reports))
                    for idx, report in enumerate(early_reports):
                        with cols[idx]:
                            report_path = os.path.join(report_dir, report)
                            with open(report_path, 'rb') as f:
                                st.download_button(
                                    label=f"⬇️ {report}",
                                    data=f.read(),
                                    file_name=report,
                                    mime="text/html",
                                    use_container_width=True,
                                    key=f"dl_early_{idx}"
                                )

                    st.divider()

                    # 리포트 선택 및 미리보기
                    st.markdown("**🔍 리포트 미리보기**")
                    selected_early = st.selectbox(
                        "보기를 원하는 리포트 선택",
                        early_reports,
                        key="select_early"
                    )

                    if selected_early:
                        report_path = os.path.join(report_dir, selected_early)
                        try:
                            with open(report_path, 'r', encoding='utf-8') as f:
                                html_content = f.read()

                            with st.expander(f"📄 {selected_early} 내용 보기 (클릭하여 펼치기)", expanded=False):
                                st.components.v1.html(html_content, height=700, scrolling=True)
                        except Exception as e:
                            st.error(f"파일 읽기 실패: {e}")
                else:
                    st.info("📄 전기고 리포트가 없습니다.\n[관리 도구]에서 생성하세요.")

            # ─── 후기고 리포트 탭 ───
            with tab_late:
                if late_reports:
                    # 다운로드 버튼 행
                    st.markdown("**📥 리포트 다운로드**")
                    cols = st.columns(len(late_reports))
                    for idx, report in enumerate(late_reports):
                        with cols[idx]:
                            report_path = os.path.join(report_dir, report)
                            with open(report_path, 'rb') as f:
                                st.download_button(
                                    label=f"⬇️ {report}",
                                    data=f.read(),
                                    file_name=report,
                                    mime="text/html",
                                    use_container_width=True,
                                    key=f"dl_late_{idx}"
                                )

                    st.divider()

                    # 리포트 선택 및 미리보기
                    st.markdown("**🔍 리포트 미리보기**")
                    selected_late = st.selectbox(
                        "보기를 원하는 리포트 선택",
                        late_reports,
                        key="select_late"
                    )

                    if selected_late:
                        report_path = os.path.join(report_dir, selected_late)
                        try:
                            with open(report_path, 'r', encoding='utf-8') as f:
                                html_content = f.read()

                            with st.expander(f"📄 {selected_late} 내용 보기 (클릭하여 펼치기)", expanded=False):
                                st.components.v1.html(html_content, height=700, scrolling=True)
                        except Exception as e:
                            st.error(f"파일 읽기 실패: {e}")
                else:
                    st.info("📄 후기고 리포트가 없습니다.\n[관리 도구]에서 생성하세요.")

            # ─── 특별전형 리포트 탭 ───
            with tab_special:
                if special_reports:
                    st.markdown("**📥 특별전형 현황 리포트 다운로드**")
                    cols = st.columns(len(special_reports))
                    for idx, report in enumerate(special_reports):
                        with cols[idx]:
                            report_path = os.path.join(report_dir, report)
                            with open(report_path, 'rb') as f:
                                st.download_button(
                                    label=f"⬇️ {report}",
                                    data=f.read(),
                                    file_name=report,
                                    mime="text/html",
                                    use_container_width=True,
                                    key=f"dl_special_{idx}"
                                )
                    st.divider()
                    selected_special = st.selectbox("보기를 원하는 리포트 선택", special_reports, key="select_special")
                    if selected_special:
                        report_path = os.path.join(report_dir, selected_special)
                        try:
                            with open(report_path, 'r', encoding='utf-8') as f:
                                html_content = f.read()
                            with st.expander(f"📄 {selected_special} 내용 보기", expanded=True):
                                st.components.v1.html(html_content, height=700, scrolling=True)
                        except Exception as e:
                            st.error(f"파일 읽기 실패: {e}")
                else:
                    st.info("🏷️ 특별전형 현황 리포트가 없습니다.\n[8️⃣ 특별전형 트래킹]에서 생성하세요.")

            # ─── Excel 파일 탭 ───
            with tab_excel:
                if excel_files:
                    st.markdown("**📥 Excel 파일 다운로드**")
                    cols = st.columns(min(3, len(excel_files)))
                    for idx, excel_file in enumerate(excel_files):
                        with cols[idx % 3]:
                            excel_path = os.path.join(report_dir, excel_file)
                            with open(excel_path, 'rb') as f:
                                st.download_button(
                                    label=f"📥 {excel_file}",
                                    data=f.read(),
                                    file_name=excel_file,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key=f"dl_excel_{idx}"
                                )
                else:
                    st.info("📊 Excel 파일이 없습니다.")

        else:
            st.info("📄 생성된 리포트가 없습니다.\n[관리 도구]에서 리포트를 생성하세요.")
    else:
        st.warning("⚠️ reports 폴더가 없습니다.")

    st.divider()

    # ── 특별전형 트래킹 ────────────────────────────────────────
    st.subheader("8️⃣ 특별전형 트래킹")
    st.caption(
        "사회통합전형 · 특례 · 보훈 · 쌍둥이 · 학폭 · 교직원자녀 · 장애 — "
        "반별 시트(301~314) 특별전형 컬럼 → **특별전형_트래킹** 시트 동기화 후 HTML 리포트 생성"
    )

    sp_col1, sp_col2 = st.columns(2)

    with sp_col1:
        if st.button("🔁 특별전형 동기화 (반별→트래킹)", use_container_width=True, key="tool_special_sync"):
            with st.spinner("특별전형_트래킹 시트 갱신 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        os.path.join(PROJECT_ROOT, "generators/sync_special_to_tracking.py"),
                        "2026"
                    ], capture_output=True, text=True, timeout=120)
                    if result.returncode == 0:
                        st.success("✅ 특별전형_트래킹 시트 갱신 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with sp_col2:
        if st.button("📊 특별전형 현황 HTML 생성", use_container_width=True, key="tool_special_report"):
            with st.spinner("특별전형 HTML 리포트 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        os.path.join(PROJECT_ROOT, "generators/generate_special_report.py"),
                        "2026"
                    ], capture_output=True, text=True, timeout=120)
                    if result.returncode == 0:
                        st.success("✅ 특별전형_현황.html 생성 완료!")
                        st.code(result.stdout, language="text")
                        # 다운로드 버튼
                        special_html_path = os.path.join("reports", "특별전형_현황.html")
                        if os.path.exists(special_html_path):
                            with open(special_html_path, "rb") as f:
                                st.download_button(
                                    label="⬇️ 특별전형_현황.html 다운로드",
                                    data=f.read(),
                                    file_name="특별전형_현황.html",
                                    mime="text/html",
                                    use_container_width=True,
                                    key="dl_special_report",
                                )
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    st.divider()

    # 추가 도구
    col7, col8 = st.columns(2)

    with col7:
        st.subheader("📊 카드 대시보드")

        if st.button("🎴 카드형 대시보드 생성", use_container_width=True, key="tool_card_dashboard"):
            st.info("합격/불합격 카드 대시보드 업데이트")
            with st.spinner("카드 대시보드 생성 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Dashboard/generators/generate_dashboard.py"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 카드 대시보드 생성 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

    with col8:
        st.subheader("🔍 심층 분석")

        if st.button("📈 심층 분석 실행", use_container_width=True, key="tool_analytics"):
            st.info("최종 배정 결과 통계 및 분석")
            with st.spinner("심층 분석 실행 중..."):
                try:
                    import subprocess
                    result = subprocess.run([
                        VENV_PYTHON,
                        "Project_HighSchool_apply_Analytics/src/advanced_analytics_engine.py"
                    ], capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        st.success("✅ 분석 완료!")
                        st.code(result.stdout, language="text")
                    else:
                        st.error(f"❌ 실패: {result.stderr}")
                except Exception as e:
                    st.error(f"❌ 오류: {str(e)}")

# ──────────────────────────────────────────
# TAB 5: 학생 조회
# ──────────────────────────────────────────
if tab5 is not None:
    with tab5:
        st.header("🔍 학생 조회")

        if not df_full.empty and '성명' in df_full.columns:
            search_name = st.text_input("📝 학생 이름 검색", placeholder="예: 강은서")

            if search_name:
                results = df_full[df_full['성명'].str.contains(search_name, na=False)]

                if len(results) > 0:
                    st.success(f"✅ {len(results)}명의 학생을 찾았습니다.")

                    for idx, (_, student) in enumerate(results.iterrows(), 1):
                        with st.container():
                            col1, col2, col3, col4 = st.columns(4)

                            with col1:
                                st.metric("반", student.get('반', '-'))
                            with col2:
                                st.metric("번호", student.get('번호', '-'))
                            with col3:
                                st.metric("성별", student.get('성별', '-'))
                            with col4:
                                st.metric("지원유형", student.get('유형', '-'))

                            st.markdown("---")

                            # 상세 정보
                            col_a, col_b, col_c, col_d = st.columns(4)

                            with col_a:
                                st.write("**1차**")
                                st.write(student.get('1차', '-'))

                            with col_b:
                                st.write("**2차**")
                                st.write(student.get('2차', '-'))

                            with col_c:
                                st.write("**최종**")
                                final = student.get('최종', '-')
                                if final == '합격':
                                    st.success(final)
                                elif final == '불합격':
                                    st.error(final)
                                else:
                                    st.write(final)

                            with col_d:
                                st.write("**배정학교**")
                                st.write(student.get('배정학교', '-'))

                            st.write(f"**비고**: {student.get('비고', '-')}")
                            st.divider()
                else:
                    st.warning("❌ 검색 결과가 없습니다.")
            else:
                st.info("💡 학생 이름을 입력하여 검색하세요.")
        else:
            st.warning("⚠️ 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 6: 반별 현황
# ──────────────────────────────────────────
if tab6 is not None:
    with tab6:
        st.header("📊 반별 진학 현황")

        if not df_full.empty and '반' in df_full.columns:
            classes = sorted(df_full['반'].dropna().unique())

            selected_class = st.selectbox(
                "반 선택",
                classes,
                format_func=lambda x: f"{x}반"
            )

            # 선택된 반의 데이터
            class_df = df_full[df_full['반'] == selected_class]

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric("👥 전체 학생", len(class_df))

            with col2:
                if '최종' in class_df.columns:
                    passed = (class_df['최종'] == '합격').sum()
                    st.metric("🎉 합격자", passed)
                else:
                    st.metric("🎉 합격자", 0)

            with col3:
                if '유형' in class_df.columns:
                    early = len(class_df[class_df['유형'].isin(['과학고', '예술계고', '특성화고', '영재고'])])
                    st.metric("🔵 전기고", early)
                else:
                    st.metric("🔵 전기고", 0)

            with col4:
                if '유형' in class_df.columns:
                    late = len(class_df[class_df['유형'].isin(['자사고', '외고/국제고', '일반고'])])
                    st.metric("🟠 후기고", late)
                else:
                    st.metric("🟠 후기고", 0)

            st.divider()

            # 지원 유형별 분포
            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                if '유형' in class_df.columns and not class_df['유형'].isna().all():
                    type_dist = class_df['유형'].value_counts().reset_index()
                    type_dist.columns = ['지원유형', '인원']

                    fig = px.pie(
                        type_dist,
                        values='인원',
                        names='지원유형',
                        title=f"{selected_class}반 지원 유형 분포"
                    )
                    st.plotly_chart(fig, use_container_width=True)

            with col_chart2:
                if '최종' in class_df.columns:
                    final_dist = class_df['최종'].value_counts().reset_index()
                    final_dist.columns = ['최종결과', '인원']
                    final_dist = final_dist[final_dist['최종결과'] != '']

                    if not final_dist.empty:
                        fig = px.bar(
                            final_dist,
                            x='최종결과',
                            y='인원',
                            title=f"{selected_class}반 최종 결과",
                            color='최종결과',
                            color_discrete_map={'합격': '#1f77b4', '불합격': '#ff7f0e'}
                        )
                        st.plotly_chart(fig, use_container_width=True)

            st.divider()

            # 반별 학생 목록
            st.subheader("📋 학생 목록")

            display_cols = ['번호', '성명', '성별', '유형', '1차', '2차', '최종', '배정학교']
            available_cols = [col for col in display_cols if col in class_df.columns]

            st.dataframe(
                class_df[available_cols].sort_values('번호'),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("⚠️ 데이터가 없습니다.")

# ──────────────────────────────────────────
# TAB 7: 동기화 상태
# ──────────────────────────────────────────
if tab7 is not None:
    with tab7:
        st.header("🔄 동기화 상태 & 데이터 현황")

        col1, col2, col3, col4 = st.columns(4)

        # 입시_트래킹 현황
        with col1:
            total_students = len(df_full) if not df_full.empty else 0
            st.metric("📊 입시_트래킹", f"{total_students}명", "총 학생 수")

        # 유형 반영도
        with col2:
            if not df_full.empty and '유형' in df_full.columns:
                type_filled = (df_full['유형'] != '').sum()
                type_pct = (type_filled / total_students * 100) if total_students > 0 else 0
                st.metric("📋 유형 반영도", f"{type_filled}/{total_students}", f"{type_pct:.1f}%")
            else:
                st.metric("📋 유형 반영도", "0/0", "0.0%")

        # 최종 반영도
        with col3:
            if not df_full.empty and '최종' in df_full.columns:
                final_filled = (df_full['최종'] != '').sum()
                final_pct = (final_filled / total_students * 100) if total_students > 0 else 0
                st.metric("✅ 최종 반영도", f"{final_filled}/{total_students}", f"{final_pct:.1f}%")
            else:
                st.metric("✅ 최종 반영도", "0/0", "0.0%")

        # 배정학교 반영도
        with col4:
            if not df_full.empty and '배정학교' in df_full.columns:
                school_filled = (df_full['배정학교'] != '').sum()
                school_pct = (school_filled / total_students * 100) if total_students > 0 else 0
                st.metric("🏫 배정학교 반영도", f"{school_filled}/{total_students}", f"{school_pct:.1f}%")
            else:
                st.metric("🏫 배정학교 반영도", "0/0", "0.0%")

        st.divider()

        # 상세 통계
        st.subheader("📈 상세 통계")

        col_stat1, col_stat2, col_stat3 = st.columns(3)

        with col_stat1:
            st.write("**지원 유형별 현황**")
            if not df_full.empty and '유형' in df_full.columns:
                type_stats = df_full['유형'].value_counts()
                st.bar_chart(type_stats)

        with col_stat2:
            st.write("**최종 결과별 현황**")
            if not df_full.empty and '최종' in df_full.columns:
                final_stats = df_full['최종'].value_counts()
                final_stats = final_stats[final_stats.index != '']  # 빈값 제외
                st.bar_chart(final_stats)

        with col_stat3:
            st.write("**배정학교 타입별 현황**")
            if not df_full.empty and '배정학교타입' in df_full.columns:
                type_stats = df_full['배정학교타입'].value_counts()
                type_stats = type_stats[type_stats.index != '']  # 빈값 제외
                st.bar_chart(type_stats)

        st.divider()

        # 데이터 품질 평가
        st.subheader("🎯 데이터 품질 평가")

        overall_pct = 0
        metrics = []

        if not df_full.empty:
            total = len(df_full)

            # 각 항목별 완성도
            checks = {
                '유형': '유형' in df_full.columns,
                '최종': '최종' in df_full.columns,
                '배정학교': '배정학교' in df_full.columns,
            }

            for name, exists in checks.items():
                if exists:
                    filled = (df_full[name] != '').sum()
                    pct = (filled / total * 100)
                    metrics.append({'항목': name, '완성도(%)': round(pct, 1)})

            if metrics:
                metrics_df = pd.DataFrame(metrics)
                overall_pct = metrics_df['완성도(%)'].mean()

                fig = px.bar(
                    metrics_df,
                    x='항목',
                    y='완성도(%)',
                    title=f"데이터 완성도 (평균: {overall_pct:.1f}%)",
                    range_y=[0, 100],
                    color_discrete_sequence=['#1f77b4']
                )
                st.plotly_chart(fig, use_container_width=True)

                # 품질 등급
                if overall_pct >= 80:
                    st.success(f"✅ 우수 (≥80%): {overall_pct:.1f}%")
                elif overall_pct >= 60:
                    st.info(f"⚠️ 보통 (60-80%): {overall_pct:.1f}%")
                else:
                    st.warning(f"❌ 부족 (<60%): {overall_pct:.1f}%")

# ──────────────────────────────────────────
# TAB 8: 입시 진행 현황 (기존)
# ──────────────────────────────────────────
if tab_progress is not None:
    with tab_progress:
        st.header("📋 입시 진행 현황 (기준: 희망 + 시기 슬롯)")

        progress_df = load_2026_progress_data()

        if not progress_df.empty:
            # 데이터 정제
            progress_df['지원유형'] = progress_df.get('지원유형', '')
            progress_df['최종'] = progress_df.get('최종', '')

            # 통계
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                total = len(progress_df)
                st.metric("📊 총 지원 현황", f"{total}명", "유형별")

            with col2:
                passed = (progress_df['최종'] == '합격').sum()
                pct = (passed / total * 100) if total > 0 else 0
                st.metric("🎉 합격자", f"{passed}명", f"{pct:.1f}%")

            with col3:
                if '지원유형' in progress_df.columns:
                    types = progress_df['지원유형'].nunique()
                    st.metric("🎯 지원 유형", f"{types}가지", "")
                else:
                    st.metric("🎯 지원 유형", "0가지", "")

            with col4:
                if '반' in progress_df.columns:
                    classes = progress_df['반'].nunique()
                    st.metric("🏫 참여 반", f"{classes}반", "")
                else:
                    st.metric("🏫 참여 반", "0반", "")

            st.divider()

            # 지원 유형별 분포
            col1, col2 = st.columns(2)

            with col1:
                if '지원유형' in progress_df.columns:
                    type_dist = progress_df['지원유형'].value_counts().reset_index()
                    type_dist.columns = ['지원유형', '인원']

                    fig = px.pie(
                        type_dist,
                        values='인원',
                        names='지원유형',
                        title="지원 유형별 분포"
                    )
                    st.plotly_chart(fig, use_container_width=True)

            with col2:
                if '최종' in progress_df.columns:
                    final_dist = progress_df['최종'].value_counts().reset_index()
                    final_dist.columns = ['최종결과', '인원']

                    # 합격/불합격만 집계 (빈 값 제외)
                    final_dist = final_dist[final_dist['최종결과'] != '']

                    if not final_dist.empty:
                        fig = px.bar(
                            final_dist,
                            x='최종결과',
                            y='인원',
                            title="최종 결과 분포",
                            color='최종결과',
                            color_discrete_map={'합격': '#1f77b4', '불합격': '#ff7f0e'}
                        )
                        st.plotly_chart(fig, use_container_width=True)

            st.divider()

            # 지원 유형별 상세 통계
            st.subheader("📈 지원 유형별 상세 통계")

            if '지원유형' in progress_df.columns:
                type_stats = progress_df.groupby('지원유형').agg({
                    '성명': 'count',
                }).rename(columns={'성명': '총인원'})

                if '최종' in progress_df.columns:
                    type_stats['합격'] = progress_df[progress_df['최종'] == '합격'].groupby('지원유형').size()
                    type_stats['합격'] = type_stats['합격'].fillna(0).astype(int)
                    type_stats['합격률(%)'] = (type_stats['합격'] / type_stats['총인원'] * 100).round(1)

                st.dataframe(type_stats, use_container_width=True)

            st.divider()

            # 상세 데이터 테이블
            st.subheader("📋 전체 지원 현황")

            # 표시할 컬럼 선택
            display_cols = ['반', '번호', '성명', '성별', '지원유형', '1차', '2차', '최종']
            available_cols = [col for col in display_cols if col in progress_df.columns]

            st.dataframe(
                progress_df[available_cols].sort_values(['반', '번호']),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("⚠️ 입시 진행 현황 데이터가 없습니다.")
            st.info("""
            입시 진행 현황 시트를 생성하려면:
            ```bash
            python generators/build_progress_from_tracking.py 2026
            ```
            """)

# ──────────────────────────────────────────
# 하단 다운로드 영역
# ──────────────────────────────────────────
st.divider()
st.subheader("📥 데이터 다운로드")

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("📊 데이터 (Excel)", use_container_width=True):
        if not df.empty:
            # Excel 파일 생성
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='전체')
                
                if '유형' in df.columns:
                    early_types = ['과학고', '예술계고', '특성화고', '영재고']
                    early_df = df[df['유형'].isin(early_types)]
                    if not early_df.empty:
                        early_df.to_excel(writer, index=False, sheet_name='전기고')
                    
                    late_types = ['자사고', '외고/국제고', '일반고', '비평준화고']
                    late_df = df[df['유형'].isin(late_types)]
                    if not late_df.empty:
                        late_df.to_excel(writer, index=False, sheet_name='후기고')
            
            output.seek(0)
            st.download_button(
                label="💾 다운로드 (Excel)",
                data=output.getvalue(),
                file_name=f"고입현황_{data_year.split()[0]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with col2:
    if st.button("🔄 데이터 새로고침", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

with col3:
    st.markdown(f"**데이터 연도:** {data_year}")
