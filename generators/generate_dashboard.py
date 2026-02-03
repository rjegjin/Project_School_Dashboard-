import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import re
import os
from datetime import datetime
from typing import List, Dict, Tuple, Any

# ==========================================
# 1. 설정 정보
# ==========================================
KEY_FILE = 'service_key.json'
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI/edit?gid=294818561#gid=294818561'

# 생성될 파일명
OUTPUT_DIR = 'reports'
OUTPUT_EARLY_HTML = os.path.join(OUTPUT_DIR, '목일중_전기고_진학현황.html')
OUTPUT_LATE_HTML = os.path.join(OUTPUT_DIR, '목일중_후기고_진학현황.html')

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# ==========================================
# 2. 데이터 가져오기 및 처리
# ==========================================
def fetch_all_data() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    구글 시트에서 데이터를 가져와 전기고/후기고 지원자 리스트로 분리하여 반환합니다.
    """
    print("🔄 구글 시트에 연결 중입니다...")
    
    try:
        creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        doc = client.open_by_url(SHEET_URL)
    except Exception as e:
        print(f"❌ 구글 시트 연결 실패: {e}")
        return [], []
    
    worksheets = doc.worksheets()
    
    # 정규표현식: "진학희망 및 지원유형 조사(3"으로 시작하고 "_Sheet1"으로 끝나는 시트 찾기
    # 예: 진학희망 및 지원유형 조사(303)_Sheet1
    target_pattern = re.compile(r"진학희망 및 지원유형 조사\(3\d{2}\)_Sheet1")
    
    early_students = []
    late_students = []
    
    for sht in worksheets:
        if target_pattern.search(sht.title):
            print(f"📑 데이터 수집 중: {sht.title}")
            
            # 전체 데이터 가져오기
            try:
                rows = sht.get_all_values()
            except Exception as e:
                print(f"⚠️ 시트 데이터 읽기 실패 ({sht.title}): {e}")
                continue
            
            # 데이터 유효성 검사 (행 개수 부족 시 패스)
            if len(rows) < 3: continue
            
            # 3행(Index 2)부터 학생 데이터 시작
            for r in rows[2:]:
                # 이름이 없으면 빈 행으로 간주
                if len(r) < 3 or not r[2].strip(): continue
                
                # 데이터 파싱 (CSV 구조 기반 인덱스 매핑)
                # 0:반, 1:번호, 2:성명, 3:성별
                # 7:영재, 8:과학, 9:예술, 10:특성화, 11:특성화학과
                # 12:자사, 13:외고, 14:일반, 15:기타
                
                # 안전한 인덱싱을 위해 길이 확장
                row = r + [''] * (25 - len(r))
                
                info = {
                    'class': row[0],
                    'num': row[1],
                    'name': row[2],
                    'gender': row[3],
                    'result': '',   # 합불 여부 (추후 확장을 위해 비워둠 or 맨 뒤 열 확인)
                    'school': '',
                    'dept': '',     # 학과
                    'type': ''
                }
                
                # 합불 여부 확인 (맨 뒤쪽 열이나 비고란 활용, 여기서는 예시로 맨 뒤쪽 스캔)
                # "합격"이라는 단어가 있는 열을 찾음
                for cell in row[16:]: 
                    if "합격" in str(cell): info['result'] = "합격"
                    elif "불합격" in str(cell): info['result'] = "불합격"

                # --- [전기고 판별] ---
                is_early = False
                
                # 1. 영재고 (7)
                if row[7].strip():
                    is_early = True; info['type'] = '영재고'; info['school'] = _clean_school_name(row[7], '영재고')
                # 2. 과학고 (8)
                elif row[8].strip():
                    is_early = True; info['type'] = '과학고'; info['school'] = _clean_school_name(row[8], '과학고')
                # 3. 예술고 (9)
                elif row[9].strip():
                    is_early = True; info['type'] = '예술고'; info['school'] = _clean_school_name(row[9], '예술고')
                # 4. 특성화고 (10)
                elif row[10].strip():
                    is_early = True; info['type'] = '특성화고'
                    info['school'] = _clean_school_name(row[10], '특성화고')
                    info['dept'] = row[11].strip() # 학과
                
                if is_early:
                    early_students.append(info)
                    continue # 전기에 속하면 후기는 체크 안 함 (우선순위)

                # --- [후기고 판별] ---
                is_late = False
                
                # 1. 자사고 (12)
                if row[12].strip():
                    is_late = True; info['type'] = '자사고'; info['school'] = _clean_school_name(row[12], '자사고')
                # 2. 외고/국제고 (13)
                elif row[13].strip():
                    is_late = True; info['type'] = '외고/국제고'; info['school'] = _clean_school_name(row[13], '외고/국제고')
                # 3. 일반고 (14) - 보통 일반고는 명단 안 만들지만 데이터 있으면 수집
                elif row[14].strip():
                    is_late = True; info['type'] = '일반고'; info['school'] = _clean_school_name(row[14], '일반고')
                # 4. 기타/대안 (15)
                elif row[15].strip():
                    is_late = True; info['type'] = '대안/기타'; info['school'] = _clean_school_name(row[15], '대안학교')
                
                if is_late:
                    late_students.append(info)
                    
    return early_students, late_students

def _clean_school_name(text: str, default_type: str) -> str:
    """'O'나 '○'만 있으면 기본 유형명을, 텍스트가 있으면 텍스트를 반환"""
    text = text.strip()
    if text in ['O', 'o', '○', '0', '']:
        return default_type
    return text

# ==========================================
# 3. HTML 생성 (카드형 대시보드)
# ==========================================
def generate_html(student_list: List[Dict[str, Any]], title: str, filename: str) -> None:
    cards_html = ""
    
    # 통계 계산
    total_count = len(student_list)
    pass_count = sum(1 for s in student_list if s['result'] == '합격')
    
    for s in student_list:
        # 디자인 요소 결정
        gender_color = "text-blue-600 bg-blue-50" if s['gender'] == '남' else "text-red-600 bg-red-50"
        
        # 상태 뱃지 (합격/불합격/지원중)
        if s['result'] == '합격':
            status_badge = '<span class="px-2 py-1 rounded bg-green-100 text-green-700 text-xs font-bold">🎉 합격</span>'
            card_border = "border-green-400 ring-2 ring-green-100"
        elif s['result'] == '불합격':
            status_badge = '<span class="px-2 py-1 rounded bg-gray-200 text-gray-600 text-xs font-bold">불합격</span>'
            card_border = "border-gray-200 opacity-70"
        else:
            status_badge = '<span class="px-2 py-1 rounded bg-indigo-50 text-indigo-600 text-xs font-bold">지원중</span>'
            card_border = "border-gray-200 hover:border-indigo-300 hover:shadow-lg"

        # 학과 표시 (있으면)
        dept_html = f'<div class="text-xs text-gray-500 mt-1">📌 {s["dept"]}</div>' if s['dept'] else ''
        
        card = f"""
        <div class="bg-white rounded-xl p-5 border {card_border} transition-all duration-300 shadow-sm flex flex-col justify-between">
            <div>
                <div class="flex justify-between items-start mb-3">
                    <div class="flex flex-col">
                        <span class="text-xs font-bold text-gray-400 mb-1">{s['class']}반 {s['num']}번</span>
                        <h3 class="text-lg font-extrabold text-gray-800">{s['name']}</h3>
                    </div>
                    <span class="px-2 py-1 rounded text-xs font-bold {gender_color}">{s['gender']}</span>
                </div>
                
                <div class="mb-4">
                    <span class="inline-block px-2 py-0.5 rounded text-xs font-medium bg-gray-100 text-gray-600 mb-2">{s['type']}</span>
                    <div class="text-gray-900 font-bold text-md leading-tight">{s['school']}</div>
                    {dept_html}
                </div>
            </div>
            
            <div class="pt-3 border-t border-gray-100 flex justify-between items-center">
                {status_badge}
            </div>
        </div>
        """
        cards_html += card

    # 전체 HTML 템플릿
    full_html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{title}</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <link rel="stylesheet" as="style" crossorigin href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
        <style>
            body {{ font-family: "Pretendard Variable", Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif; }}
        </style>
    </head>
    <body class="bg-slate-50 min-h-screen p-6 md:p-12">
        <div class="max-w-7xl mx-auto">
            <header class="mb-10 flex flex-col md:flex-row md:items-end justify-between gap-4">
                <div>
                    <h1 class="text-3xl md:text-4xl font-black text-slate-800 mb-2">{title}</h1>
                    <p class="text-slate-500 font-medium">
                        총 <span class="text-indigo-600 font-bold">{total_count}</span>명 지원 
                        {' | <span class="text-green-600 font-bold">🎉 ' + str(pass_count) + '명 합격</span>' if pass_count > 0 else ''}
                    </p>
                </div>
                <div class="text-right text-xs text-gray-400">
                    업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}
                </div>
            </header>

            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
                {cards_html}
            </div>
            
            <footer class="mt-12 text-center text-gray-400 text-sm">
                2025학년도 목일중학교 진학현황 대시보드
            </footer>
        </div>
    </body>
    </html>
    """
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(full_html)
    print(f"✅ 파일 생성 완료: {filename}")

# ==========================================
# 4. 실행
# ==========================================
if __name__ == "__main__":
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    early_list, late_list = fetch_all_data()
    
    if early_list:
        generate_html(early_list, "2025학년도 전기고 지원 현황", OUTPUT_EARLY_HTML)
    else:
        print("⚠️ 전기고 지원자가 없습니다.")

    if late_list:
        generate_html(late_list, "2025학년도 후기고 지원 현황", OUTPUT_LATE_HTML)
    else:
        print("⚠️ 후기고 지원자가 없습니다.")