import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import re
import os
from datetime import datetime

# ==========================================
# 1. 설정 정보
# ==========================================
KEY_FILE = 'service_key.json'
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI/edit?gid=294818561#gid=294818561'

# 컬럼 인덱스 (A=0 기준)
COL = {
    'GIFTED': 7,       # H: 영재고
    'SCIENCE': 8,      # I: 과학고
    'ARTS': 9,         # J: 예술고
    'MEISTER': 10,     # K: 특성화고(교명)
    'DEPT': 11,        # L: 특성화고(학과)
    'JASA': 12,        # M: 자사고
    'FOREIGN': 13,     # N: 외고/국제고
    'GENERAL': 14,     # O: 일반고
    'ETC': 15,         # P: 기타
    'RES_GIFTED': 20,  # U: 영재고 합불
    'RES_EARLY': 21,   # V: 전기고 합불
    'RES_LATE': 22     # W: 후기고 합불
}

def get_data_with_waterfall():
    print("🔄 데이터 수집 및 상태별 배지 로직 적용 중...")
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
    client = gspread.authorize(creds)
    doc = client.open_by_url(SHEET_URL)
    
    early_report = {'gifted': [], 'science': [], 'arts': [], 'meister': []}
    late_report = {'jasa': [], 'foreign': [], 'etc': []}
    
    target_pattern = re.compile(r"진학희망 및 지원유형 조사\(3\d{2}\)_Sheet1")
    
    for sht in doc.worksheets():
        if target_pattern.search(sht.title):
            rows = sht.get_all_values()
            if len(rows) < 3: continue
            
            for r in rows[2:]:
                if len(r) < 3 or not r[2].strip(): continue
                row = r + [''] * (30 - len(r))
                base = {'class': row[0], 'name': row[2], 'gender': row[3]}
                history_note = []
                
                # --- 1. 영재고 ---
                sch = row[COL['GIFTED']].strip()
                res = row[COL['RES_GIFTED']].strip()
                
                if sch and sch != 'nan':
                    sch_name = sch if sch not in ['O','o'] else "영재학교"
                    # 상태 판별
                    if "합격" in res and "불합" not in res: status = "최종합격"
                    elif "2차" in res: status = "2차합격"
                    elif "1차" in res: status = "1차합격"
                    elif "불합" in res: status = "불합격"
                    else: status = "지원" # 기본값

                    if status == "최종합격":
                        early_report['gifted'].append({**base, 'school': sch_name, 'status': status, 'note': ''})
                        continue
                    elif status == "불합격":
                        history_note.append("영재불합")
                    else: # 진행중 (1차, 2차, 지원)
                        early_report['gifted'].append({**base, 'school': sch_name, 'status': status, 'note': ''})
                        continue

                # --- 2. 전기고 ---
                sch_sci = row[COL['SCIENCE']].strip()
                sch_art = row[COL['ARTS']].strip()
                sch_mei = row[COL['MEISTER']].strip()
                res_early = row[COL['RES_EARLY']].strip()
                
                if sch_sci or sch_art or sch_mei:
                    if "합격" in res_early and "불합" not in res_early: status = "최종합격"
                    elif "2차" in res_early: status = "2차합격"
                    elif "1차" in res_early: status = "1차합격"
                    elif "불합" in res_early: status = "불합격"
                    else: status = "지원"

                    final_note = "/".join(history_note)

                    if sch_sci:
                        sch_name = sch_sci if sch_sci not in ['O','o'] else "과학고"
                        if status != "불합격":
                            early_report['science'].append({**base, 'school': sch_name, 'status': status, 'note': final_note})
                            continue
                        else: history_note.append("과고불합")
                    
                    elif sch_art:
                        sch_name = sch_art if sch_art not in ['O','o'] else "예술고"
                        if status != "불합격":
                            early_report['arts'].append({**base, 'school': sch_name, 'status': status, 'note': final_note})
                            continue
                        else: history_note.append("예고불합")

                    elif sch_mei:
                        sch_name = sch_mei if sch_mei not in ['O','o'] else "특성화고"
                        dept = row[COL['DEPT']].strip()
                        if status != "불합격":
                            early_report['meister'].append({**base, 'school': sch_name, 'dept': dept, 'status': status, 'note': final_note})
                            continue
                        else: history_note.append("특성불합")

                # --- 3. 후기고 ---
                sch_jasa = row[COL['JASA']].strip()
                sch_for = row[COL['FOREIGN']].strip()
                sch_etc = row[COL['ETC']].strip()
                res_late = row[COL['RES_LATE']].strip()
                
                if "합격" in res_late and "불합" not in res_late: status = "최종합격"
                elif "1차" in res_late or "면접" in res_late: status = "1차합격"
                elif "불합" in res_late: status = "불합격"
                else: status = "지원"
                
                final_note = "/".join(history_note)

                if sch_jasa:
                    sch_name = sch_jasa if sch_jasa not in ['O','o'] else "자사고"
                    late_report['jasa'].append({**base, 'school': sch_name, 'status': status, 'note': final_note})
                elif sch_for:
                    sch_name = sch_for if sch_for not in ['O','o'] else "외고/국제고"
                    late_report['foreign'].append({**base, 'school': sch_name, 'status': status, 'note': final_note})
                elif sch_etc:
                    sch_name = sch_etc if sch_etc not in ['O','o'] else "기타"
                    late_report['etc'].append({**base, 'school': sch_name, 'status': status, 'note': final_note})

    return early_report, late_report

# ==========================================
# 2. HTML 생성 (컬러 배지 적용)
# ==========================================
def generate_html_with_badges(data_dict, title, filename, mode='early'):
    
    # [핵심] 상태별 배지 디자인 함수
    def make_badge(status):
        # Tailwind CSS 클래스 조합
        base_cls = "inline-flex items-center px-2 py-0.5 rounded text-xs font-bold border"
        
        if "최종합격" in status or "합격" == status:
            # 초록색 (Green)
            return f'<span class="{base_cls} bg-green-100 text-green-700 border-green-200">🎉 최종합격</span>'
        elif "2차" in status:
            # 보라색 (Purple) - 최종 직전
            return f'<span class="{base_cls} bg-purple-100 text-purple-700 border-purple-200">2차 합격</span>'
        elif "1차" in status:
            # 파란색 (Blue) - 시작
            return f'<span class="{base_cls} bg-blue-100 text-blue-700 border-blue-200">1차 합격</span>'
        elif "지원" in status:
            # 회색 (Gray) - 대기중
            return f'<span class="{base_cls} bg-gray-100 text-gray-600 border-gray-200">지원 완료</span>'
        elif "불합격" in status:
            # 붉은색 (Red) + 취소선
            return f'<span class="{base_cls} bg-red-50 text-red-500 border-red-100 line-through">불합격</span>'
        else:
            return f'<span class="text-xs text-gray-400">{status}</span>'

    def make_table(section_title, data, cols):
        rows = ""
        if not data:
            rows = f'<tr><td colspan="{len(cols)+5}" class="text-center py-8 text-gray-300">해당 없음</td></tr>'
        
        for idx, s in enumerate(data):
            badge = make_badge(s['status'])
            note_html = f'<div class="text-[10px] text-gray-400 mt-0.5">({s["note"]})</div>' if s['note'] else ""
            
            # 학교명이 길어질 경우를 대비해 truncate 적용 가능
            school_display = s.get('school','-')
            
            rows += f"""
            <tr class="hover:bg-gray-50 border-b border-gray-200 transition-colors">
                <td class="text-center border-r border-gray-200 py-2.5 font-mono text-gray-500">{idx+1}</td>
                <td class="text-center border-r border-gray-200 py-2.5">{s['class']}</td>
                <td class="text-center border-r border-gray-200 py-2.5 font-semibold text-gray-700">{s['name']}</td>
                <td class="text-center border-r border-gray-200 py-2.5 text-xs text-gray-500">{s['gender']}</td>
                <td class="text-center border-r border-gray-200 py-2.5">
                    <span class="font-medium">{school_display}</span>
                    {note_html}
                </td>
                {'<td class="text-center border-r border-gray-200 py-2.5 text-xs text-gray-600">' + s.get('dept','-') + '</td>' if '학과' in cols else ''}
                <td class="text-center py-2.5">{badge}</td>
            </tr>
            """

        return f"""
        <div class="flex-1 min-w-0 bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            <h3 class="text-center font-bold bg-slate-50 py-3 border-b border-gray-200 text-slate-700">
                {section_title} 
                <span class="ml-1 inline-flex items-center justify-center px-2 py-0.5 rounded-full text-xs font-medium bg-slate-200 text-slate-600">{len(data)}</span>
            </h3>
            <table class="w-full text-xs">
                <thead class="bg-slate-100 border-b border-gray-200 text-slate-500 uppercase tracking-wider">
                    <tr>
                        <th class="py-2 w-8 font-semibold">No</th>
                        <th class="py-2 w-10 font-semibold">반</th>
                        <th class="py-2 w-16 font-semibold">이름</th>
                        <th class="py-2 w-10 font-semibold">성별</th>
                        <th class="py-2 font-semibold">지원학교</th>
                        {'<th class="py-2 w-24 font-semibold">학과</th>' if '학과' in cols else ''}
                        <th class="py-2 w-24 font-semibold">진행상황</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-gray-100">{rows}</tbody>
            </table>
        </div>
        """

    content = ""
    if mode == 'early':
        content += make_table("영재학교", data_dict['gifted'], [])
        content += '<div class="w-6"></div>'
        content += make_table("과학고/예술고", data_dict['science'] + data_dict['arts'], [])
        content += '<div class="w-6"></div>'
        content += make_table("특성화/마이스터고", data_dict['meister'], ['학과'])
    else:
        content += make_table("자사고", data_dict['jasa'], [])
        content += '<div class="w-6"></div>'
        content += make_table("외고/국제고", data_dict['foreign'], [])
        content += '<div class="w-6"></div>'
        content += make_table("기타/비평준", data_dict['etc'], [])

    full_html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <title>{title}</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
        <style>
            body {{ font-family: "Pretendard Variable", Pretendard, sans-serif; -webkit-print-color-adjust: exact; }}
            @media print {{ 
                @page {{ size: landscape; margin: 10mm; }} 
                .no-print {{ display:none !important; }}
                body {{ background: white; padding: 0; }}
                .shadow-sm {{ box-shadow: none; }}
            }}
        </style>
    </head>
    <body class="p-8 bg-slate-50 min-h-screen">
        <div class="max-w-[297mm] mx-auto">
            <header class="flex justify-between items-end mb-8 border-b border-slate-300 pb-4">
                <div>
                    <h1 class="text-3xl font-extrabold text-slate-800 tracking-tight">{title}</h1>
                    <div class="flex gap-3 mt-2 text-xs font-medium text-slate-500">
                        <span class="flex items-center"><span class="w-2 h-2 rounded-full bg-blue-400 mr-1.5"></span>1차합격</span>
                        <span class="flex items-center"><span class="w-2 h-2 rounded-full bg-purple-400 mr-1.5"></span>2차합격</span>
                        <span class="flex items-center"><span class="w-2 h-2 rounded-full bg-green-500 mr-1.5"></span>최종합격</span>
                    </div>
                </div>
                <div class="text-right">
                    <p class="text-xs text-slate-400 mb-2 font-mono">업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
                    <button onclick="window.print()" class="no-print bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-lg text-sm font-bold transition shadow-lg flex items-center gap-2 ml-auto">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
                        인쇄하기
                    </button>
                </div>
            </header>
            
            <div class="flex flex-row items-start justify-between gap-6">
                {content}
            </div>
            
            <footer class="mt-8 text-center border-t border-slate-200 pt-4">
                <p class="text-[11px] text-slate-400">
                    * 학교명 하단 괄호 안의 내용은 이전 단계 전형 탈락 이력을 나타냅니다. (예: 영재불합) <br>
                    * 본 문서는 행정 및 입시 관리 목적으로 생성되었습니다.
                </p>
            </footer>
        </div>
    </body>
    </html>
    """
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(full_html)
    print(f"✅ 리포트 생성 완료: {filename}")

if __name__ == "__main__":
    early, late = get_data_with_waterfall()
    
    output_dir = "reports"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    generate_html_with_badges(early, "2025학년도 전기고 전형 진행 현황", os.path.join(output_dir, "목일중_전기고_컬러리포트.html"), mode='early')
    generate_html_with_badges(late, "2025학년도 후기고 전형 진행 현황", os.path.join(output_dir, "목일중_후기고_컬러리포트.html"), mode='late')