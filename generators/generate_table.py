#!/usr/bin/env python3
"""
2026 고입 컬러 리포트 생성기 (워터폴 배지형)

데이터 소스: 입시_트래킹 시트
출력:
  reports/전기고_컬러리포트.html  — 영재고/과학고/예술계고/특성화고 (인쇄 최적화)
  reports/후기고_컬러리포트.html  — 자사고/외고·국제고/비평준화고/기타

실행:
  python generate_table.py         # 2026 기본
  python generate_table.py 2026
  python generate_table.py 2025
"""

import sys
import os
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ==========================================
# 설정
# ==========================================
KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_IDS = {
    "2025": "1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI",
    "2026": "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0",
}

# 전기고 → 우선순위순
EARLY_TYPES = ["영재고", "과학고", "예술계고", "특성화고"]
# 후기고
LATE_TYPES  = ["자사고", "외고/국제고", "비평준화고", "기타/대안"]

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "reports")


# ==========================================
# 데이터 수집
# ==========================================
def fetch_from_tracking(year: str):
    """입시_트래킹 시트에서 유형 있는 학생 로드"""
    ss_id = SPREADSHEET_IDS.get(year)
    if not ss_id:
        print(f"❌ 지원하지 않는 연도: {year}")
        sys.exit(1)

    print(f"🔄 {year} 스프레드시트 연결 중...")
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    sht = ss.worksheet("입시_트래킹")
    rows = sht.get_all_values()
    if not rows:
        print("❌ 입시_트래킹 시트가 비어 있습니다.")
        return []

    header = rows[0]

    def col(name):
        try:
            return header.index(name)
        except ValueError:
            return None

    c_cls    = col("반")      or 0
    c_num    = col("번호")    or 1
    c_name   = col("이름")    or 2
    c_gender = col("성별")    or 3
    c_type   = col("유형")
    c_school = col("지원학교")
    c_final  = col("최종")

    if c_type is None:
        print("❌ '유형' 컬럼을 찾을 수 없습니다.")
        return []

    def get(r, idx):
        return r[idx].strip() if idx is not None and idx < len(r) else ""

    # 최종 값 → 워터폴 상태 정규화
    def parse_status(val):
        v = val.strip().lower()
        if "최종합격" in v or (("합격" in v) and "불합" not in v and "1차" not in v and "2차" not in v):
            return "최종합격"
        elif "2차" in v and "합격" in v:
            return "2차합격"
        elif "1차" in v and "합격" in v:
            return "1차합격"
        elif "불합" in v or "fail" in v:
            return "불합격"
        elif v:
            return val.strip()  # 기타 텍스트 그대로
        return "지원중"

    students = []
    for r in rows[1:]:
        if len(r) < 3 or not get(r, c_name):
            continue
        type_val = get(r, c_type)
        if not type_val:
            continue

        students.append({
            "class":  get(r, c_cls),
            "num":    get(r, c_num),
            "name":   get(r, c_name),
            "gender": get(r, c_gender),
            "type":   type_val,
            "school": get(r, c_school) or type_val,
            "status": parse_status(get(r, c_final)),
        })

    print(f"  → {len(students)}명 로드 완료")
    return students


# ==========================================
# HTML 생성
# ==========================================
def make_badge(status: str) -> str:
    base = "inline-flex items-center px-2 py-0.5 rounded text-xs font-bold border"
    mapping = {
        "최종합격": f"{base} bg-green-100 text-green-700 border-green-200",
        "2차합격":  f"{base} bg-purple-100 text-purple-700 border-purple-200",
        "1차합격":  f"{base} bg-blue-100 text-blue-700 border-blue-200",
        "지원중":   f"{base} bg-gray-100 text-gray-500 border-gray-200",
        "불합격":   f"{base} bg-red-50 text-red-400 border-red-100 line-through",
    }
    cls = mapping.get(status, f"{base} bg-slate-50 text-slate-500 border-slate-200")
    emoji = {"최종합격": "🎉 ", "2차합격": "✅ ", "1차합격": "📋 "}.get(status, "")
    return f'<span class="{cls}">{emoji}{status}</span>'


def make_section(title: str, students: list) -> str:
    if not students:
        rows_html = f'<tr><td colspan="6" class="text-center py-6 text-gray-300 text-xs">해당 없음</td></tr>'
    else:
        rows_html = ""
        for i, s in enumerate(students):
            gender_cls = "text-blue-500" if s["gender"] == "남" else "text-red-400"
            rows_html += f"""
            <tr class="hover:bg-slate-50 transition-colors">
                <td class="text-center text-gray-400 font-mono py-2 border-r border-gray-100">{i+1}</td>
                <td class="text-center py-2 border-r border-gray-100 text-gray-600">{s['class']}반</td>
                <td class="text-center py-2 border-r border-gray-100 font-semibold text-gray-800">{s['name']}</td>
                <td class="text-center py-2 border-r border-gray-100 text-xs {gender_cls}">{s['gender']}</td>
                <td class="py-2 border-r border-gray-100 pl-2 font-medium text-gray-700">{s['school']}</td>
                <td class="text-center py-2">{make_badge(s['status'])}</td>
            </tr>"""

    return f"""
    <div class="flex-1 min-w-0 bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <h3 class="text-center font-bold bg-slate-50 py-2.5 border-b border-gray-200 text-slate-700 text-sm">
            {title}
            <span class="ml-1 inline-flex items-center justify-center px-2 py-0.5 rounded-full text-xs font-medium bg-slate-200 text-slate-600">{len(students)}</span>
        </h3>
        <table class="w-full text-xs">
            <thead class="bg-slate-100 border-b border-gray-200 text-slate-500">
                <tr>
                    <th class="py-2 w-7">No</th>
                    <th class="py-2 w-10">반</th>
                    <th class="py-2 w-14">이름</th>
                    <th class="py-2 w-9">성별</th>
                    <th class="py-2 text-left pl-2">지원학교</th>
                    <th class="py-2 w-20">진행상황</th>
                </tr>
            </thead>
            <tbody class="divide-y divide-gray-100">{rows_html}</tbody>
        </table>
    </div>"""


def generate_html(sections: list, title: str, filepath: str, year: str) -> str:
    """
    sections: [ (section_title, [student, ...]), ... ]
    """
    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    content_html = ""
    total = 0
    pass_cnt = 0
    for sec_title, students in sections:
        content_html += make_section(sec_title, students)
        content_html += '<div class="w-4 shrink-0"></div>'
        total += len(students)
        pass_cnt += sum(1 for s in students if s["status"] == "최종합격")

    pass_badge = f'<span class="text-green-600 font-bold ml-3">🎉 {pass_cnt}명 합격</span>' if pass_cnt else ""

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet"
          href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css"/>
    <style>
        body {{ font-family: "Pretendard Variable", Pretendard, -apple-system, sans-serif;
               -webkit-print-color-adjust: exact; }}
        @media print {{
            @page {{ size: landscape; margin: 8mm; }}
            .no-print {{ display: none !important; }}
            body {{ background: white; padding: 0; }}
            .shadow-sm {{ box-shadow: none; }}
        }}
    </style>
</head>
<body class="p-8 bg-slate-50 min-h-screen">
    <div class="max-w-[297mm] mx-auto">
        <header class="flex justify-between items-end mb-6 border-b border-slate-300 pb-4">
            <div>
                <h1 class="text-2xl font-extrabold text-slate-800">{title}</h1>
                <div class="flex gap-3 mt-1.5 text-xs font-medium text-slate-500">
                    <span>총 <b class="text-slate-700">{total}</b>명{pass_badge}</span>
                </div>
                <div class="flex gap-3 mt-1 text-[11px] text-slate-400">
                    <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-blue-400"></span>1차합격</span>
                    <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-purple-400"></span>2차합격</span>
                    <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-green-500"></span>최종합격</span>
                    <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-red-300"></span>불합격</span>
                </div>
            </div>
            <div class="text-right">
                <p class="text-xs text-slate-400 mb-2 font-mono">업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
                <button onclick="window.print()"
                        class="no-print bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-bold shadow">
                    🖨 인쇄하기
                </button>
            </div>
        </header>

        <div class="flex flex-row items-start gap-0">
            {content_html}
        </div>

        <footer class="mt-6 text-center border-t border-slate-200 pt-3">
            <p class="text-[10px] text-slate-400">
                {year}학년도 목일중학교 진학현황 컬러리포트 &nbsp;|&nbsp;
                <a href="전기고_컬러리포트.html" class="underline">전기고</a> |
                <a href="후기고_컬러리포트.html" class="underline">후기고</a> |
                <a href="전체_현황.html" class="underline">전체</a>
            </p>
        </footer>
    </div>
</body>
</html>"""

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)

    abs_path = os.path.abspath(filepath)
    print(f"  ✅ {os.path.basename(filepath)}  ({total}명)")
    return abs_path


# ==========================================
# 실행
# ==========================================
def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    students = fetch_from_tracking(year)
    if not students:
        print("⚠️ 데이터 없음.")
        return

    by_type = {}
    for s in students:
        by_type.setdefault(s["type"], []).append(s)

    out = lambda name: os.path.join(OUTPUT_DIR, name)
    print("\n[HTML 생성]")

    # 전기고
    early_sections = [(t, by_type.get(t, [])) for t in EARLY_TYPES if t in by_type]
    if early_sections:
        generate_html(early_sections, f"{year} 전기고 전형 진행 현황", out("전기고_컬러리포트.html"), year)

    # 후기고
    late_sections = [(t, by_type.get(t, [])) for t in LATE_TYPES if t in by_type]
    # 목록에 없는 기타 후기 유형
    known = set(EARLY_TYPES + LATE_TYPES)
    extra = [(t, v) for t, v in by_type.items() if t not in known]
    late_sections.extend(extra)

    if late_sections:
        generate_html(late_sections, f"{year} 후기고 전형 진행 현황", out("후기고_컬러리포트.html"), year)

    print(f"\n완료! reports/ 폴더에 컬러리포트 생성됨.")

    try:
        import webbrowser
        webbrowser.open(f"file://{os.path.abspath(out('전기고_컬러리포트.html'))}")
    except Exception:
        pass


if __name__ == "__main__":
    main()
