#!/usr/bin/env python3
"""
특별전형 현황 HTML 리포트 생성기

데이터 소스:
  특별전형_트래킹 시트  — O/X 특별전형 여부
  입시_트래킹 시트      — 유형(전기/후기고), 지원학교

출력:
  reports/특별전형_현황.html

실행:
  python generate_special_report.py         # 2026 기본
  python generate_special_report.py 2026
"""

import sys
import os
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_IDS = {
    "2025": "1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI",
    "2026": "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0",
}

SPECIAL_NAMES = ["사회통합전형", "특례", "보훈", "쌍둥이", "학폭", "교직원자녀", "장애", "다자녀(3인+)"]

# 전기고 유형
EARLY_TYPES = {"영재고", "과학고", "예술계고", "특성화고"}
# 후기고
LATE_TYPES = {"자사고", "외고/국제고", "비평준화고"}

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "reports")

# 카테고리별 색상 (Tailwind)
CATEGORY_COLORS = {
    "사회통합전형": ("bg-sky-100",    "text-sky-700",    "border-sky-200",    "bg-sky-500"),
    "특례":        ("bg-violet-100", "text-violet-700", "border-violet-200", "bg-violet-500"),
    "보훈":        ("bg-amber-100",  "text-amber-700",  "border-amber-200",  "bg-amber-500"),
    "쌍둥이":      ("bg-pink-100",   "text-pink-700",   "border-pink-200",   "bg-pink-500"),
    "학폭":        ("bg-red-100",    "text-red-700",    "border-red-200",    "bg-red-500"),
    "교직원자녀":  ("bg-green-100",  "text-green-700",  "border-green-200",  "bg-green-500"),
    "장애":        ("bg-orange-100", "text-orange-700", "border-orange-200", "bg-orange-500"),
    "다자녀(3인+)": ("bg-teal-100",  "text-teal-700",   "border-teal-200",   "bg-teal-500"),
}


def school_group(type_val: str) -> str:
    if type_val in EARLY_TYPES:
        return "전기고"
    elif type_val in LATE_TYPES:
        return "후기고"
    elif type_val == "일반고":
        return "일반고"
    elif type_val:
        return "기타"
    return "미분류"


def fetch_data(year: str):
    ss_id = SPREADSHEET_IDS.get(year)
    if not ss_id:
        print(f"❌ 지원하지 않는 연도: {year}")
        sys.exit(1)

    print(f"🔄 {year} 스프레드시트 연결 중...")
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(ss_id)

    # 1) 특별전형_트래킹 로드
    try:
        special_sht = ss.worksheet("특별전형_트래킹")
    except gspread.exceptions.WorksheetNotFound:
        print("❌ 특별전형_트래킹 시트를 찾을 수 없습니다.")
        print("   먼저 sync_special_to_tracking.py 를 실행하세요.")
        sys.exit(1)

    special_rows = special_sht.get_all_values()
    if not special_rows:
        print("❌ 특별전형_트래킹 시트가 비어 있습니다.")
        sys.exit(1)

    special_header = special_rows[0]

    def col(name):
        try:
            return special_header.index(name)
        except ValueError:
            return None

    c_cls    = col("반")    or 0
    c_num    = col("번호")  or 1
    c_name   = col("이름")  or 2
    c_gender = col("성별")  or 3

    special_data = {}  # (반, 번호) → {name, gender, categories: [...]}
    for r in special_rows[1:]:
        if len(r) < 3 or not r[c_name].strip():
            continue
        key = (r[c_cls].strip(), r[c_num].strip())
        cats = []
        for cat in SPECIAL_NAMES:
            ci = col(cat)
            if ci is not None and ci < len(r) and r[ci].strip().upper() == "O":
                cats.append(cat)
        special_data[key] = {
            "cls":    r[c_cls].strip(),
            "num":    r[c_num].strip(),
            "name":   r[c_name].strip(),
            "gender": r[c_gender].strip() if c_gender < len(r) else "",
            "cats":   cats,
        }

    print(f"  → 특별전형_트래킹: {len(special_data)}명 로드")

    # 2) 입시_트래킹 로드 (유형, 지원학교, 최종)
    try:
        tracking_sht = ss.worksheet("입시_트래킹")
        tracking_rows = tracking_sht.get_all_values()
    except Exception:
        tracking_rows = []

    tracking_map = {}  # (반, 번호) → {type, school, status}
    if tracking_rows:
        th = tracking_rows[0]
        tc_cls    = th.index("반")    if "반"    in th else 0
        tc_num    = th.index("번호")  if "번호"  in th else 1
        tc_type   = th.index("유형")  if "유형"  in th else None
        tc_school = th.index("지원학교") if "지원학교" in th else None
        tc_final  = th.index("최종")  if "최종"  in th else None

        for r in tracking_rows[1:]:
            if len(r) < 3 or not r[tc_num if tc_num < len(r) else 1].strip():
                continue
            key = (r[tc_cls].strip(), r[tc_num].strip())
            tracking_map[key] = {
                "type":   r[tc_type].strip()   if tc_type   is not None and tc_type   < len(r) else "",
                "school": r[tc_school].strip() if tc_school is not None and tc_school < len(r) else "",
                "status": r[tc_final].strip()  if tc_final  is not None and tc_final  < len(r) else "",
            }

    print(f"  → 입시_트래킹: {len(tracking_map)}명 로드")

    # 3) 합치기
    students = []
    for key, sp in special_data.items():
        tr = tracking_map.get(key, {})
        students.append({
            **sp,
            "type":   tr.get("type", ""),
            "school": tr.get("school", ""),
            "status": tr.get("status", ""),
            "group":  school_group(tr.get("type", "")),
        })

    return students


# ── HTML 생성 ──────────────────────────────────────────────────────────────────

def make_summary_cards(students: list) -> str:
    all_with_special = [s for s in students if s["cats"]]
    total_special = len(all_with_special)

    cards_html = f"""
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
        <div class="bg-white rounded-xl border border-gray-200 shadow-sm p-4 text-center">
            <div class="text-3xl font-extrabold text-slate-800">{total_special}</div>
            <div class="text-xs text-slate-500 mt-1">특별전형 해당 학생</div>
        </div>"""

    for cat in SPECIAL_NAMES:
        cnt = sum(1 for s in students if cat in s["cats"])
        if cnt == 0:
            continue
        bg, text, border, _ = CATEGORY_COLORS[cat]
        cards_html += f"""
        <div class="bg-white rounded-xl border {border} shadow-sm p-4 text-center">
            <div class="text-2xl font-extrabold {text}">{cnt}</div>
            <div class="text-xs text-slate-500 mt-1">{cat}</div>
        </div>"""

    cards_html += "\n    </div>"
    return cards_html


def make_cross_matrix(students: list) -> str:
    groups = ["전기고", "후기고", "일반고", "기타", "미분류"]
    # 집계
    matrix = {}
    for cat in SPECIAL_NAMES:
        matrix[cat] = {}
        for g in groups:
            matrix[cat][g] = sum(
                1 for s in students
                if cat in s["cats"] and s["group"] == g
            )

    col_totals = {g: sum(matrix[c][g] for c in SPECIAL_NAMES) for g in groups}
    row_totals = {cat: sum(matrix[cat][g] for g in groups) for cat in SPECIAL_NAMES}

    # 사용된 그룹만 표시
    active_groups = [g for g in groups if col_totals[g] > 0]
    active_cats   = [c for c in SPECIAL_NAMES if row_totals[c] > 0]

    if not active_cats:
        return '<p class="text-gray-400 text-sm text-center py-8">데이터 없음</p>'

    # 헤더
    th_cells = "".join(
        f'<th class="py-2 px-3 text-center text-xs font-semibold text-slate-500">{g}</th>'
        for g in active_groups
    )
    th_cells += '<th class="py-2 px-3 text-center text-xs font-bold text-slate-700">합계</th>'

    rows_html = ""
    for cat in active_cats:
        _, text, border, _ = CATEGORY_COLORS[cat]
        total = row_totals[cat]
        cells = ""
        for g in active_groups:
            v = matrix[cat][g]
            cell_cls = f"text-center py-2 px-3 text-sm {text} font-semibold" if v > 0 else "text-center py-2 px-3 text-sm text-gray-300"
            cells += f'<td class="{cell_cls}">{v if v > 0 else "—"}</td>'
        cells += f'<td class="text-center py-2 px-3 text-sm font-bold text-slate-700">{total}</td>'
        rows_html += f"""
        <tr class="border-b border-gray-100 hover:bg-slate-50">
            <td class="py-2 px-3 text-sm font-semibold {text} border-r border-gray-100">{cat}</td>
            {cells}
        </tr>"""

    # 합계 행
    total_cells = "".join(
        f'<td class="py-2 px-3 text-center text-sm font-bold text-slate-600">{col_totals[g]}</td>'
        for g in active_groups
    )
    grand_total = sum(row_totals[c] for c in active_cats)
    total_cells += f'<td class="py-2 px-3 text-center text-sm font-bold text-slate-800">{grand_total}</td>'

    return f"""
    <div class="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden mb-6">
        <h3 class="text-sm font-bold text-slate-700 bg-slate-50 px-4 py-2.5 border-b border-gray-200">
            특별전형 × 지원계열 교차표
        </h3>
        <div class="overflow-x-auto">
            <table class="w-full text-xs">
                <thead class="bg-slate-100 border-b border-gray-200">
                    <tr>
                        <th class="py-2 px-3 text-left text-xs font-semibold text-slate-500 border-r border-gray-200">전형</th>
                        {th_cells}
                    </tr>
                </thead>
                <tbody class="divide-y divide-gray-50">{rows_html}</tbody>
                <tfoot class="bg-slate-50 border-t border-gray-200">
                    <tr>
                        <td class="py-2 px-3 text-sm font-bold text-slate-600 border-r border-gray-200">합계</td>
                        {total_cells}
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>"""


def make_student_table(students: list) -> str:
    special_students = [s for s in students if s["cats"]]
    if not special_students:
        return '<p class="text-gray-400 text-sm text-center py-8">해당 학생 없음</p>'

    rows_html = ""
    for i, s in enumerate(sorted(special_students, key=lambda x: (x["cls"], x["num"]))):
        gender_cls = "text-blue-500" if s["gender"] == "남" else "text-red-400"
        cat_badges = ""
        for cat in s["cats"]:
            _, text, border, _ = CATEGORY_COLORS[cat]
            cat_badges += f'<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold border {text} {border} mr-1">{cat}</span>'

        status_cls = {
            "최종합격": "text-green-600 font-bold",
            "불합격":   "text-red-400",
        }.get(s["status"], "text-gray-500")

        rows_html += f"""
        <tr class="hover:bg-slate-50 transition-colors" data-cats="{' '.join(s['cats'])}">
            <td class="text-center text-gray-400 font-mono py-2 border-r border-gray-100 text-xs">{i+1}</td>
            <td class="text-center py-2 border-r border-gray-100 text-xs text-gray-600">{s['cls']}반</td>
            <td class="text-center py-2 border-r border-gray-100 text-xs font-semibold text-gray-800">{s['name']}</td>
            <td class="text-center py-2 border-r border-gray-100 text-xs {gender_cls}">{s['gender']}</td>
            <td class="py-2 border-r border-gray-100 pl-2 text-xs">{cat_badges}</td>
            <td class="text-center py-2 border-r border-gray-100 text-xs text-gray-600">{s['group']}</td>
            <td class="py-2 pl-2 text-xs text-gray-600">{s['school'] or s['type'] or '—'}</td>
            <td class="text-center py-2 text-xs {status_cls}">{s['status'] or '—'}</td>
        </tr>"""

    # 필터 버튼
    filter_btns = """<button onclick="filterTable('')"
        class="filter-btn px-3 py-1 rounded-full text-xs font-bold border border-gray-300 text-gray-600 hover:bg-gray-100 active" data-cat="">
        전체
    </button>"""
    for cat in SPECIAL_NAMES:
        cnt = sum(1 for s in special_students if cat in s["cats"])
        if cnt == 0:
            continue
        _, text, border, _ = CATEGORY_COLORS[cat]
        filter_btns += f"""<button onclick="filterTable('{cat}')"
            class="filter-btn px-3 py-1 rounded-full text-xs font-bold border {border} {text} hover:opacity-80" data-cat="{cat}">
            {cat} ({cnt})
        </button>"""

    return f"""
    <div class="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
        <div class="px-4 py-2.5 bg-slate-50 border-b border-gray-200 flex items-center justify-between flex-wrap gap-2">
            <h3 class="text-sm font-bold text-slate-700">
                특별전형 해당 학생 목록
                <span class="ml-2 inline-flex items-center justify-center px-2 py-0.5 rounded-full text-xs font-medium bg-slate-200 text-slate-600">{len(special_students)}</span>
            </h3>
            <div class="flex flex-wrap gap-1.5 no-print">
                {filter_btns}
            </div>
        </div>
        <div class="overflow-x-auto">
            <table class="w-full text-xs" id="studentTable">
                <thead class="bg-slate-100 border-b border-gray-200 text-slate-500">
                    <tr>
                        <th class="py-2 w-7">No</th>
                        <th class="py-2 w-10">반</th>
                        <th class="py-2 w-14">이름</th>
                        <th class="py-2 w-9">성별</th>
                        <th class="py-2 text-left pl-2">특별전형</th>
                        <th class="py-2 w-14">계열</th>
                        <th class="py-2 text-left pl-2">지원학교</th>
                        <th class="py-2 w-16">진행상황</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-gray-100" id="studentBody">
                    {rows_html}
                </tbody>
            </table>
        </div>
    </div>"""


def generate_html(students: list, year: str, filepath: str):
    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    special_cnt = sum(1 for s in students if s["cats"])
    summary_cards = make_summary_cards(students)
    cross_matrix  = make_cross_matrix(students)
    student_table = make_student_table(students)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>{year} 특별전형 현황</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet"
          href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css"/>
    <style>
        body {{ font-family: "Pretendard Variable", Pretendard, -apple-system, sans-serif;
               -webkit-print-color-adjust: exact; }}
        @media print {{
            @page {{ size: A4 landscape; margin: 10mm; }}
            .no-print {{ display: none !important; }}
            body {{ background: white; padding: 0; }}
            .shadow-sm {{ box-shadow: none; }}
        }}
        tr.hidden {{ display: none; }}
        .filter-btn.active {{ ring: 2px; opacity: 1; filter: brightness(0.9); }}
    </style>
</head>
<body class="p-8 bg-slate-50 min-h-screen">
    <div class="max-w-5xl mx-auto">
        <header class="flex justify-between items-end mb-6 border-b border-slate-300 pb-4">
            <div>
                <h1 class="text-2xl font-extrabold text-slate-800">{year} 특별전형 현황</h1>
                <p class="text-xs text-slate-500 mt-1">
                    사회통합전형 · 특례 · 보훈 · 쌍둥이 · 학폭 · 교직원자녀 · 장애 · 다자녀(3인+)
                </p>
            </div>
            <div class="text-right">
                <p class="text-xs text-slate-400 mb-2 font-mono">업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
                <button onclick="window.print()"
                        class="no-print bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-bold shadow">
                    🖨 인쇄하기
                </button>
            </div>
        </header>

        {summary_cards}
        {cross_matrix}
        {student_table}

        <footer class="mt-6 text-center border-t border-slate-200 pt-3">
            <p class="text-[10px] text-slate-400">
                {year}학년도 목일중학교 고입 특별전형 현황 &nbsp;|&nbsp;
                <a href="전기고_현황.html" class="underline">전기고</a> |
                <a href="후기고_현황.html" class="underline">후기고</a> |
                <a href="전체_현황.html" class="underline">전체</a>
            </p>
        </footer>
    </div>

    <script>
    function filterTable(cat) {{
        document.querySelectorAll('.filter-btn').forEach(btn => {{
            btn.classList.toggle('active', btn.dataset.cat === cat);
        }});
        document.querySelectorAll('#studentBody tr').forEach(row => {{
            if (!cat) {{
                row.classList.remove('hidden');
            }} else {{
                const cats = row.dataset.cats || '';
                row.classList.toggle('hidden', !cats.includes(cat));
            }}
        }});
    }}
    </script>
</body>
</html>"""

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)

    abs_path = os.path.abspath(filepath)
    print(f"  ✅ {os.path.basename(filepath)}  (특별전형 {special_cnt}명)")
    return abs_path


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"
    students = fetch_data(year)

    out = os.path.join(OUTPUT_DIR, "특별전형_현황.html")
    print("\n[HTML 생성]")
    path = generate_html(students, year, out)
    print(f"\n완료! {path}")

    try:
        import webbrowser
        webbrowser.open(f"file://{path}")
    except Exception:
        pass


if __name__ == "__main__":
    main()
