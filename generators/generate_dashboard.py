#!/usr/bin/env python3
"""
2026 고입 진학현황 HTML 대시보드 생성기

데이터 소스: 입시_트래킹 시트 (auto_sync.py 파이프라인으로 최신 유지)
출력:
  reports/전기고_현황.html   — 영재고, 과학고, 예술계고, 특성화고
  reports/후기고_현황.html   — 자사고, 외고/국제고, 비평준화고, 기타/대안
  reports/전체_현황.html     — 전체 (유형 필터 포함)

실행:
  python generate_dashboard.py         # 2026 기본
  python generate_dashboard.py 2026
  python generate_dashboard.py 2025
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

EARLY_TYPES = ["영재고", "과학고", "예술계고", "특성화고"]
LATE_TYPES  = ["자사고", "외고/국제고", "비평준화고", "기타/대안", "일반고"]

# 유형별 색상 (Tailwind 클래스)
TYPE_COLOR = {
    "영재고":     ("bg-purple-100", "text-purple-700"),
    "과학고":     ("bg-blue-100",   "text-blue-700"),
    "예술계고":   ("bg-pink-100",   "text-pink-700"),
    "특성화고":   ("bg-orange-100", "text-orange-700"),
    "자사고":     ("bg-indigo-100", "text-indigo-700"),
    "외고/국제고":("bg-teal-100",   "text-teal-700"),
    "비평준화고": ("bg-cyan-100",   "text-cyan-700"),
    "기타/대안":  ("bg-gray-100",   "text-gray-600"),
    "일반고":     ("bg-slate-100",  "text-slate-600"),
}

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "reports")


# ==========================================
# 데이터 수집
# ==========================================
def fetch_from_tracking(year: str):
    """입시_트래킹 시트에서 전체 학생 데이터 로드"""
    ss_id = SPREADSHEET_IDS.get(year)
    if not ss_id:
        print(f"❌ 지원하지 않는 연도: {year}  (사용 가능: {', '.join(SPREADSHEET_IDS)})")
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

    # 필수 컬럼 인덱스
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

    students = []
    for r in rows[1:]:
        if len(r) < 3 or not r[c_name].strip():
            continue

        def get(idx):
            return r[idx].strip() if idx is not None and idx < len(r) else ""

        type_val   = get(c_type)
        school_val = get(c_school)
        final_val  = get(c_final)

        if not type_val:
            continue  # 유형 없는 학생 제외

        # 최종 결과 정규화
        result = ""
        if final_val.lower() in ["합격", "pass", "o", "○", "yes", "v"]:
            result = "합격"
        elif final_val.lower() in ["불합격", "fail", "x", "no"]:
            result = "불합격"

        students.append({
            "class":  get(c_cls),
            "num":    get(c_num),
            "name":   get(c_name),
            "gender": get(c_gender),
            "type":   type_val,
            "school": school_val or type_val,
            "result": result,
        })

    print(f"  → {len(students)}명 로드 완료")
    return students


# ==========================================
# HTML 생성
# ==========================================
def _type_badge(type_val: str) -> str:
    bg, text = TYPE_COLOR.get(type_val, ("bg-gray-100", "text-gray-600"))
    return f'<span class="px-2 py-0.5 rounded text-xs font-semibold {bg} {text}">{type_val}</span>'


def _status_badge(result: str) -> str:
    if result == "합격":
        return '<span class="px-2 py-1 rounded bg-green-100 text-green-700 text-xs font-bold">🎉 합격</span>'
    elif result == "불합격":
        return '<span class="px-2 py-1 rounded bg-gray-200 text-gray-500 text-xs font-bold">불합격</span>'
    else:
        return '<span class="px-2 py-1 rounded bg-indigo-50 text-indigo-500 text-xs font-bold">지원중</span>'


def _card_border(result: str) -> str:
    if result == "합격":
        return "border-green-400 ring-2 ring-green-100"
    elif result == "불합격":
        return "border-gray-200 opacity-60"
    else:
        return "border-gray-200 hover:border-indigo-300 hover:shadow-lg"


def _stat_bar(students) -> str:
    """유형별 통계 바"""
    from collections import Counter
    counts = Counter(s["type"] for s in students)
    pass_count = sum(1 for s in students if s["result"] == "합격")

    type_pills = ""
    for t, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        bg, text = TYPE_COLOR.get(t, ("bg-gray-100", "text-gray-600"))
        type_pills += f"""
        <button onclick="filterType('{t}')"
                class="type-btn px-3 py-1 rounded-full text-xs font-semibold {bg} {text} hover:opacity-80 transition cursor-pointer border border-transparent"
                data-type="{t}">
            {t} {cnt}명
        </button>"""

    pass_html = f' | <span class="text-green-600 font-bold">🎉 {pass_count}명 합격</span>' if pass_count > 0 else ""

    return f"""
    <div class="mb-6 flex flex-wrap items-center gap-2">
        <button onclick="filterType('전체')" class="type-btn px-3 py-1 rounded-full text-xs font-semibold bg-slate-800 text-white hover:opacity-80 transition cursor-pointer" data-type="전체">
            전체 {len(students)}명
        </button>
        {type_pills}
        <span class="text-sm text-gray-400 ml-2">{pass_html}</span>
    </div>"""


def _filter_script() -> str:
    return """
    <script>
    function filterType(type) {
        document.querySelectorAll('.student-card').forEach(card => {
            if (type === '전체' || card.dataset.type === type) {
                card.style.display = '';
            } else {
                card.style.display = 'none';
            }
        });
        document.querySelectorAll('.type-btn').forEach(btn => {
            btn.classList.toggle('ring-2', btn.dataset.type === type);
            btn.classList.toggle('ring-offset-1', btn.dataset.type === type);
        });
    }
    </script>"""


def generate_html(students, title: str, filepath: str, year: str) -> None:
    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    cards_html = ""
    for s in students:
        gender_cls = "text-blue-600 bg-blue-50" if s["gender"] == "남" else "text-red-500 bg-red-50"
        card = f"""
        <div class="student-card bg-white rounded-xl p-5 border {_card_border(s['result'])} transition-all duration-300 shadow-sm flex flex-col justify-between"
             data-type="{s['type']}">
            <div>
                <div class="flex justify-between items-start mb-3">
                    <div>
                        <span class="text-xs font-bold text-gray-400">{s['class']}반 {s['num']}번</span>
                        <h3 class="text-lg font-extrabold text-gray-800 mt-0.5">{s['name']}</h3>
                    </div>
                    <span class="px-2 py-1 rounded text-xs font-bold {gender_cls}">{s['gender']}</span>
                </div>
                <div class="mb-4">
                    {_type_badge(s['type'])}
                    <div class="text-gray-900 font-bold text-base mt-1 leading-tight">{s['school']}</div>
                </div>
            </div>
            <div class="pt-3 border-t border-gray-100">
                {_status_badge(s['result'])}
            </div>
        </div>"""
        cards_html += card

    stat_bar = _stat_bar(students)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" as="style" crossorigin
          href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css"/>
    <style>
        body {{ font-family: "Pretendard Variable", Pretendard, -apple-system, BlinkMacSystemFont, system-ui, sans-serif; }}
    </style>
</head>
<body class="bg-slate-50 min-h-screen p-6 md:p-12">
    <div class="max-w-7xl mx-auto">
        <header class="mb-8 flex flex-col md:flex-row md:items-end justify-between gap-4">
            <div>
                <h1 class="text-3xl md:text-4xl font-black text-slate-800 mb-1">{title}</h1>
                <p class="text-slate-500 text-sm">{year}학년도 목일중학교 진학현황</p>
            </div>
            <div class="text-right text-xs text-gray-400">
                업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}
            </div>
        </header>

        {stat_bar}

        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
            {cards_html}
        </div>

        <footer class="mt-12 text-center text-gray-400 text-sm">
            {year}학년도 목일중학교 진학현황 대시보드 —
            <a href="전기고_현황.html" class="underline hover:text-indigo-500">전기고</a> |
            <a href="후기고_현황.html" class="underline hover:text-indigo-500">후기고</a> |
            <a href="전체_현황.html" class="underline hover:text-indigo-500">전체</a>
        </footer>
    </div>
    {_filter_script()}
</body>
</html>"""

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)

    abs_path = os.path.abspath(filepath)
    print(f"  ✅ {os.path.basename(filepath)}  ({len(students)}명)")
    return abs_path


# ==========================================
# 실행
# ==========================================
def main():
    year = sys.argv[1] if len(sys.argv) > 1 else "2026"

    students = fetch_from_tracking(year)
    if not students:
        print("⚠️ 데이터 없음. 종료.")
        return

    early = [s for s in students if s["type"] in EARLY_TYPES]
    late  = [s for s in students if s["type"] in LATE_TYPES]

    out = lambda name: os.path.join(OUTPUT_DIR, name)

    print(f"\n[HTML 생성]")
    paths = []
    if early:
        paths.append(generate_html(early,    f"{year} 전기고 지원 현황", out("전기고_현황.html"), year))
    if late:
        paths.append(generate_html(late,     f"{year} 후기고 지원 현황", out("후기고_현황.html"), year))
    if students:
        paths.append(generate_html(students, f"{year} 전체 진학 현황",   out("전체_현황.html"),  year))

    print(f"\n완료! reports/ 폴더에 HTML {len(paths)}개 생성됨.")
    print(f"  경로: {os.path.abspath(OUTPUT_DIR)}/")

    # 브라우저 자동 열기 (전체 현황 우선)
    try:
        import webbrowser
        target = paths[-1] if paths else None
        if target:
            webbrowser.open(f"file://{target}")
    except Exception:
        pass


if __name__ == "__main__":
    main()
