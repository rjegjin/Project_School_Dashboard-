"""normalize_school 표준화 규칙 검증."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "generators"))
from school_types import normalize_school as n

# 같은 학교는 반드시 한 이름으로 모인다
assert n("신목고") == n("신목고등학교") == n("신목") == "신목고"
assert n("양정고") == n("양정고등학교") == "양정고"
assert n("백암고") == n("백암고등학교") == "백암고"
assert n("양천고") == n("양천고등학교") == "양천고"
assert n("명덕외고") == n("명덕외국어고등학교") == "명덕외고"
assert n("진명여고") == n("진명여자고등학교") == "진명여고"
assert n("세종과고") == n("세종과학고") == "세종과학고"
assert n("대구과학고") == n("대구과학고등학교") == n("대구과학고등학고") == "대구과학고"
assert n("영상고") == n("영상고등하교") == "영상고"

# 영재학교는 '고'가 아니다 — 꼬리표를 붙이면 안 된다
assert n("세종과학예술영재학교") == "세종과학예술영재학교"
assert n("인천과학예술영재학교") == "인천과학예술영재학교"

# 플레이스홀더 / 복수기재는 학교명이 아니다
for junk in ("", "  ", "ㅇ", "O", "미정", "O(하나고 고민)", "신목고, 양정고",
             "진명여고 목동고 금옥고", "광역자사/이화여고", "하나고 상상고"):
    assert n(junk) == "", f"{junk!r} → {n(junk)!r}"

print("ok")

# 내부 공백 표기도 같은 학교로 모인다
assert n("한성 과학 고") == n("한성과고") == "한성과학고"
assert n("양정 고") == "양정고"
assert n("미진학") == ""
print("ok2")
