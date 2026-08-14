"""record_requirement 판정 검증."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "generators"))
from school_types import record_requirement as rr

# 교육감선발 일반고는 학생부 점검 대상이 아니다
assert rr("일반고", "신목고") == {"제출": "X", "반영학기": "", "출력유형": "", "점검필요": ""}

# 과고·영재고는 3-1까지
for t in ("영재고", "과학고"):
    r = rr(t, "서울과학고")
    assert (r["제출"], r["반영학기"], r["점검필요"]) == ("O", "3-1", "O"), r

# 외고·국제고, 예술계고는 3-2까지
assert rr("외고/국제고", "명덕외고")["반영학기"] == "3-2"
assert rr("예술계고", "선화예고")["반영학기"] == "3-2"

# 자사고는 학교로 전국/서울을 가른다
assert rr("자사고", "하나고등학교")["출력유형"] == "전국자사"
assert rr("자사고", "이화여고")["출력유형"] == "서울자사"
assert rr("자사고", "")["출력유형"] == "자사고(전국/서울 확인)"   # 학교 모르면 찍지 않는다

# 모르는 유형은 조용히 통과시키지 않는다
assert rr("", "")["제출"] == "확인필요"
assert rr("듣보고", "")["제출"] == "확인필요"

print("ok")
