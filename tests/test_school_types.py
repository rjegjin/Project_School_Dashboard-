import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "generators"))

from school_types import classify_school


def test_keyword_classification():
    assert classify_school("한성과학고") == "과학고"
    assert classify_school("서울예술고등학교") == "예술계고"
    assert classify_school("수도전기공업고(마이스터)") == "특성화고"
    assert classify_school("대원외고") == "외고/국제고"
    assert classify_school("동탄국제고") == "외고/국제고"
    assert classify_school("경기북과학고") == "과학고"
    assert classify_school("서울과학영재학교") == "영재고"


def test_override_wins_over_keyword():
    assert classify_school("하나고") == "자사고"
    assert classify_school("민족사관고") == "자사고"


def test_unknown_returns_empty():
    assert classify_school("알수없는고등학교") == ""
    assert classify_school("") == ""
    assert classify_school("  ") == ""


def test_whitespace_and_suffix_normalization():
    assert classify_school(" 한성과학고등학교 ") == "과학고"
