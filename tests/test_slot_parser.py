import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "generators"))

from sync_type_to_tracking import detect_slots, data_status


def _row(v="", w="", x="", y="", z="", aa=""):
    """반별 시트 한 행: A:U(21칸) 더미 + V:AA 슬롯값."""
    return [""] * 21 + [v, w, x, y, z, aa]


def test_empty_row_has_no_slots():
    slots = detect_slots(_row())
    assert all(s["school"] == "" and s["result"] == "" for s in slots.values())


def test_basic_slot_parsing():
    slots = detect_slots(_row(v="서울과학영재학교", w="1차합", x="한성과학고", y="", z="하나고", aa="최종불"))
    assert slots["영재고"] == {"school": "서울과학영재학교", "result": "1차합", "flags": []}
    assert slots["전기"] == {"school": "한성과학고", "result": "", "flags": []}
    assert slots["후기"] == {"school": "하나고", "result": "최종불", "flags": []}


def test_nonstandard_result_code_flagged():
    slots = detect_slots(_row(x="한성과학고", y="합격했음"))
    assert "결과코드비표준" in slots["전기"]["flags"]
    assert slots["전기"]["result"] == "합격했음"  # 값은 보존


def test_unclassified_school_flagged_in_data_status():
    slots = detect_slots(_row(z="알수없는고"))
    status = data_status("일반고", slots, final_assigned="")
    assert status == "확인필요"


def test_short_row_padding():
    slots = detect_slots([""] * 23 + ["한성과학고"])  # X열까지만 존재
    assert slots["전기"]["school"] == "한성과학고"


def test_data_status_progression():
    empty = detect_slots(_row())
    assert data_status("", empty, "") == "미입력"
    assert data_status("과학고", empty, "") == "희망만"
    assert data_status("과학고", detect_slots(_row(v="경기과학영재학교")), "") == "영재고진행"
    assert data_status("과학고", detect_slots(_row(x="한성과학고", y="1차합")), "") == "전기진행"
    assert data_status("과학고", detect_slots(_row(z="대원외고")), "") == "후기진행"
    assert data_status("과학고", detect_slots(_row(x="한성과학고")), "서울고") == "배정완료"


def test_later_slot_wins_for_status():
    slots = detect_slots(_row(v="경기과학영재학교", w="최종불", x="한성과학고"))
    assert data_status("과학고", slots, "") == "전기진행"
