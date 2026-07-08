"""학교명 → 고입 유형 분류기 (시기 슬롯 구조 공용).

키워드로 못 잡는 학교(자사고는 이름에 '자사'가 없다)는
SCHOOL_TYPE_OVERRIDES에 등록한다. 미분류("")는 sync가 '확인필요'로
플래그하므로, 플래그가 뜰 때마다 여기에 추가하면 된다.
"""

# ponytail: 수동 사전 = 미분류 플래그가 뜰 때 채우는 calibration knob
SCHOOL_TYPE_OVERRIDES: dict[str, str] = {
    "하나고": "자사고",
    "민족사관고": "자사고",
    "민사고": "자사고",
    "상산고": "자사고",
    "외대부고": "자사고",
    "북일고": "자사고",
    "인천하늘고": "자사고",
}

# 순서 중요: 먼저 매칭되는 키워드가 이긴다 (영재 > 과학, 국제중학교 같은 오탐 없음)
_KEYWORD_RULES = [
    ("영재", "영재고"),
    ("과학고", "과학고"),
    ("과고", "과학고"),
    ("예술", "예술계고"),
    ("예고", "예술계고"),
    ("체육", "예술계고"),
    ("체고", "예술계고"),
    ("마이스터", "특성화고"),
    ("외국어", "외고/국제고"),
    ("외고", "외고/국제고"),
    ("국제고", "외고/국제고"),
]


def classify_school(name: str) -> str:
    """학교명에서 유형 표준값을 판정한다. 미분류는 ""."""
    cleaned = str(name).strip()
    if not cleaned:
        return ""
    # 등학교 제거: "고등학교" -> "고"
    base = cleaned.replace("등학교", "").rstrip()
    for key, type_name in SCHOOL_TYPE_OVERRIDES.items():
        if key in base:
            return type_name
    for keyword, type_name in _KEYWORD_RULES:
        if keyword in base:
            return type_name
    return ""
