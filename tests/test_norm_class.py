"""반 표기 정규화 검증."""
import pandas as pd


def _norm_class(df):
    if "반" in df.columns:
        df["반"] = df["반"].astype(str).str.strip().str.lstrip("0").replace("", pd.NA).fillna(df["반"])
    return df


def test_norm_class():
    df = _norm_class(pd.DataFrame({"반": ["01", "1", " 03 ", "10", "", "0"]}))
    assert list(df["반"]) == ["1", "1", "3", "10", "", "0"], list(df["반"])
    assert "반" not in _norm_class(pd.DataFrame({"번호": [1]})).columns


if __name__ == "__main__":
    test_norm_class()
    print("ok")
