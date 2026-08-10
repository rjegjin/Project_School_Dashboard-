"""Option A 리포트 스캔: Analytics output/ 재귀 탐색이 실제 산출물을 잡는지."""
import glob
import os

ANALYTICS_OUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "..", "Project_HighSchool_apply_Analytics", "output",
)


def test_scan_finds_reports_and_charts():
    assert os.path.isdir(ANALYTICS_OUT), ANALYTICS_OUT
    reports = glob.glob(os.path.join(ANALYTICS_OUT, "**", "*.html"), recursive=True)
    charts = glob.glob(os.path.join(ANALYTICS_OUT, "**", "*.png"), recursive=True)
    assert reports, "HTML 리포트를 못 찾음"
    assert charts, "차트 PNG를 못 찾음 (advanced_plots/ 하위 재귀 실패)"
    # 하위 폴더까지 내려가는지 확인
    assert any(os.sep + "advanced_plots" + os.sep in p for p in charts)


if __name__ == "__main__":
    test_scan_finds_reports_and_charts()
    print("ok")
