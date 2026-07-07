#!/usr/bin/env python3
"""
export_class_sheets_pdf.py
301~314 반별 시트를 A~P열 범위로 PDF 일괄 내보내기

출력: ~/다운로드/진학희망조사_PDF/XXX반_진학희망조사.pdf
실행: /home/rjegj/projects/unified_venv/bin/python generators/export_class_sheets_pdf.py
"""

import os
import sys
import time
import logging
import requests
import gspread
from pypdf import PdfWriter
from google.oauth2.service_account import Credentials
import google.auth.transport.requests

# pypdf의 "Object N M not defined" 경고 억제
logging.getLogger("pypdf").setLevel(logging.ERROR)

KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
SPREADSHEET_ID = "14VeC3Dxj0Ou5-ddWTwfzktuWfB0Eoz_2CcDwNZPVEH0"
OUTPUT_DIR = os.path.expanduser("~/다운로드/진학희망조사_PDF")
RANGE = "A:P"
SHEET_RANGE = (301, 314)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_creds():
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    auth_req = google.auth.transport.requests.Request()
    creds.refresh(auth_req)
    return creds


def export_pdf(spreadsheet_id, gid, title, token, output_dir):
    url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
        f"?format=pdf"
        f"&gid={gid}"
        f"&range={RANGE}"
        f"&portrait=false"   # 가로 방향 (A~P 16열이므로)
        f"&scale=4"          # 페이지에 맞춤 (가로+세로 동시)
        f"&size=A4"
        f"&top_margin=0.25"
        f"&bottom_margin=0.25"
        f"&left_margin=0.25"
        f"&right_margin=0.25"
        f"&gridlines=true"
        f"&printtitle=true"
        f"&sheetnames=true"
        f"&fzr=false"        # 고정 행 반복 없음
    )
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    if resp.status_code == 200:
        path = os.path.join(output_dir, f"{title}반_진학희망조사.pdf")
        with open(path, "wb") as f:
            f.write(resp.content)
        return path
    else:
        raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:200]}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"출력 폴더: {OUTPUT_DIR}\n")

    creds = get_creds()
    gc = gspread.authorize(creds)
    doc = gc.open_by_key(SPREADSHEET_ID)

    # 301~314 시트만 필터링 & 정렬
    target = [
        (sht.title, sht.id)
        for sht in doc.worksheets()
        if sht.title.isdigit() and SHEET_RANGE[0] <= int(sht.title) <= SHEET_RANGE[1]
    ]
    target.sort(key=lambda x: int(x[0]))

    if not target:
        print("❌ 301~314 시트를 찾을 수 없습니다.")
        sys.exit(1)

    print(f"내보낼 시트: {[t for t, _ in target]}\n")

    DELAY = 6        # 기본 요청 간격 (초)
    RETRY_DELAYS = [15, 30]  # 재시도 대기 (초) — 최대 2회

    success, fail = [], []
    for i, (title, gid) in enumerate(target):
        if i > 0:
            time.sleep(DELAY)
        for attempt, wait in enumerate([0] + RETRY_DELAYS):
            if attempt > 0:
                print(f"  ↻ {title}반 재시도 {attempt}/{len(RETRY_DELAYS)} ({wait}초 대기)...")
                time.sleep(wait)
            try:
                path = export_pdf(SPREADSHEET_ID, gid, title, creds.token, OUTPUT_DIR)
                mark = "✓" if attempt == 0 else "✓ (재시도)"
                print(f"  {mark} {title}반 → {os.path.basename(path)}")
                success.append(title)
                break
            except Exception as e:
                if attempt == len(RETRY_DELAYS):
                    print(f"  ✗ {title}반 최종 실패: {e}")
                    fail.append(title)

    print(f"\n완료: 성공 {len(success)}개 / 실패 {len(fail)}개")
    if fail:
        print(f"실패 반: {fail}")
    if success:
        print(f"저장 위치: {OUTPUT_DIR}")

    # 성공한 PDF를 반 순서대로 하나로 합치기
    if len(success) > 1:
        merged_path = os.path.join(OUTPUT_DIR, "진학희망조사_전체.pdf")
        writer = PdfWriter()
        for title in sorted(success, key=lambda x: int(x)):
            pdf_path = os.path.join(OUTPUT_DIR, f"{title}반_진학희망조사.pdf")
            writer.append(pdf_path)
        with open(merged_path, "wb") as f:
            writer.write(f)
        print(f"합본 저장: {merged_path}")


if __name__ == "__main__":
    main()
