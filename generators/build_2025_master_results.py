#!/usr/bin/env python3
"""
build_2025_master_results.py
2025학년도 최종 입시 결과 → 마스터 시트 생성

방식: 결과 우선(results-first)
Stage 1: 전기고/후기고/일반고 결과 시트 → {성명: (유형, 학교, 합격여부)}
Stage 2: 학생 기본정보(반별 조사) + Stage 1 결과 매칭 → 마스터 시트

마스터 시트 구조:
  반 | 번호 | 성명 | 성별 | 최종유형 | 배정학교 | 합격여부 | 데이터출처
"""

import sys
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import os
import re
from collections import defaultdict

def get_sheets_client():
    """Google Sheets 클라이언트"""
    KEY_FILE = "/home/rjegj/projects/.secrets/service_key.json"
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

def parse_early_high_final(rows):
    """
    전기고 최종 합불 시트 파싱
    3섹션 (과학고, 예고, 특성화고) 병렬 배치

    반환: {성명: (유형, 학교)}
    """
    results = {}

    # 헤더 행 찾기 (행2)
    if len(rows) < 3:
        return results

    # 각 섹션별 파싱
    for row_idx in range(3, len(rows)):
        row = rows[row_idx]
        # 패딩
        row = row + [''] * (16 - len(row))

        # 과학고: cols 0-4
        if len(row) > 4 and row[0].strip() and row[0].strip().isdigit():
            name = row[2].strip() if row[2] else ""
            school = row[4].strip() if row[4] else ""
            if name and school:
                results[name] = ('과학고', school)

        # 예고: cols 5-9
        if len(row) > 9 and row[5].strip() and row[5].strip().isdigit():
            name = row[7].strip() if row[7] else ""
            school = row[9].strip() if row[9] else ""
            if name and school:
                results[name] = ('예고', school)

        # 특성화고: cols 10-15
        if len(row) > 15 and row[10].strip() and row[10].strip().isdigit():
            name = row[12].strip() if row[12] else ""
            school = row[14].strip() if row[14] else ""
            if name and school:
                results[name] = ('특성화고', school)

    return results

def parse_late_high_final(rows):
    """
    후기고 최종 합불 시트 파싱
    3섹션 (자사고, 외고/국제고, 비평준화고)

    반환: {성명: (유형, 학교, 합격여부)}
    """
    results = {}

    if len(rows) < 3:
        return results

    for row_idx in range(3, len(rows)):
        row = rows[row_idx]
        row = row + [''] * (18 - len(row))

        # 자사고: cols 0-5
        if len(row) > 5 and row[0].strip() and row[0].strip().isdigit():
            name = row[2].strip() if row[2] else ""
            school = row[4].strip() if row[4] else ""
            result = row[5].strip().lower() if row[5] else ""
            pass_yn = '합격' if result == '합' else '불합격' if result == '불' else '미결'
            if name and school:
                results[name] = ('자사고', school, pass_yn)

        # 외고/국제고: cols 6-11
        if len(row) > 11 and row[6].strip() and row[6].strip().isdigit():
            name = row[8].strip() if row[8] else ""
            school = row[10].strip() if row[10] else ""
            result = row[11].strip().lower() if row[11] else ""
            pass_yn = '합격' if result == '합' else '불합격' if result == '불' else '미결'
            if name and school:
                results[name] = ('외고/국제고', school, pass_yn)

        # 비평준화고/중점고: cols 12-17
        if len(row) > 17 and row[12].strip() and row[12].strip().isdigit():
            name = row[14].strip() if row[14] else ""
            school = row[16].strip() if row[16] else ""
            result = row[17].strip().lower() if row[17] else ""
            pass_yn = '합격' if result == '합' else '불합격' if result == '불' else '미결'
            if name and school:
                results[name] = ('비평준화고', school, pass_yn)

    return results

def parse_general_high_xlsx():
    """
    일반고 배정 xlsx 파일 파싱

    반환: {성명: (일반고, 학교, 합격여부)}
    """
    results = {}
    final_file = "/home/rjegj/다운로드/2026학년도 후기고 최종결과.xlsx"

    if not os.path.exists(final_file):
        print(f"⚠ xlsx 파일 없음: {final_file}")
        return results

    try:
        df = pd.read_excel(final_file, sheet_name=0)

        # 컬럼명 정규화
        df.columns = df.columns.str.replace('\r\n', ' ').str.strip()

        # 헤더 중복 제거 (첫 행이 헤더면 스킵)
        if '성명' not in df.columns and '반' in df.iloc[0].values:
            df = df.iloc[1:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)

        # 성명 컬럼 확인
        if '성명' not in df.columns:
            print("⚠ xlsx 파일에서 '성명' 컬럼을 찾을 수 없음")
            return results

        # 사정 결과 컬럼 찾기
        result_col = None
        for col in df.columns:
            if '사정' in col and '결과' in col:
                result_col = col
                break

        # 배정 학교 컬럼 찾기
        school_col = None
        for col in df.columns:
            if '배정' in col and '학교' in col:
                school_col = col
                break

        if not result_col or not school_col:
            print(f"⚠ xlsx 파일에서 필요 컬럼을 찾을 수 없음 (사정: {result_col}, 학교: {school_col})")
            return results

        # 데이터 추출
        for _, row in df.iterrows():
            name = str(row['성명']).strip()
            result_str = str(row[result_col]).strip()
            school = str(row[school_col]).strip()

            if name and name != 'nan':
                pass_yn = '합격' if result_str == '합격' else '불합격' if result_str == '불합격' else '미결'
                results[name] = ('일반고', school, pass_yn)

        print(f"✓ xlsx에서 {len(results)}명 추출")
        return results
    except Exception as e:
        print(f"✗ xlsx 파싱 오류: {e}")
        return results

def load_student_base_info(client):
    """
    반별 조사 시트에서 학생 기본정보 로드
    반환: DataFrame (반, 번호, 성명, 성별)
    """
    doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')

    all_students = []
    target_pattern = re.compile(r"진학희망 및 지원유형 조사\(3\d{2}\)_Sheet1")

    for sht in doc.worksheets():
        if target_pattern.search(sht.title):
            try:
                rows = sht.get_all_values()
                if len(rows) < 3:
                    continue

                for row in rows[2:]:
                    if len(row) >= 4 and row[2].strip():
                        all_students.append({
                            '반': row[0].strip(),
                            '번호': row[1].strip(),
                            '성명': row[2].strip(),
                            '성별': row[3].strip() if len(row) > 3 else '',
                        })
            except:
                pass

    return pd.DataFrame(all_students)

def build_master_sheet(client):
    """
    마스터 결과 시트 생성 및 Google Sheets에 업로드
    """
    print("\n" + "=" * 80)
    print("Step 1: 최종 결과 데이터 수집")
    print("=" * 80)

    doc = client.open_by_key('1I_Cy5TZEnG0GmoThLPJJR7ZrXxUgXzsDDzu2zOtmjQI')

    # 전기고 파싱
    print("  전기고 최종 합불 파싱...", end=" ")
    try:
        early_sht = doc.worksheet("전기고 최종 합불")
        early_rows = early_sht.get_all_values()
        early_results = parse_early_high_final(early_rows)
        print(f"✓ {len(early_results)}명")
    except Exception as e:
        print(f"✗ {e}")
        early_results = {}

    # 후기고 파싱
    print("  후기고 최종 합불 파싱...", end=" ")
    try:
        late_sht = doc.worksheet("후기고 최종 합불")
        late_rows = late_sht.get_all_values()
        late_results = parse_late_high_final(late_rows)
        print(f"✓ {len(late_results)}명")
    except Exception as e:
        print(f"✗ {e}")
        late_results = {}

    # 일반고 파싱
    print("  일반고 배정(xlsx) 파싱...", end=" ")
    general_results = parse_general_high_xlsx()

    # 병합: 후기고와 일반고 합치기
    all_results = {**early_results}
    for name, data in late_results.items():
        all_results[name] = data
    for name, data in general_results.items():
        if name not in all_results:
            all_results[name] = data

    print("\n" + "=" * 80)
    print("Step 2: 학생 기본정보와 결과 병합")
    print("=" * 80)

    # 학생 기본정보 로드
    print("  반별 조사 시트 로드...", end=" ")
    student_df = load_student_base_info(client)
    print(f"✓ {len(student_df)}명")

    # 결과 데이터 추가
    print("  결과 정보 병합...", end=" ")
    student_df['최종유형'] = ''
    student_df['배정학교'] = ''
    student_df['합격여부'] = ''
    student_df['데이터출처'] = ''

    for idx, row in student_df.iterrows():
        name = row['성명']
        if name in all_results:
            result = all_results[name]
            if len(result) == 3:  # (유형, 학교, 합격여부)
                student_df.at[idx, '최종유형'] = result[0]
                student_df.at[idx, '배정학교'] = result[1]
                student_df.at[idx, '합격여부'] = result[2]
                if name in early_results:
                    student_df.at[idx, '데이터출처'] = '전기고합불시트'
                elif name in late_results:
                    student_df.at[idx, '데이터출처'] = '후기고합불시트'
                elif name in general_results:
                    student_df.at[idx, '데이터출처'] = '일반고xlsx'
            else:  # (유형, 학교) - 전기고
                student_df.at[idx, '최종유형'] = result[0]
                student_df.at[idx, '배정학교'] = result[1]
                student_df.at[idx, '합격여부'] = '합격'
                student_df.at[idx, '데이터출처'] = '전기고합불시트'
        else:
            student_df.at[idx, '합격여부'] = '미결'
            student_df.at[idx, '데이터출처'] = '기본정보만'

    print(f"✓ 완료")

    # Google Sheets 업로드
    print("\n" + "=" * 80)
    print("Step 3: Google Sheets에 마스터 시트 생성")
    print("=" * 80)

    print("  시트 생성/업데이트...", end=" ")
    try:
        # 기존 시트 삭제 (있으면)
        try:
            old_sht = doc.worksheet("2025_입시결과_마스터")
            doc.del_worksheet(old_sht)
        except:
            pass

        # 새 시트 생성
        new_sht = doc.add_worksheet(title="2025_입시결과_마스터", rows=len(student_df) + 1, cols=8)

        # 헤더
        header = ['반', '번호', '성명', '성별', '최종유형', '배정학교', '합격여부', '데이터출처']
        new_sht.append_row(header)

        # 데이터 추가 (배치)
        data_rows = []
        for _, row in student_df.iterrows():
            data_rows.append([
                row['반'],
                row['번호'],
                row['성명'],
                row['성별'],
                row['최종유형'],
                row['배정학교'],
                row['합격여부'],
                row['데이터출처']
            ])

        # 배치 업로드 (gspread의 배치 API)
        if data_rows:
            new_sht.append_rows(data_rows)

        print("✓ 완료")
    except Exception as e:
        print(f"✗ {e}")
        return False

    print("\n" + "=" * 80)
    print("✓ 마스터 시트 생성 완료")
    print("=" * 80)
    print(f"총 {len(student_df)}명")
    print(f"  - 전기고 합격: {(student_df['데이터출처'] == '전기고합불시트').sum()}명")
    print(f"  - 후기고 합격: {(student_df['데이터출처'] == '후기고합불시트').sum()}명")
    print(f"  - 일반고 배정: {(student_df['데이터출처'] == '일반고xlsx').sum()}명")
    print(f"  - 기본정보만: {(student_df['데이터출처'] == '기본정보만').sum()}명")
    print(f"\n합격자 총: {(student_df['합격여부'] == '합격').sum()}명")
    print(f"불합격: {(student_df['합격여부'] == '불합격').sum()}명")
    print(f"미결정: {(student_df['합격여부'] == '미결').sum()}명")

    return True

if __name__ == "__main__":
    try:
        print("2025학년도 입시결과 마스터 시트 생성 시작...")
        client = get_sheets_client()
        build_master_sheet(client)
    except KeyboardInterrupt:
        print("\n중단됨")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ 오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
