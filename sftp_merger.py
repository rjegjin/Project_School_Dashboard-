"""
SFTP 원격 파일을 다운로드해서 2025 데이터와 병합하는 모듈
"""

import os
import paramiko
import pandas as pd
from io import BytesIO
import streamlit as st

class SFTPDataMerger:
    def __init__(self, host, username, password=None, key_file=None):
        """
        SFTP 연결 초기화
        
        Args:
            host: SFTP 서버 주소 (예: 100.68.98.81)
            username: 사용자명 (예: user)
            password: 비밀번호 (옵션)
            key_file: SSH 키 파일 경로 (옵션)
        """
        self.host = host
        self.username = username
        self.password = password
        self.key_file = key_file
    
    def connect(self):
        """SFTP 연결"""
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            if self.key_file:
                ssh.connect(self.host, username=self.username, key_filename=self.key_file)
            elif self.password:
                ssh.connect(self.host, username=self.username, password=self.password)
            else:
                raise ValueError("인증 정보가 없습니다 (비밀번호 또는 SSH 키 필요)")
            
            sftp = ssh.open_sftp()
            return sftp, ssh
        except Exception as e:
            st.error(f"❌ SFTP 연결 실패: {e}")
            return None, None
    
    def download_file(self, remote_path):
        """
        SFTP에서 파일 다운로드
        
        Args:
            remote_path: 원격 파일 경로
            
        Returns:
            pandas DataFrame
        """
        sftp, ssh = self.connect()
        if not sftp:
            return None
        
        try:
            # 파일 다운로드 (메모리에)
            file_obj = BytesIO()
            sftp.getfo(remote_path, file_obj)
            file_obj.seek(0)
            
            # Excel 읽기
            df = pd.read_excel(file_obj)
            
            sftp.close()
            ssh.close()
            
            return df
        except Exception as e:
            st.error(f"❌ 파일 다운로드 실패: {e}")
            return None
    
    def merge_with_legacy(self, legacy_df, final_results_df, merge_key='성명'):
        """
        레거시 데이터와 최종 결과 병합
        
        Args:
            legacy_df: 2025 레거시 데이터
            final_results_df: 원격 최종 결과 파일
            merge_key: 병합 기준 컬럼 (성명)
            
        Returns:
            병합된 DataFrame
        """
        try:
            # 최종 결과 파일 정규화
            final_results_df = final_results_df.rename(columns=str.strip)
            
            # 병합 (left join)
            merged = legacy_df.merge(
                final_results_df,
                on=merge_key,
                how='left'
            )
            
            # 최종 컬럼 우선순위: 최종 (원격) > 최종 (로컬)
            if '최종_x' in merged.columns and '최종_y' in merged.columns:
                merged['최종'] = merged['최종_y'].fillna(merged['최종_x'])
                merged = merged.drop(['최종_x', '최종_y'], axis=1)
            
            return merged
        except Exception as e:
            st.error(f"❌ 병합 실패: {e}")
            return legacy_df

# ==========================================
# 사용 예시
# ==========================================
if __name__ == "__main__":
    # 설정
    SFTP_HOST = "100.68.98.81"
    SFTP_USER = "user"
    SFTP_PASSWORD = "YOUR_PASSWORD"  # 여기 입력
    REMOTE_FILE = "C:/Users/user/OneDrive/문서/학교근무/목일중/(담당업무) 진로/2026학년도 후기고 최종결과.xlsx"
    
    # 병합
    merger = SFTPDataMerger(SFTP_HOST, SFTP_USER, password=SFTP_PASSWORD)
    final_data = merger.download_file(REMOTE_FILE)
    
    if final_data is not None:
        print("✅ 원격 파일 다운로드 성공")
        print(final_data.head())
