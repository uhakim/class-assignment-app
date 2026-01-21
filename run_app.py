# -*- coding: utf-8 -*-
"""
Streamlit 앱 실행 래퍼
exe로 패키징할 때 사용
"""
import sys
import os
import subprocess

# 현재 디렉토리를 Python 경로에 추가
if getattr(sys, 'frozen', False):
    # exe로 실행 중인 경우
    application_path = os.path.dirname(sys.executable)
else:
    # 스크립트로 실행 중인 경우
    application_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(application_path)

# Streamlit 앱 실행
if __name__ == "__main__":
    # streamlit run app.py --server.headless=true
    subprocess.run([
        sys.executable, "-m", "streamlit", "run", "app.py",
        "--server.headless=true",
        "--server.port=8501",
        "--browser.gatherUsageStats=false"
    ])
