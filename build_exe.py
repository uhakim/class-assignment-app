# -*- coding: utf-8 -*-
"""
PyInstaller를 사용하여 exe 파일 생성 스크립트
"""
import PyInstaller.__main__
import os
import sys

# 현재 디렉토리
current_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller 옵션
options = [
    'app.py',  # 메인 스크립트
    '--name=반편성배정프로그램',  # exe 파일 이름
    '--onefile',  # 단일 exe 파일로 생성
    '--windowed',  # 콘솔 창 숨기기 (GUI만 표시)
    '--add-data=출력서식.xlsx;.',  # 출력서식.xlsx 포함 (Windows는 ; 사용)
    '--hidden-import=streamlit',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=class_assignment_logic',
    '--hidden-import=create_final_assignment',
    '--hidden-import=fill_output_format',
    '--collect-all=streamlit',  # Streamlit 관련 모든 파일 수집
    '--collect-all=pandas',
    '--collect-all=openpyxl',
]

# 필요한 파일들 추가
required_files = [
    'class_assignment_logic.py',
    'create_final_assignment.py',
    'fill_output_format.py',
]

for file in required_files:
    if os.path.exists(file):
        options.append(f'--add-data={file};.')

PyInstaller.__main__.run(options)
