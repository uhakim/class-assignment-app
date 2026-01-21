# -*- coding: utf-8 -*-
"""
EXE 빌드를 위한 설정 파일
"""
import PyInstaller.__main__
import os

# PyInstaller 옵션
PyInstaller.__main__.run([
    'app.py',
    '--name=반편성배정프로그램',
    '--onefile',
    '--windowed',
    '--add-data=출력서식.xlsx;.',
    '--hidden-import=streamlit',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=class_assignment_logic',
    '--hidden-import=create_final_assignment',
    '--hidden-import=fill_output_format',
    '--collect-all=streamlit',
    '--collect-all=pandas',
    '--collect-all=openpyxl',
    '--noconsole',  # 콘솔 창 숨기기
])
