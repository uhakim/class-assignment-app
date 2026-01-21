@echo off
chcp 65001
echo ========================================
echo 반편성 배정 프로그램 EXE 빌드
echo ========================================
echo.

echo 1. PyInstaller 설치 확인...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller가 설치되어 있지 않습니다. 설치 중...
    pip install pyinstaller
)

echo.
echo 2. EXE 파일 생성 중...
echo 이 작업은 몇 분 정도 걸릴 수 있습니다.
echo.

python build_exe.py

echo.
echo ========================================
echo 빌드 완료!
echo ========================================
echo.
echo 생성된 파일 위치: dist\반편성배정프로그램.exe
echo.
pause
