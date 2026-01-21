@echo off
chcp 65001
echo ========================================
echo 반편성 배정 프로그램 EXE 빌드
echo ========================================
echo.

echo PyInstaller 설치 확인 중...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller 설치 중...
    pip install pyinstaller
    echo.
)

echo.
echo EXE 파일 생성 중...
echo 이 작업은 몇 분 정도 걸릴 수 있습니다.
echo.

pyinstaller --name=반편성배정프로그램 ^
    --onefile ^
    --windowed ^
    --noconsole ^
    --add-data="출력서식.xlsx;." ^
    --add-data="class_assignment_logic.py;." ^
    --add-data="create_final_assignment.py;." ^
    --add-data="fill_output_format.py;." ^
    --hidden-import=streamlit ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=class_assignment_logic ^
    --hidden-import=create_final_assignment ^
    --hidden-import=fill_output_format ^
    --collect-all=streamlit ^
    --collect-all=pandas ^
    --collect-all=openpyxl ^
    app.py

echo.
echo ========================================
echo 빌드 완료!
echo ========================================
echo.
echo 생성된 파일: dist\반편성배정프로그램.exe
echo.
echo 배포 시 함께 복사할 파일:
echo   - 출력서식.xlsx (exe와 같은 폴더에)
echo.
pause
