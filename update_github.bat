@echo off
chcp 65001 >nul
cd /d "%~dp0"
"C:\Program Files\Git\bin\git.exe" status
"C:\Program Files\Git\bin\git.exe" add .
"C:\Program Files\Git\bin\git.exe" commit -m "반 정보 채우기 로직 수정: 2행/16행/30행/44행에 A/B/C/D 반 정보 추가, A/I/J/K/P열 제외"
"C:\Program Files\Git\bin\git.exe" push origin main
pause
