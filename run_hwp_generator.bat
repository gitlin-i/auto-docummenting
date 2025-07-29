@echo off
chcp 65001 > nul
echo ========================================
echo    청년이룸 HWP 생성기
echo ========================================
echo.

REM Python이 설치되어 있는지 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo 오류: Python이 설치되어 있지 않습니다.
    echo Python을 설치한 후 다시 실행해주세요.
    pause
    exit /b 1
)

REM 필요한 패키지가 설치되어 있는지 확인
python -c "import win32com.client" > nul 2>&1
if errorlevel 1 (
    echo 필요한 패키지를 설치합니다...
    pip install pywin32
)

echo HWP 생성기를 실행합니다...
python hwp_generator.py

echo.
echo 실행이 완료되었습니다.
pause 