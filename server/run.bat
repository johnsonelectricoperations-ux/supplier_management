@echo off
cd /d %~dp0

if not exist venv (
    echo [설치] 가상환경 생성 중...
    python -m venv venv
)

call venv\Scripts\activate
echo [설치] 패키지 확인 중...
pip install -q -r requirements.txt

if not exist pdfs mkdir pdfs
if not exist db mkdir db

echo.
echo ============================================
echo   검사성적서 관리 시스템 시작
echo   접속 주소: http://10.80.101.200:5002
echo ============================================
echo.

uvicorn main:app --host 0.0.0.0 --port 5002 --workers 1 --access-log --log-level info
pause
