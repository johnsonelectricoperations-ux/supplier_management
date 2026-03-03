#!/bin/bash
# ============================================================
#  검사성적서 관리 시스템 - 서버 실행 스크립트
#  서버PC: 10.80.101.200:5002
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Python 가상환경 확인/생성
if [ ! -d "venv" ]; then
  echo "[설치] 가상환경 생성 중..."
  python3 -m venv venv
fi

# 가상환경 활성화
source venv/bin/activate

# 패키지 설치 (최초 1회)
echo "[설치] 패키지 확인 중..."
pip install -q -r requirements.txt

# 필요 디렉토리 생성
mkdir -p pdfs db

echo ""
echo "============================================"
echo "  검사성적서 관리 시스템 시작"
echo "  접속 주소: http://10.80.101.200:5002"
echo "============================================"
echo ""

# 서버 실행 (0.0.0.0 = 사내망 전체 접근 허용)
uvicorn main:app \
  --host 0.0.0.0 \
  --port 5002 \
  --workers 1 \
  --access-log \
  --log-level info
