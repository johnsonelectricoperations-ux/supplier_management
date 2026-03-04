"""
main.py - FastAPI 서버 (서버PC용)
접속: http://10.80.101.200:5002
"""
import os
import json
import base64
import mimetypes
from pathlib import Path
from typing import Optional
from datetime import datetime

from fastapi import FastAPI, Query, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import database as db
from models import SyncBatchRequest, PdfSyncRequest, AppConfigRequest

# ─── 경로 설정 ────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
PDF_DIR    = BASE_DIR / "pdfs"
STATIC_DIR = BASE_DIR / "static"
CONFIG_FILE = BASE_DIR / "config.json"
SA_FILE     = BASE_DIR / "service_account.json"


def _load_cfg() -> dict:
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    return {}


def _save_cfg(data: dict):
    CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

PDF_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

# ─── FastAPI 초기화 ───────────────────────────────────
app = FastAPI(title="검사성적서 관리 시스템", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],        # 사내망이므로 전체 허용
    allow_methods=["*"],
    allow_headers=["*"],
)

# DB 초기화
db.init_db()

# ─── 정적 파일 ────────────────────────────────────────
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=FileResponse)
def index():
    return FileResponse(str(STATIC_DIR / "index.html"))


# ─── 설정 API ─────────────────────────────────────────

@app.get("/api/config")
def get_config():
    """설정 조회 (Spreadsheet ID, Drive Folder ID, 서비스 계정 여부)"""
    cfg = _load_cfg()
    return {
        "success":       True,
        "sheet_id":      cfg.get("sheet_id", ""),
        "folder_id":     cfg.get("folder_id", ""),
        "sa_configured": SA_FILE.exists(),
    }


@app.post("/api/config")
def save_config(req: AppConfigRequest):
    """설정 저장"""
    cfg = _load_cfg()
    cfg["sheet_id"]  = req.sheet_id
    cfg["folder_id"] = req.folder_id
    _save_cfg(cfg)
    return {"success": True, "message": "설정이 서버에 저장되었습니다."}


@app.get("/api/service-account")
def get_service_account():
    """서비스 계정 키 정보 반환 (사내망 전용)"""
    if not SA_FILE.exists():
        raise HTTPException(
            status_code=404,
            detail="service_account.json 파일이 없습니다. 서버PC의 server/ 폴더에 파일을 배치해 주세요."
        )
    try:
        sa = json.loads(SA_FILE.read_text(encoding="utf-8"))
        return {
            "success":      True,
            "client_email": sa["client_email"],
            "private_key":  sa["private_key"],
            "token_uri":    sa.get("token_uri", "https://oauth2.googleapis.com/token"),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"서비스 계정 파일 읽기 실패: {e}")


# ─── 동기화 API ───────────────────────────────────────

@app.post("/api/sync/batch")
def sync_batch(req: SyncBatchRequest):
    """
    브라우저(인터넷PC)에서 Google Sheets 데이터를 읽어 일괄 전송
    - incoming_data: [업체명]_Data 시트 전체
    - inspection_results: [업체명]_Result 시트 전체
    - item_list: [업체명]_ItemList 시트 전체
    """
    try:
        data_rows    = [r.dict() for r in req.incoming_data]
        result_rows  = [r.dict() for r in req.inspection_results]
        item_rows    = [r.dict() for r in req.item_list]

        dc = db.upsert_incoming_data(data_rows)
        rc = db.upsert_inspection_results(result_rows)
        ic = db.upsert_item_list(item_rows)

        db.add_sync_log(dc, rc, ic, 0, "success",
                        f"데이터:{dc}건, 결과:{rc}건, 품목:{ic}건 동기화 완료")

        return {
            "success": True,
            "data_count": dc,
            "result_count": rc,
            "item_count": ic,
            "message": f"동기화 완료 (입고:{dc}, 검사결과:{rc}, 품목:{ic})"
        }
    except Exception as e:
        db.add_sync_log(0, 0, 0, 0, "error", str(e))
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/sync/pdf")
def sync_pdf(req: PdfSyncRequest):
    """
    브라우저에서 Google Drive PDF를 다운로드한 뒤 서버PC에 저장
    Drive 구조: 검사성적서/[업체명]/[년도]/[월]/파일명
    """
    try:
        # 저장 경로 생성
        save_dir = PDF_DIR / req.company_name / req.year / req.month
        save_dir.mkdir(parents=True, exist_ok=True)

        file_path = save_dir / req.file_name

        # base64 디코딩 후 저장
        file_bytes = base64.b64decode(req.file_data)
        with open(file_path, "wb") as f:
            f.write(file_bytes)

        # 상대 경로 (API 서빙용)
        rel_path = f"{req.company_name}/{req.year}/{req.month}/{req.file_name}"

        # incoming_data의 local_pdf_path 업데이트 (tm_no 기준)
        if req.tm_no:
            conn = db.get_conn()
            rows = conn.execute(
                "SELECT id FROM incoming_data WHERE tm_no=? AND company_name=? LIMIT 1",
                (req.tm_no, req.company_name)
            ).fetchall()
            conn.close()
            for row in rows:
                db.update_local_pdf_path(row["id"], rel_path)

        return {
            "success": True,
            "path": rel_path,
            "size": len(file_bytes),
            "message": f"{req.file_name} 저장 완료"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/sync/status")
def sync_status():
    """마지막 동기화 상태 조회"""
    last = db.get_last_sync()
    return {"success": True, "last_sync": last}


# ─── 이력 조회 API ────────────────────────────────────

@app.get("/api/history")
def get_history(
    company_name: Optional[str] = Query(None),
    date_from:    Optional[str] = Query(None),
    date_to:      Optional[str] = Query(None),
    tm_no:        Optional[str] = Query(None),
    inspection_type: Optional[str] = Query(None),
    page:         int = Query(1, ge=1),
    page_size:    int = Query(50, ge=1, le=200),
):
    """검사성적서 이력 목록 조회 (필터 + 페이징)"""
    result = db.search_history(
        company_name=company_name or "",
        date_from=date_from or "",
        date_to=date_to or "",
        tm_no=tm_no or "",
        inspection_type=inspection_type or "",
        page=page,
        page_size=page_size,
    )
    return {"success": True, **result}


@app.get("/api/history/results")
def get_results(
    date:         str = Query(...),
    company_name: str = Query(...),
    tm_no:        str = Query(...),
):
    """특정 입고건의 검사결과 상세 조회"""
    results = db.get_inspection_results_by_key(date, company_name, tm_no)
    return {"success": True, "data": results}


@app.get("/api/companies")
def get_companies():
    """업체 목록 조회"""
    return {"success": True, "data": db.get_companies()}


# ─── PDF 서빙 ─────────────────────────────────────────

@app.get("/api/pdf/{file_path:path}")
def serve_pdf(file_path: str):
    """
    로컬 PDF 파일 서빙
    경로: /api/pdf/[업체명]/[년도]/[월]/파일명
    """
    full_path = PDF_DIR / file_path

    # 경로 순회 공격 방지
    try:
        full_path = full_path.resolve()
        PDF_DIR.resolve()
        full_path.relative_to(PDF_DIR.resolve())
    except (ValueError, RuntimeError):
        raise HTTPException(status_code=403, detail="접근 거부")

    if not full_path.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")

    return FileResponse(
        str(full_path),
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{full_path.name}"'}
    )


@app.get("/api/pdf-list/{company_name}")
def list_pdfs(company_name: str):
    """업체별 저장된 PDF 목록 조회"""
    company_dir = PDF_DIR / company_name
    if not company_dir.exists():
        return {"success": True, "data": []}

    files = []
    for f in sorted(company_dir.rglob("*.pdf")):
        rel = f.relative_to(PDF_DIR)
        parts = list(rel.parts)
        files.append({
            "path": str(rel),
            "name": f.name,
            "year": parts[1] if len(parts) > 1 else "",
            "month": parts[2] if len(parts) > 2 else "",
            "size": f.stat().st_size,
        })
    return {"success": True, "data": files}


# ─── 실행 ─────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=5002, reload=False)
