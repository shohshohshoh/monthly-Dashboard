#!/usr/bin/env python3
"""
ダッシュボード生成 FastAPI バックエンド

起動方法:
  cd template
  uvicorn server:app --port 8000 --reload

必要パッケージ:
  pip install fastapi uvicorn
"""
import base64
import io
import json
import os
import subprocess
import sys
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "https://shohshohshoh.github.io",
    ],
    allow_methods=["*"],
    allow_headers=["*"],
)

ROOT      = Path(__file__).parent.parent
DATA      = ROOT / "frontend" / "public" / "data"
SRC_DIR   = ROOT / "data"
SCRIPTS   = Path(__file__).parent
PY        = sys.executable


class Req(BaseModel):
    year: int
    month: int


def _run_pipeline(y: int, m: int):
    for script in ["create_daily.py", "create_report.py",
                   "create_dashboard_img.py", "create_pptx.py"]:
        res = subprocess.run(
            [PY, str(SCRIPTS / script), str(y), str(m)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            cwd=str(SCRIPTS),
        )
        if res.returncode != 0:
            raise HTTPException(
                500,
                detail=f"{script} の実行中にエラー:\n{res.stderr or res.stdout}",
            )


def _base64_result(y: int, m: int) -> dict:
    png_path  = DATA / f"dashboard_{y}_{m}.png"
    pptx_path = DATA / f"dashboard_{y}_{m}.pptx"
    return {
        "success":     True,
        "png_base64":  base64.b64encode(png_path.read_bytes()).decode(),
        "pptx_base64": base64.b64encode(pptx_path.read_bytes()).decode(),
        "filename":    f"dashboard_{y}_{m}.pptx",
    }


@app.post("/api/check")
def check(req: Req) -> dict:
    y, m = req.year, req.month
    return {
        "source_exists":    (SRC_DIR / f"★営業日報{y}年{m}月.xlsx").exists(),
        "daily_exists":     (DATA / f"daily_{y}_{m}.xlsx").exists(),
        "report_exists":    (DATA / f"report_{y}_{m}.xlsx").exists(),
        "dashboard_exists": (DATA / f"dashboard_{y}_{m}.png").exists(),
        "pptx_exists":      (DATA / f"dashboard_{y}_{m}.pptx").exists(),
    }


@app.post("/api/generate")
def generate(req: Req) -> dict:
    y, m = req.year, req.month

    src = SRC_DIR / f"★営業日報{y}年{m}月.xlsx"
    if not src.exists():
        raise HTTPException(
            404,
            detail=f"★営業日報{y}年{m}月.xlsx が data/ フォルダに見つかりません。",
        )

    _run_pipeline(y, m)

    return {
        "success": True,
        "message": (f"daily_{y}_{m}.xlsx / report_{y}_{m}.xlsx / "
                    f"dashboard_{y}_{m}.png / dashboard_{y}_{m}.pptx を生成しました"),
    }


@app.post("/api/upload-and-generate")
async def upload_and_generate(
    file: UploadFile = File(...),
    year: int = Form(...),
    month: int = Form(...),
) -> dict:
    y, m = year, month

    SRC_DIR.mkdir(exist_ok=True)
    DATA.mkdir(parents=True, exist_ok=True)
    (SRC_DIR / f"★営業日報{y}年{m}月.xlsx").write_bytes(await file.read())

    _run_pipeline(y, m)
    return _base64_result(y, m)


@app.post("/api/drive-generate")
async def drive_generate(req: Req) -> dict:
    """Google Drive の指定フォルダから Excel を取得して生成する"""
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError:
        raise HTTPException(500, detail="Google Drive SDK が未インストールです")

    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    folder_id  = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")
    if not creds_json:
        raise HTTPException(500, detail="環境変数 GOOGLE_SERVICE_ACCOUNT_JSON が設定されていません")
    if not folder_id:
        raise HTTPException(500, detail="環境変数 GOOGLE_DRIVE_FOLDER_ID が設定されていません")

    credentials = service_account.Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    service = build("drive", "v3", credentials=credentials)

    y, m = req.year, req.month
    filename = f"★営業日報{y}年{m}月.xlsx"
    query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get("files", [])
    if not files:
        raise HTTPException(404, detail=f"{filename} が Google Drive フォルダに見つかりません")

    # ダウンロード
    SRC_DIR.mkdir(exist_ok=True)
    DATA.mkdir(parents=True, exist_ok=True)
    request = service.files().get_media(fileId=files[0]["id"])
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    (SRC_DIR / filename).write_bytes(buf.getvalue())

    _run_pipeline(y, m)
    return _base64_result(y, m)
