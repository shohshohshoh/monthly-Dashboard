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
                detail=f"{script} の実行中にエラーが発生しました:\n"
                       f"{res.stderr or res.stdout}",
            )

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

    # アップロードされた Excel を data/ に保存
    src_dir = ROOT / "data"
    src_dir.mkdir(exist_ok=True)
    DATA.mkdir(parents=True, exist_ok=True)
    src_path = src_dir / f"★営業日報{y}年{m}月.xlsx"
    src_path.write_bytes(await file.read())

    # パイプライン実行
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

    png_path  = DATA / f"dashboard_{y}_{m}.png"
    pptx_path = DATA / f"dashboard_{y}_{m}.pptx"

    return {
        "success":      True,
        "png_base64":   base64.b64encode(png_path.read_bytes()).decode(),
        "pptx_base64":  base64.b64encode(pptx_path.read_bytes()).decode(),
        "filename":     f"dashboard_{y}_{m}.pptx",
    }
