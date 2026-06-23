#!/usr/bin/env python3
"""
ダッシュボード生成 FastAPI バックエンド

起動方法:
  cd template
  uvicorn server:app --port 8000 --reload

必要パッケージ:
  pip install fastapi uvicorn
"""
import subprocess
import sys
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

ROOT    = Path(__file__).parent.parent
DATA    = ROOT / "frontend" / "public" / "data"
SCRIPTS = Path(__file__).parent
PY      = sys.executable


class Req(BaseModel):
    year: int
    month: int


@app.post("/api/check")
def check(req: Req) -> dict:
    y, m = req.year, req.month
    return {
        "input_exists":     (DATA / f"daily_{y}_{m}.xlsx").exists(),
        "dashboard_exists": (DATA / f"dashboard_{y}_{m}.png").exists(),
        "report_exists":    (DATA / f"report_{y}_{m}.xlsx").exists(),
    }


@app.post("/api/generate")
def generate(req: Req) -> dict:
    y, m = req.year, req.month

    input_file = DATA / f"daily_{y}_{m}.xlsx"
    if not input_file.exists():
        raise HTTPException(
            404,
            detail=f"daily_{y}_{m}.xlsx が見つかりません。"
                   f"frontend/public/data/ フォルダに配置してください。",
        )

    for script in ["create_report.py", "create_dashboard_img.py"]:
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

    res = subprocess.run(
        [PY, str(SCRIPTS / "create_pptx.py"), str(y), str(m)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(SCRIPTS),
    )
    if res.returncode != 0:
        raise HTTPException(
            500,
            detail=f"create_pptx.py の実行中にエラーが発生しました:\n"
                   f"{res.stderr or res.stdout}",
        )

    return {
        "success": True,
        "message": f"dashboard_{y}_{m}.png / report_{y}_{m}.xlsx / dashboard_{y}_{m}.pptx を生成しました",
    }
