@echo off
echo バックエンドサーバーをインストール・起動します...
pip install fastapi uvicorn -q

echo.
echo [1/2] バックエンド起動中 (http://localhost:8000)
start "Backend - FastAPI" cmd /k "cd /d %~dp0template && uvicorn server:app --port 8000"

timeout /t 2 /nobreak > nul

echo [2/2] フロントエンド起動中 (http://localhost:5173)
start "Frontend - Vite" cmd /k "cd /d %~dp0frontend && npm run dev"

echo.
echo 起動完了。ブラウザで http://localhost:5173 を開いてください。
