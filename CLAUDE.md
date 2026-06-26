# test-01 プロジェクト

## 概要

営業日報（Excelファイル）を集計し、ダッシュボードPNG・レポートExcel・PowerPointを自動生成するWebアプリ。
FastAPI（バックエンド）＋ React/Vite（フロントエンド）構成。

- **フロントエンド**: GitHub Pages（静的ホスティング）
- **バックエンド**: Render（クラウド Python サーバー）
- **Excel ソース**: Google Drive（サービスアカウント経由で自動取得）

## ディレクトリ構成

```
test-01/
├── data/                          # 営業日報Excelファイル（★営業日報YYYY年M月.xlsx）※git管理外
├── template/
│   ├── server.py                  # FastAPI バックエンド（port 8000）
│   ├── create_daily.py            # ① 営業日報 → daily_{Y}_{M}.xlsx
│   ├── create_report.py           # ② daily → report_{Y}_{M}.xlsx
│   ├── create_dashboard_img.py    # ③ daily → dashboard_{Y}_{M}.png
│   ├── create_pptx.py             # ④ PNG → dashboard_{Y}_{M}.pptx
│   ├── requirements.txt           # Python 依存パッケージ
│   ├── download_fonts.py          # フォントダウンロードスクリプト（現在は不使用）
│   └── fonts/
│       ├── BIZUDGothic-Bold.ttf   # 日本語フォント Bold（Render・ローカル共用）
│       └── BIZUDGothic-Regular.ttf
├── frontend/
│   ├── src/
│   │   ├── App.jsx                # メイン React コンポーネント
│   │   └── App.css                # スタイル
│   └── public/data/               # 生成ファイルの出力先（ローカル）
├── render.yaml                    # Render デプロイ設定
├── start.bat                      # バックエンド＋フロントエンド同時起動（ローカル用）
└── .github/workflows/deploy.yml   # GitHub Pages 自動デプロイ
```

## 生成パイプライン

```
data/★営業日報{Y}年{M}月.xlsx  ← Google Drive から自動取得（クラウド時）
  └─ create_daily.py ─→ daily_{Y}_{M}.xlsx
       └─ create_report.py ─→ report_{Y}_{M}.xlsx
       └─ create_dashboard_img.py ─→ dashboard_{Y}_{M}.png
            └─ create_pptx.py ─→ dashboard_{Y}_{M}.pptx
```

すべての出力は `frontend/public/data/` に保存される（ローカル）。  
クラウド時は base64 エンコードして API レスポンスで返す。

## 動作モード

### クラウドモード（GitHub Pages + Render）

- フロントエンド: `https://shohshohshoh.github.io/monthly-Dashboard/`
- バックエンド: Render の URL（`VITE_API_URL` GitHub リポジトリ変数で設定）
- Excel: Google Drive フォルダから自動取得
- 出力: base64 でブラウザに返し、その場でダウンロード

### ローカルモード

- フロントエンド: `http://localhost:5173`
- バックエンド: `http://localhost:8000`
- Excel: `data/` フォルダに配置
- 出力: `frontend/public/data/` に保存

## 起動方法

```bat
# 両サーバーを同時起動（推奨）
start.bat

# 個別起動
cd template
python -m uvicorn server:app --port 8000 --reload

cd frontend
npm run dev
```

## APIエンドポイント（server.py）

| エンドポイント | 内容 |
|---|---|
| `POST /api/check` | ソース・出力ファイルの存在確認（ローカル用） |
| `POST /api/generate` | パイプライン全体を実行して生成（ローカル用） |
| `POST /api/drive-generate` | Google Drive から Excel 取得 → 生成 → base64 返却（クラウド用） |
| `POST /api/upload-and-generate` | ブラウザから Excel アップロード → 生成 → base64 返却 |
| `GET /api/debug-fonts` | フォント診断（文字化け調査用） |

リクエスト形式（drive-generate）:
```json
{ "year": 2026, "month": 5 }
```

レスポンス形式（クラウド生成時）:
```json
{
  "success": true,
  "png_base64": "...",
  "pptx_base64": "...", "pptx_filename": "dashboard_2026_5.pptx",
  "daily_base64": "...", "daily_filename": "daily_2026_5.xlsx",
  "report_base64": "...", "report_filename": "report_2026_5.xlsx"
}
```

## 環境変数

### Render（バックエンド）

| 変数名 | 内容 |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | サービスアカウント JSON キーの内容（文字列）|
| `GOOGLE_DRIVE_FOLDER_ID` | Excel ファイルを置く Google Drive フォルダ ID |

### GitHub リポジトリ変数（Settings → Variables → Actions）

| 変数名 | 内容 |
|---|---|
| `VITE_API_URL` | Render のバックエンド URL（例: `https://test01-backend.onrender.com`）|

`VITE_API_URL` が未設定の場合、フロントエンドはローカルモード（`http://localhost:8000`）になる。

## セキュリティ注意事項

- **サービスアカウント JSON キーは絶対にコミットしない**（`.gitignore` に `*.json` を設定済み）
- Render の環境変数に文字列として設定する
- `.env` ファイルもコミット禁止

## 技術スタック

- **バックエンド**: Python 3.11、FastAPI、uvicorn
- **データ処理**: openpyxl、numpy、matplotlib、python-pptx
- **日本語フォント**: BIZ UDGothic Bold（`template/fonts/` に同梱）
- **フロントエンド**: React 18、Vite、CSS（カスタム）
- **クラウド**: Render（バックエンド）、GitHub Pages（フロントエンド）
- **ストレージ**: Google Drive API（サービスアカウント認証）

## スクリプト個別実行

```bash
cd template
python create_daily.py 2026 5
python create_report.py 2026 5
python create_dashboard_img.py 2026 5
python create_pptx.py 2026 5
```

---

## Git 運用ルール

### 基本方針

**コードを変更するたびに、必ずGitHubへプッシュする。**

### 手順

1. **変更後は即コミット＆プッシュ**
   ```bash
   git add <変更ファイル>
   git commit -m "変更内容の要約"
   git push origin main
   ```

2. **コミットメッセージ規則**
   - 日本語OK。変更の「なぜ」を一言で書く。
   - 例: `create_daily: 翌月混入日付を除外`, `①グラフY軸上限を最大値+15万に変更`

3. **Git 管理対象**
   - 対象: スクリプト（.py）、React（.jsx/.css）、設定ファイル、フォント（.ttf）
   - 除外: `data/` 配下の営業日報Excel、生成物（PNG・PPTX）、サービスアカウントJSON

4. **`.gitignore` 主要設定**
   ```
   /data/
   *.xlsx
   !/frontend/public/data/*.xlsx
   *.json
   !/frontend/public/data/*.json
   __pycache__/
   *.pyc
   .env
   ```

5. **ブランチ戦略**
   - 通常作業は `main` ブランチで直接運用。
   - 大きな機能追加・実験的変更は `feature/XXX` ブランチを切る。

### 注意事項

- プッシュ前に `git status` で変更内容を確認する。
- `.env` や認証情報を含むファイルは絶対にコミットしない。
- 出力ファイル（PNG・PPTX）はコミット対象外。
- GitHub へプッシュすると GitHub Actions が自動でフロントエンドをビルド・デプロイする。
- バックエンドは Render で **Manual Deploy** が必要（自動デプロイが有効な場合は不要）。
