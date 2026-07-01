# test-01 プロジェクト

## 概要

営業日報（Excelファイル）を集計し、ダッシュボードPNG・日次Excel・PowerPointを自動生成するWebアプリ。
FastAPI（バックエンド）＋ React/Vite（フロントエンド）構成。

- **フロントエンド**: GitHub Pages（静的ホスティング）
- **バックエンド**: Render（クラウド Python サーバー）
- **Excel ソース**: Google Drive（サービスアカウント経由で自動取得）

## ディレクトリ構成

```
test-01/
├── data/                          # 営業日報Excelファイル（ローカル開発テスト用のみ）※git管理外
├── template/
│   ├── server.py                  # FastAPI バックエンド（port 8000）
│   ├── create_daily.py            # ① 営業日報 → daily_{Y}_{M}.xlsx
│   ├── create_dashboard_img.py    # ② daily → dashboard_{Y}_{M}.png
│   ├── create_pptx.py             # ③ PNG → dashboard_{Y}_{M}.pptx
│   ├── requirements.txt           # Python 依存パッケージ
│   └── fonts/
│       ├── BIZUDGothic-Bold.ttf   # 日本語フォント Bold（Render・ローカル共用）
│       └── BIZUDGothic-Regular.ttf
├── frontend/
│   ├── src/
│   │   ├── App.jsx                # メイン React コンポーネント
│   │   └── App.css                # スタイル
│   └── public/data/               # 生成ファイルの出力先（ローカル）
├── render.yaml                    # Render デプロイ設定
└── .github/workflows/deploy.yml   # GitHub Pages 自動デプロイ
```

## 生成パイプライン

```
data/★営業日報{Y}年{M}月.xlsx  ← Google Drive から自動取得
  └─ create_daily.py ─→ daily_{Y}_{M}.xlsx
       └─ create_dashboard_img.py ─→ dashboard_{Y}_{M}.png
            └─ create_pptx.py ─→ dashboard_{Y}_{M}.pptx
```

出力（PNG・PPTX・日次Excel）は base64 エンコードして API レスポンスで返し、ブラウザでダウンロードする。

## daily_{Y}_{M}.xlsx の列構成

### 日次データシート（テーブル形式）

| 列 | 内容 |
|---|---|
| 日付 | 日付（yyyy/mm/dd） |
| 曜日 | 月〜日 |
| 定休 | 休日は "休" |
| 祝日 | 祝日フラグ |
| 純売上高 | 税抜売上 |
| 消費税 | 消費税額 |
| 総売上高 | 税込売上（純売上高＋消費税） |
| 現金〜夕食売上 | 支払方法・カテゴリ別・客数・食事別売上 |

### 商品別シート（テーブル形式）

FOOD / DRINK の商品名・単価・数量・金額（日別・順位別）

## 運用方法（常にクラウドモード）

- フロントエンド: `https://shohshohshoh.github.io/monthly-Dashboard/`
- バックエンド: Render（`VITE_API_URL` GitHub リポジトリ変数で指定）
- Excel ソース: Google Drive フォルダに `★営業日報{Y}年{M}月.xlsx` を置くだけ
- 既存レポート一覧: `GOOGLE_DRIVE_OUTPUT_FOLDER_ID` フォルダの PPTX を自動読み取り・一覧表示

ダッシュボードの生成はクラウドのみ。ローカルでは生成しない。

## 開発時のローカル起動（コード確認用）

```bat
cd template
python -m uvicorn server:app --port 8000 --reload

cd frontend
npm run dev
```

## APIエンドポイント（server.py）

| エンドポイント | 内容 |
|---|---|
| `POST /api/drive-generate` | Google Drive から Excel 取得 → 生成（daily・PNG・PPTX）→ base64 返却 |
| `GET /api/list-reports` | Drive の PPTX フォルダを読み取り、年月降順で一覧返却 |
| `GET /api/get-pptx-image/{file_id}` | Drive の PPTX から最初のスライド画像を PNG base64 で返却 |
| `GET /api/get-file/{file_id}` | Drive のファイルを base64 で返す |
| `GET /api/debug-drive` | Drive 接続・フォルダアクセス・環境変数の診断 |
| `GET /api/debug-fonts` | フォント診断（文字化け調査用） |

リクエスト形式（drive-generate）:
```json
{ "year": 2026, "month": 5 }
```

レスポンス形式（drive-generate）:
```json
{
  "success": true,
  "png_base64": "...",
  "pptx_base64": "...", "pptx_filename": "dashboard_2026_5.pptx",
  "daily_base64": "...", "daily_filename": "daily_2026_5.xlsx"
}
```

## 環境変数

### Render（バックエンド）

| 変数名 | 内容 |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | サービスアカウント JSON キーの内容（文字列）|
| `GOOGLE_DRIVE_FOLDER_ID` | 営業日報 Excel を置く Google Drive フォルダ ID（読み取り元）|
| `GOOGLE_DRIVE_OUTPUT_FOLDER_ID` | 既存 PPTX 一覧を読み取る Google Drive フォルダ ID |

> **注意**: Google Drive への書き込みは行わない（サービスアカウントのスコープは `drive.readonly`）。  
> `GOOGLE_DRIVE_OUTPUT_FOLDER_ID` には、手動で配置した `dashboard_YYYY_M.pptx` を読み取り専用で参照する。

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
- **日本語フォント**: BIZ UDGothic Bold（`template/fonts/` に同梱、Render 環境で確実に利用可能）
- **フロントエンド**: React 18、Vite、CSS（カスタム）
- **クラウド**: Render（バックエンド）、GitHub Pages（フロントエンド）
- **ストレージ**: Google Drive API（サービスアカウント認証、読み取り専用）

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
   - 例: `create_daily: 翌月混入日付を除外`, `UI: レポート一覧を3列表示に変更`

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
