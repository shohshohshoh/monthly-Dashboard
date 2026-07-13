# test-01 プロジェクト

営業日報（Excelファイル）を集計し、ダッシュボードPNG・日次Excel・PowerPointを自動生成するWebアプリ。

- **フロントエンド**: React/Vite → GitHub Pages（`https://shohshohshoh.github.io/monthly-Dashboard/`）
- **バックエンド**: FastAPI → Render（`https://test01-backend.onrender.com` 相当。`VITE_API_URL` で指定）
- **Excel ソース**: Google Drive（サービスアカウント経由で自動取得、生成物も自動アップロード）

ダッシュボードの生成は常にクラウド（Render）のみで行う。ローカルでは生成しない。

## ディレクトリ構成

```
test-01/
├── data/                          # 営業日報Excel（ローカル開発テスト用のみ、git管理外）
├── template/
│   ├── server.py                  # FastAPI バックエンド（port 8000）
│   ├── create_daily.py            # ① 営業日報 → daily_{Y}_{M}.xlsx
│   ├── create_dashboard_img.py    # ② daily → dashboard_{Y}_{M}.png
│   ├── create_pptx.py             # ③ PNG → dashboard_{Y}_{M}.pptx
│   └── fonts/                     # 日本語フォント（BIZ UDGothic）
├── frontend/
│   ├── src/App.jsx                # メイン React コンポーネント
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

`server.py` は3スクリプトをモジュールとして直接呼び出す（サブプロセス起動はしない）。出力（PNG・PPTX・日次Excel）は base64 でレスポンス返却し、同時に Google Drive の指定フォルダへ自動アップロードする。

## 技術スタック

Python 3.11 / FastAPI / openpyxl / numpy / matplotlib / python-pptx / React 18 / Vite / Google Drive API（サービスアカウント）

## 詳細情報（用途に応じてSkillを参照）

- APIエンドポイント・リクエスト/レスポンス形式を触るとき → `api-reference` skill
- `daily_{Y}_{M}.xlsx` の列構成や集計ロジックを触るとき → `daily-xlsx-format` skill
- 環境変数・Google Drive権限・ローカル起動・コミット/デプロイ手順 → `ops-deploy` skill

## 作業の進め方（必ず守ること）

- **設計判断は人間が決める**: アーキテクチャ変更・認証方式の変更・依存ライブラリの大きな入れ替えなど、影響範囲が広い判断は着手前にユーザーに確認する。実装の詳細（関数分割・変数名など）は任されている。
- **計画はファイルに書き出す**: 複数ステップにまたがる作業に着手する前に、作業計画を `PLAN.md`（コミット対象外・`.gitignore`済み）に書き出してから進める。`/compact` は作業の区切り（1つのタスク完了時）でのみ使い、実装途中では使わない。
- **1指示に詰め込まない**: 範囲の広い依頼は独立したタスクに分割して進める。レビューが必要な変更は `reviewer` サブエージェントに委任してからコミットする。
- **コードを変更するたびに、必ずGitHubへプッシュする**（手順は `ops-deploy` skill）。
- **サービスアカウント JSON キー・`.env` は絶対にコミットしない**（`.gitignore` 設定済み）。
