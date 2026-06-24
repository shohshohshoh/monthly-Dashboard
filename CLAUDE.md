# test-01 プロジェクト

## 概要

営業日報（Excelファイル）を集計し、ダッシュボードPNG・レポートExcel・PowerPointを自動生成するWebアプリ。
FastAPI（バックエンド）＋ React/Vite（フロントエンド）構成。

## ディレクトリ構成

```
test-01/
├── data/                        # 営業日報Excelファイル（★営業日報YYYY年M月.xlsx）
├── template/
│   ├── server.py                # FastAPI バックエンド（port 8000）
│   ├── create_daily.py          # ① 営業日報 → daily_{Y}_{M}.xlsx
│   ├── create_report.py         # ② daily → report_{Y}_{M}.xlsx
│   ├── create_dashboard_img.py  # ③ daily → dashboard_{Y}_{M}.png
│   └── create_pptx.py           # ④ PNG → dashboard_{Y}_{M}.pptx
├── frontend/
│   ├── src/
│   │   ├── App.jsx              # メイン React コンポーネント
│   │   └── App.css              # スタイル
│   └── public/data/             # 生成ファイルの出力先
└── start.bat                    # バックエンド＋フロントエンド同時起動
```

## 生成パイプライン

```
data/★営業日報{Y}年{M}月.xlsx
  └─ create_daily.py ─→ daily_{Y}_{M}.xlsx
       └─ create_report.py ─→ report_{Y}_{M}.xlsx
       └─ create_dashboard_img.py ─→ dashboard_{Y}_{M}.png
            └─ create_pptx.py ─→ dashboard_{Y}_{M}.pptx
```

すべての出力は `frontend/public/data/` に保存される。

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

- バックエンド: http://localhost:8000
- フロントエンド: http://localhost:5173

## APIエンドポイント（server.py）

| エンドポイント | 内容 |
|---|---|
| `POST /api/check` | ソース・出力ファイルの存在確認 |
| `POST /api/generate` | パイプライン全体を実行して生成 |

## 技術スタック

- **バックエンド**: Python 3.x、FastAPI、uvicorn
- **データ処理**: openpyxl、numpy、matplotlib、python-pptx
- **フロントエンド**: React 18、Vite、CSS（カスタム）

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

3. **大きなExcelファイル・生成物はGit管理対象外**
   - `data/` 配下の営業日報Excelは除外。
   - スクリプト（.py）、React（.jsx/.css）のみバージョン管理。

4. **`.gitignore` 設定**
   ```
   /data/
   *.xlsx
   !/frontend/public/data/*.xlsx
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
