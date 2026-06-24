# 営業日報 ダッシュボード生成 — フロントエンド

React + Vite で構築されたダッシュボード生成Webアプリのフロントエンドです。

## 概要

年月を入力すると、バックエンド（FastAPI）経由でPythonスクリプトが実行され、以下のファイルを自動生成します。

- `daily_{Y}_{M}.xlsx` — 日次データ
- `report_{Y}_{M}.xlsx` — 集計レポート
- `dashboard_{Y}_{M}.png` — ダッシュボード画像
- `dashboard_{Y}_{M}.pptx` — PowerPoint

## 起動方法

```bash
npm install
npm run dev
# → http://localhost:5173
```

バックエンド（port 8000）が別途起動している必要があります。  
ルートの `start.bat` を使うと両方まとめて起動できます。

## 主なファイル

| ファイル | 内容 |
|---|---|
| `src/App.jsx` | メインコンポーネント（入力・生成・履歴・ライトボックス） |
| `src/App.css` | スタイル定義 |
| `public/data/` | 生成されたExcel・PNG・PPTXの配信ディレクトリ |

## 主な機能

- 年月入力（例: `2026/5`）による一括生成
- 既存ファイルがある場合の上書き確認モーダル
- 生成済みダッシュボードのライトボックス全画面表示（Escキーで閉じる）
- PPTXダウンロード（File System Access API 対応）
- 生成履歴の一覧表示・削除
