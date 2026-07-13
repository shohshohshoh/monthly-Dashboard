---
name: ops-deploy
description: 環境変数の設定、Google Driveの権限設定、ローカル起動、コミット・プッシュ、Renderへのデプロイなど運用まわりの作業をするときに使う。コードを変更してコミットする前には必ずこのSkillの手順に従う。
---

# 環境変数

## Render（バックエンド）

| 変数名 | 内容 |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | サービスアカウント JSON キーの内容（文字列）|
| `GOOGLE_DRIVE_FOLDER_ID` | 営業日報 Excel を置く Google Drive フォルダ ID（読み取り元）|
| `GOOGLE_DRIVE_OUTPUT_FOLDER_ID` | 既存 PPTX 一覧を読み取る Google Drive フォルダ ID（手動配置分の参照用）|
| `GOOGLE_DRIVE_UPLOAD_FOLDER_ID` | 生成した pptx・daily xlsx を自動アップロードする Google Drive フォルダ ID |

> サービスアカウントのDriveスコープは `drive`（読み書き）。`GOOGLE_DRIVE_UPLOAD_FOLDER_ID` で指定するフォルダは、サービスアカウントのメールアドレス（`client_email`）に **編集者** 権限で共有しておくこと（閲覧者権限のみだとアップロードに失敗する）。

## GitHub リポジトリ変数（Settings → Variables → Actions）

| 変数名 | 内容 |
|---|---|
| `VITE_API_URL` | Render のバックエンド URL（例: `https://test01-backend.onrender.com`）|

`VITE_API_URL` が未設定の場合、フロントエンドはローカルモード（`http://localhost:8000`）になる。

# セキュリティ注意事項

- **サービスアカウント JSON キーは絶対にコミットしない**（`.gitignore` に `*.json` を設定済み）
- Render の環境変数に文字列として設定する
- `.env` ファイルもコミット禁止

# 開発時のローカル起動（コード確認用）

```bat
cd template
python -m uvicorn server:app --port 8000 --reload

cd frontend
npm run dev
```

ダッシュボードの生成自体はクラウド（Render）のみで行う。ローカルでは生成しない。

# Git 運用ルール

## 基本方針

**コードを変更するたびに、必ずGitHubへプッシュする。**

## 手順

1. 変更後は即コミット＆プッシュ
   ```bash
   git add <変更ファイル>
   git commit -m "変更内容の要約"
   git push origin main
   ```
2. コミットメッセージ規則: 日本語OK。変更の「なぜ」を一言で書く。
   例: `create_daily: 翌月混入日付を除外`, `UI: レポート一覧を3列表示に変更`
3. Git 管理対象: スクリプト（.py）、React（.jsx/.css）、設定ファイル、フォント（.ttf）
   除外: `data/` 配下の営業日報Excel、生成物（PNG・PPTX）、サービスアカウントJSON
4. `.gitignore` 主要設定
   ```
   /data/
   *.xlsx
   !/frontend/public/data/*.xlsx
   *.json
   !/frontend/public/data/*.json
   __pycache__/
   *.pyc
   .env
   PLAN.md
   ```
5. ブランチ戦略: 通常作業は `main` ブランチで直接運用。大きな機能追加・実験的変更は `feature/XXX` ブランチを切る。

## 注意事項

- プッシュ前に `git status` で変更内容を確認する。
- `.env` や認証情報を含むファイルは絶対にコミットしない。
- 出力ファイル（PNG・PPTX）はコミット対象外。
- コミット前に `reviewer` サブエージェントで差分をレビューする（大きい変更・複数ファイルにまたがる変更のときは特に）。
- GitHub へプッシュすると GitHub Actions が自動でフロントエンドをビルド・デプロイする。
- バックエンドは Render で **Manual Deploy** が必要（自動デプロイが有効な場合は不要）。
