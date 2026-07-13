---
name: api-reference
description: server.py (FastAPI) のAPIエンドポイント一覧・リクエスト/レスポンス形式を確認・変更するときに使う。エンドポイントの追加、レスポンス項目の変更、フロントエンドとの連携調整などで参照する。
---

# APIエンドポイント（template/server.py）

| エンドポイント | 内容 |
|---|---|
| `POST /api/drive-generate` | Google Drive から Excel 取得 → 生成（daily・PNG・PPTX）→ base64 返却＋`GOOGLE_DRIVE_UPLOAD_FOLDER_ID`へ自動アップロード |
| `POST /api/upload-and-generate` | アップロードされたExcelから生成（ローカル開発用）→ base64 返却＋Drive自動アップロード |
| `GET /api/list-reports` | Drive の PPTX フォルダ（`GOOGLE_DRIVE_OUTPUT_FOLDER_ID`）を読み取り、年月降順で一覧返却 |
| `GET /api/get-pptx-image/{file_id}` | Drive の PPTX から最初のスライド画像を PNG base64 で返却 |
| `GET /api/get-file/{file_id}` | Drive のファイルを base64 で返す |
| `GET /api/combine-daily` | Drive出力フォルダの daily_*.xlsx を全て読み込み、日次データシートを縦結合して1ファイルで返す |
| `GET /api/debug-drive` | Drive 接続・フォルダアクセス・環境変数の診断 |
| `GET /api/debug-fonts` | フォント診断（文字化け調査用） |

## リクエスト形式（drive-generate / upload-and-generate 共通の年月指定）

```json
{ "year": 2026, "month": 5 }
```

## レスポンス形式（drive-generate）

```json
{
  "success": true,
  "png_base64": "...",
  "pptx_base64": "...", "pptx_filename": "dashboard_2026_5.pptx",
  "daily_base64": "...", "daily_filename": "daily_2026_5.xlsx",
  "drive_upload": { "uploaded": true }
}
```

- 生成のたびに pptx・daily xlsx を `GOOGLE_DRIVE_UPLOAD_FOLDER_ID` フォルダへ自動アップロードする（同名ファイルがあれば上書き）。
- `drive_upload.uploaded` が `false` の場合は `reason` に失敗理由が入るが、アップロード失敗はレスポンス自体を失敗させない（base64でのダウンロード提供は継続する）。

## 実装上の注意

- 生成パイプライン（create_daily → create_dashboard_img → create_pptx）は `server.py` 起動時にモジュールとして import 済みで、リクエストごとに関数を直接呼び出す（`subprocess` は使わない。プロセス起動・再import のコストを避けるため）。
- ブロッキング処理（パイプライン実行・Driveアップロード）は `asyncio.to_thread` でスレッドプール実行し、他のリクエストをブロックしないようにする。
- Google Drive サービス（`_get_drive_service()`）はモジュールレベルでキャッシュされており、毎リクエストで再構築しない。
