# 作業記録ライブラリ

不具合対応・改修ごとの記録と学びを蓄積する。新しい作業を始める前に、必ずここを確認する。

| 日付 | 件名 | 要点 |
|---|---|---|
| 2026-11-16 | [create_daily「invalid literal for int()」修正](2026-11-16_create_daily_int-error.md) | 商品別シートの空白セル。Excel由来の値は必ず `_int()` で変換する |

## 共通の学び
- Excel由来の値は「None・空白・"-"・カンマ付き文字列」が混ざる前提で扱う。
- 修正したことと本番に反映されたことは別。GitHub の main と Render の Live コミットで確認する。
- Render のサービス名は `monthly-Dashboard`（Auto-Deploy 有効。main へのプッシュで自動反映）。
