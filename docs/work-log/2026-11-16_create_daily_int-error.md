# 作業記録: create_daily「invalid literal for int() with base 10: ' '」修正（2026-11-16）

## 現象
- 画面: https://shohshohshoh.github.io/monthly-Dashboard/ で 対象年月 2024/12 → 生成
- エラー: `create_daily の実行中にエラー: invalid literal for int() with base 10: ' '`

## 原因
- `template/create_daily.py` の商品別シート読み込み（旧 169〜172行）が `int(vals[n] or 0)` で変換していた。
- セルが半角スペース `' '` の場合、`' ' or 0` → `' '`（空でない文字列は真）となり、`int(' ')` で ValueError。
- 日次データ側は安全な関数 `_int()`（try/except で 0 返却）を使っていたため落ちず、商品別側だけ漏れていた。

## 修正
- 商品別の単価・数量・金額の変換を `_int()` に統一。
- `_int()` を強化: 空白（全角スペース含む）・カンマ・円記号を除去し、"12.0" のような小数文字列にも対応。
- パッチ: `2026-11-16_create_daily_blank-cell.patch`（ブランチ `fix/create-daily-blank-cell`）

## 根拠（仕様・参考）
- Python `int()`: 文字列は整数リテラルでなければ ValueError
  https://docs.python.org/3/library/functions.html#int
- Python ブール演算 `x or y`: x が真なら x を返す（空でない文字列は真）
  https://docs.python.org/3/reference/expressions.html#boolean-operations
  https://docs.python.org/3/library/stdtypes.html#truth-value-testing

## 学び
1. Excel由来の値は「None・空白文字・"-"・カンマ付き文字列」が混入する前提で、変換は必ず共通の安全関数を経由させる。
2. `int(x or 0)` は None/空文字には効くが、空白文字には効かない。
3. 新しく入力シートを読む処理を追加するときは `grep -n "int(" template/*.py` で直接変換が残っていないか確認する。
4. 元Excel側の対処（スペースだけのセルを Delete で空にする）でも回避できるが、コード側で吸収する方が再発しない。

## 追記（同日）: 「まだエラーが出る」
- 確認結果: GitHub の origin/main は 350f810 のまま。修正はプッシュされておらず、Render は旧コードで動作中。
- 修正版の create_daily.py には、直接の int() 変換は残っていない（コマンドライン引数の処理を除く）。
- このセッションからのプッシュは、Claude GitHub App の権限がないため 403 で拒否される。
- 学び: 「修正した」と「本番に反映された」は別物。ユーザーに報告するときは、反映の確認方法（GitHub のコミット履歴・Render の Deploy ログ）まで案内する。

## 追記（同日）: 反映完了
- Claude GitHub App インストール後、main へプッシュ（350f810 → d3f8770）。ops-deploy skill の手順「git push origin main」に従った。
- バックエンドは Render 側で Manual Deploy が必要な場合あり（ops-deploy skill 記載）。
- 学び: GitHub App 未インストールだとクローン（読み取り）はできてもプッシュは 403 になる。最初に push 権限の有無を確認する。

## 追記: Render ダッシュボードの説明
- URL: https://dashboard.render.com （Render の管理画面。test01-backend がバックエンド）
- 手動デプロイ: サービス画面右上「Manual Deploy」→「Deploy latest commit」
- 根拠: https://render.com/docs/deploys （Manual deploys / Automatic deploys）
- 学び: ユーザーは Render の存在を意識していない可能性がある。デプロイ案内では URL と画面上のボタン名まで示す。
- 画面構成メモ: Overview → Projects「My project」→ その中に test01-backend。サービスはプロジェクト配下にある。
- 確認結果: Render サービス名は実際には「monthly-Dashboard」（render.yaml の test01-backend ではない）。URL https://monthly-dashboard-n31f.onrender.com 。Auto-Deploy 有効で、d3f8770 が Live（手動デプロイ不要）。
- 学び: render.yaml の name と実サービス名が一致しないことがある。Auto-Deploy 有効なので main へのプッシュだけで反映される。Free プランは無操作時に停止し、初回リクエストが50秒以上遅れることがある。

## 結果: 解決（2026-11-16）
- 2024/12 の生成に成功。原因（商品別シートの空白セル）と修正（_int() への統一）が正しかったことを確認した。
