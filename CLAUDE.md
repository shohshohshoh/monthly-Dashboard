# test-01 プロジェクト

## 概要

営業日報（Excelファイル）を集計し、ダッシュボード形式のPowerPointを自動生成するPythonプロジェクト。

## ディレクトリ構成

```
test-01/
├── data/               # 営業日報Excelファイル（★営業日報YYYY年M月.xlsx）
├── template/
│   ├── dashboard_template.py      # ダッシュボード生成スクリプト（メインテンプレート）
│   └── ダッシュボードテンプレート.pptx  # PowerPointテンプレート
└── list_all.xlsx       # 集計済み一覧データ
```

## 主要ファイル

- [template/dashboard_template.py](template/dashboard_template.py) — ダッシュボード生成スクリプト。`[CONFIG]` タグのついた箇所がカスタマイズポイント。
- [data/](data/) — 月次営業日報Excelファイル（2024年8月〜2026年現在）

## 技術スタック

- Python 3.x
- matplotlib / mpl_toolkits（3Dグラフ・ネオングロウグラフ）
- python-pptx（PowerPoint生成）
- numpy
- openpyxl / pandas（Excelデータ読み込み用途）

## スクリプトの実行方法

```bash
cd template
python dashboard_template.py
# → my_dashboard.pptx が生成される
```

## カスタマイズポイント（dashboard_template.py）

| タグ | 内容 |
|------|------|
| `[CONFIG] 基本設定` | タイトル・サブタイトル・出力ファイル名 |
| `[CONFIG] カラーパレット` | 全体の配色 |
| `[CONFIG] KPI カード` | KPI名と値（最大5個） |
| `[CONFIG] 外的要因イベント` | イベントラベル（最大9個） |
| `[CONFIG] 季節性ラベル` | 12ヶ月分の季節指数（超/高/中/低） |
| `[CONFIG] データ読み込み` | `load_data()` 関数 — 実データへの差し替え箇所 |
| `[CONFIG] チャート定義` | 各 `chart_N()` 関数 — グラフ内容の差し替え箇所 |

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
   - 例: `KPIカードに売上達成率を追加`, `load_data()を実データExcelに切り替え`

3. **大きなExcelファイルはGit管理対象外**
   - `data/` 配下のExcelファイルは `.gitignore` で除外する。
   - スクリプト（.py）とテンプレート（.pptx）のみバージョン管理する。

4. **`.gitignore` 推奨設定**
   ```
   data/
   *.xlsx
   __pycache__/
   *.pyc
   my_dashboard.pptx
   ```

5. **ブランチ戦略**
   - 通常作業は `main` ブランチで直接運用。
   - 大きな機能追加や実験的な変更は `feature/XXX` ブランチを切る。

### 注意事項

- プッシュ前に `git status` で変更内容を確認する。
- `.env` や認証情報を含むファイルは絶対にコミットしない。
- 出力ファイル（`my_dashboard.pptx` など）はコミット対象外。
