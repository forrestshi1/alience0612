# alience0612

総務省 電波利用ポータル（無線局等情報検索）から営業リストを作成するスクレイピング＆エンリッチメントパイプライン。

## 構成

- `tele_soumu_full_pipeline.py` — メインパイプライン（収集 → 名寄せ → 企業情報補完 → 納品形式エクスポート）
- `tele_soumu_oneclick_enrich.py` — 企業情報補完エンジン（法人番号公表サイト / gBizINFO / IRBank / 公式サイトクロール / PDF解析）
- `_count_check*.py` / `_probe*.py` — 検索条件ごとの件数・ページ構造の調査スクリプト

## 使い方

```bash
# ① 全検索条件の収集（約1時間、中断後は再開可能）
python tele_soumu_full_pipeline.py collect

# ② 企業情報の補完（並列実行、中断後は再開可能）
python tele_soumu_full_pipeline.py enrich --workers 12

# ③ 納品ファイル出力（pipeline_out/deliverable.xlsx）
python tele_soumu_full_pipeline.py export
```

gBizINFO REST API のトークンを環境変数 `GBIZINFO_API_TOKEN` に設定すると、資本金・従業員数・決算月の取得率が向上します。

## 依存パッケージ

```bash
pip install requests beautifulsoup4 pandas openpyxl PyPDF2
```

データ出力（CSV/XLSX）はリポジトリに含まれません。
