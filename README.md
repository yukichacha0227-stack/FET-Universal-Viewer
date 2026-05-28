# FET Universal Viewer

FET（電界効果トランジスタ）の測定データを一括で読み込み、Excelブックへの統合とグラフ作成を支援するデスクトップアプリです。

研究室で発生する `.Dat` / `.txt` / `.csv` 形式の測定ファイルを、手作業で整形してから作図する負担を減らすことを目的にしています。Pythonでデータを正規化し、プレビューで傾向を確認したうえで、Excel上で編集できるグラフを出力できます。

## このリポジトリで示すこと

- 実験装置由来のタブ区切りデータを読み込み、列名ゆれや不要列を補正する処理
- 複数ファイルをまとめてExcelブックへ出力するバッチ処理
- FETの出力特性・伝達特性を切り替えて確認できるGUI
- 掃引方向を分けて可視化するヒステリシス確認用のプロット
- Windows + Microsoft Excel環境で、Excel編集可能な散布図を自動生成する処理

## 主な機能

- 複数の `.Dat` / `.txt` / `.csv` ファイルを一括読み込み
- `温度` / `磁場` などの日本語ヘッダーを `Temp` / `Mag` に正規化
- 必須列 `Isd` / `Vsd` / `Vbg` を数値化し、欠損値やオーバーフロー表記を除外
- `Output Characteristics (Isd - Vsd)` と `Transfer Characteristics (Isd - Vbg)` のプレビュー
- 線形スケールと `|Isd|` のログスケール表示
- Excelブックに `Merged_Data` と `Raw_Data` シートを出力
- Excel COM Automationによる編集可能なグラフ作成

## 動作環境

- Python 3.10以降
- Windows
- Microsoft Excel（Excelグラフ自動生成を使う場合）

PythonプレビューとExcelブック出力はPythonだけで動作します。Excelグラフ自動生成のみ、Windows版Microsoft Excelと `pywin32` が必要です。

## セットアップ

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 使い方

```bash
python FET-Universal-Viewer.py
```

1. `Select Files & Merge` から測定データを複数選択します。
2. 統合されたExcelブックが、選択したデータと同じフォルダに作成されます。
3. `Preview (Python)` でグラフを確認します。
4. 必要に応じて `Create Excel Graphs (Japanese Legend)` でExcel上に編集可能なグラフを作成します。

## サンプルデータ

`sample_data/` に動作確認用の `.Dat` ファイルを同梱しています。これらは実際のFET測定ファイルと同じ形式のサンプルとして、読み込み処理とテストで使用します。

想定する主要列は次の通りです。

| Column | Meaning |
| --- | --- |
| `No.` | 測定点番号 |
| `Temp` | 温度 |
| `Mag` | 磁場 |
| `Isd` | ソース-ドレイン電流 |
| `Vsd` | ソース-ドレイン電圧 |
| `Vbg` | バックゲート電圧 |

## テスト

```bash
python -m unittest discover
```

テストでは `sample_data/*.Dat` を読み込み、列名の正規化、数値変換、複数ファイルの結合、Excel出力を確認します。

## 構成

```text
FET-Universal-Viewer.py      # Tkinter GUI entry point
FET-Universal-Viewer.ipynb   # Data loader demo notebook
src/data_loader.py           # Dat/csv/txt reader and Excel export utilities
sample_data/                 # Example Dat files for tests and demos
tests/test_data_loader.py    # Regression tests for the data loader
requirements.txt             # Runtime dependencies
LICENSE                      # MIT License
```

## ライセンス

MIT License
