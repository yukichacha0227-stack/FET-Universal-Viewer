# FET Universal Viewer

**FET（電界効果トランジスタ）の測定データ解析・グラフ化自動化ツール**

研究室での実験データを「一括で読み込み」「整理してExcelに結合」「論文品質のグラフを自動生成」までワンストップで行うPythonアプリケーションです。

![App Screenshot](images/demo.png) ## 🚀 主な機能 (Features)

* **データ一括処理 (Batch Processing)**
    * 複数の測定データ（`.Dat`, `.txt`, `.csv`）を一度に選択し、1つのExcelファイルに結合します。
    * ファイル間に空白列を自動挿入し、視認性を高めます。
    * ヘッダーがないファイルには自動で列名（No., Temp, Mag, Isd, Vsd, Vbg）を補完します。

* **Excelグラフ完全自動化 (Native Excel Charting)**
    * Pythonで画像を作るのではなく、**Excelの機能（散布図）を使ってグラフを自動生成**します。
    * 生成後のグラフはExcel上で自由に編集可能です。

* **ヒステリシス解析 (Hysteresis Analysis)**
    * **順掃引 (Forward)** と **逆掃引 (Reverse)** を自動判定。
    * **青色（行き）** と **オレンジ色（帰り）** に色分けしてプロットします。

* **論文仕様のデザイン (Publication-Ready Design)**
    * **内向き目盛り (Inside Ticks)**
    * **指数表記 (Scientific Notation)**
    * **枠線のみ・グリッドなし (Box Style, No Grid)**
    * 横軸・縦軸の交点をグラフ枠の端に固定。

* **強力なデータクリーニング**
    * 測定器特有のオーバーフロー表記（`#######`）や文字化けを自動で除去・修復して読み込みます。

## 🛠 動作要件 (Requirements)

Windows環境（Excelがインストールされていること）が必要です。

* Python 3.x
* Microsoft Excel

### 必要なライブラリ
```bash
pip install pandas numpy matplotlib openpyxl pywin32
