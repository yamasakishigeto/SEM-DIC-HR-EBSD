# SEM-DIC / HR-EBSD Tools

SEM-DIC と HR-EBSD を併用した実験のための Python スクリプト集です。前処理済みデータ（pre-processed *.mat / *.xlsx）の読み込み、パターン置換、可視化、データ変換（.mat→.xlsx）などを含みます。

> ⚠️ **大容量データについて**  
> 数GB規模の `*.tif` や `*.mat` は Git に含めず、ローカルの `data/` などで管理してください。> 追跡除外のための `.gitignore` 例は下部を参照。

---

## 📦 主要スクリプトと役割

### EBSD PatRep/ 配下
- **pattern_replacer_allpoints_batch_250709.py**  
  - 0th / nth の前処理済みファイル群を指定し、**位相（phase）ごとに結晶対称群を選択**して    参照パターンとの **misorientation マッチング** を一括実行。    マッチ結果をCSVに出力し、対応するTIFを `replacing_*/renamed_*/replaced_*` へコピー。    最後に **グレインマップを一括可視化**します。｛GUI: Tkinter｝
- **preprocessed_loader.py**  
  - `pre-processed {nth}*.xlsx / *.mat` をパターン検索で読み込むユーティリティ。    Excelから **ラベルに“ゆるく一致”する値を取得** する `get_value_by_label` も提供。
- **reference_search_module_allpoints_250709.py**  
  - `pattern_replacer_allpoints_batch` から呼ばれる **参照点抽出** / **マッチング** の実装（関数名：    `extract_target_points`, `run_misorientation_matching_all_vs_targets` など）。
- **visualize_grain_map_overlay_250709.py**  
  - `visualize_grain_map(mat_path, excel_path, csv_path)` を想定。    マッチ結果と前処理データを用いて **グレインマップを表示・保存**。

### ルート直下
- **mat_to_excel_batch_exporter_250828.py**  
  - フォルダ内の `*.mat` を一括読み込みし、選択した変数を列として統合、**Excel（.xlsx）へ書き出し**。    変数長の不一致がある場合の扱い（スキップ／NaNパディング等）はコード設定に従います。｛GUI: Tkinter｝
- **stress_strain_mapper_250828.py**  
  - 前処理済みExcelを読み込み、**散布図（粒IDカラー）** と **応力–ひずみ曲線** を連動表示する    インタラクティブ可視化ツール。クリック／ホバー切替、境界線の表示切替などに対応。

---

## 🧩 依存関係（例）

- Python 3.9+（推奨）
- numpy / pandas / scipy / matplotlib / openpyxl / tkinter（標準）

仮想環境のセットアップ例：
```bash
python -m venv .venv
# Windows PowerShell
.venv\Scripts\Activate.ps1
pip install -U pip
pip install numpy pandas scipy matplotlib openpyxl
```

---

## 🚀 使い方（クイック）

### 1) .mat → .xlsx に一括変換
```bash
python mat_to_excel_batch_exporter_250828.py
```
- ダイアログで `*.mat` のあるフォルダを選択 → 変数を複数選択して出力。

### 2) 参照パターン置換（全点・バッチ）
```bash
python "EBSD PatRep/pattern_replacer_allpoints_batch_250709.py"
```
- 0th を含む **親フォルダ**を選択 → **nth フォルダを複数選択** → **phaseごとに対称群を選択** →   misorientation マッチング → CSV出力 & TIFコピー → **グレインマップ可視化**まで自動実行。

### 3) 応力–ひずみ可視化
```bash
# 大規模データでは hover より click 推奨
python stress_strain_mapper_250828.py --mode click
# 直接ファイルを渡す場合（例）
python stress_strain_mapper_250828.py --file "E:\path\to\preprocessed.xlsx" --mode click
```

---

## 📂 推奨ディレクトリ構成
```
SEM-DIC-HR-EBSD/
├─ EBSD PatRep/
├─ data/                 # 大容量データ（Git管理外）
├─ README.md
└─ .gitignore
```

**.gitignore 例**
```
data/
*.tif
*.h5
*.npy
*.mat
__pycache__/
*.pyc
.venv/
```

---

## 📝 メモ
- `pattern_replacer_allpoints_batch_250709.py` では、位相ごとの **結晶対称群（cubic/hexagonal/...）** を   GUIで選択し、`scipy.spatial.transform.Rotation` の **群演算** を用いて対称操作を生成します。  マッチ結果は CSV へ出力され、TIF は最寄座標検索によりコピー＆置換されます。
- `preprocessed_loader.py` のラベル検索は、**空白やアンダースコアを無視**する緩い一致で隣セルを返すため、  表記ゆれに強い仕様です。

---

## 🙋 連絡
- Author: Shigeto Yamasaki (@yamasakishigeto)
- Issue / Pull Request による改善提案歓迎（Private の場合は Collaborator 招待にて）
