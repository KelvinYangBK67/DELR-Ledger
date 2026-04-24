# DELR Ledger

[English](README.md)

DELR Ledger 是一個輕量的桌面記賬工具（Tkinter），支援多語系介面，並以本地檔案為核心保存資料。

本專案在 Codex 的協助下完成。

## 版本

目前版本：`v0.1.1`

## 主要功能

- Python + Tkinter 桌面 GUI
- 介面語言：`繁體中文`、`English`、`Deutsch`
- 賬本格式：`.delr`（相容 CSV）
- 可新建、開啟、導入、導出賬本
- 支援剪貼板智能匯入
- 支援賬本導入/導出格式：
  - `.delr`
  - `.csv`
  - `.tsv`
  - `.xlsx`
  - `.json`
  - `.xml`
  - `.yaml` / `.yml`
- 支援依照目前表格視圖導出文檔：
  - `.md`
  - `.docx`
  - `.pdf`
- 預設條目排序為日期優先；同日期下按商家分組，商家順序依照該日期資料中首次出現順序
- 可勾選「不記入收支」，用於轉賬、取現等需要保留記錄但不影響統計的條目
- 底部按貨幣分別統計總計 / 收入 / 支出
- 表頭排序與篩選

## 專案結構

- `delr_ledger_app.py`：主 GUI 程式
- `expenses.py`：早期 CLI 工具（保留）
- `scripts/`：打包與發佈腳本
- `VERSION`：版本號
- `LICENSE`：MIT 授權

執行或打包後產生的資料夾：

- `dist/`：輸出結果
- `build/`：中間建置產物

## 開發啟動

1. 建立並啟用虛擬環境。
2. 安裝依賴。
3. 執行 GUI。

PowerShell 範例：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl pyyaml python-docx reportlab
python .\delr_ledger_app.py
```

## 打包 EXE

沿用目前 `scripts/` 下既有腳本：

```powershell
.\scripts\build.bat
```

或發佈流程：

```powershell
.\scripts\release.bat
```

說明：

- 發佈產物位於 `dist/`
- 正式發佈可生成帶版本號的 zip

## 資料說明

- 預設資料夾可在程式內設定。
- `.delr` 檔案以表格形式儲存，欄位如下：
  - `date, amount, item, unit, payment, merchant, category, excluded`
- 舊檔案沒有 `excluded` 欄位也可正常讀取。`1`、`true`、`yes`、`y`、`on` 會被視為「不記入收支」。
- `settings.json` 用於儲存應用偏好（語言、上次路徑、上次檔案等）。

## 可選依賴

部分格式需要額外套件：

- `XLSX` 支援：`openpyxl`
- `YAML` 支援：`PyYAML`
- `DOCX` 文檔導出：`python-docx`
- `PDF` 文檔導出：`reportlab`

若缺少依賴，程式在使用對應格式時會顯示錯誤訊息。

## 授權

MIT License，詳見 [LICENSE](LICENSE)。
