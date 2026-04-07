# DELR Ledger

[Englsih](README.md)

DELR Ledger 是一個輕量的桌面記賬工具（Tkinter），支援多語系介面，並以本地檔案為核心保存資料。

本專案在 Codex 的輔助下完成。

## 版本

目前版本：`v0.1.0`

## 主要功能

- Python + Tkinter 桌面 GUI
- 介面語言：`繁體中文`、`English`、`Deutsch`
- 賬本格式：`.delr`（內容相容 CSV）
- 可新建 / 開啟 / 導入 / 導出賬本
- 支援導入/導出格式：
  - `.delr`
  - `.csv`
  - `.tsv`
  - `.xlsx`
  - `.json`
  - `.xml`
  - `.yaml` / `.yml`
- 底部按貨幣分別統計：總計 / 收入 / 支出
- 表頭篩選（點擊啟用，再點擊可取消）

## 專案結構

- `delr_ledger_app.py`：主 GUI 程式
- `expenses.py`：早期 CLI 工具（保留）
- `scripts/`：打包與發佈腳本
- `VERSION`：版本號
- `LICENSE`：MIT 授權

執行或打包後產生目錄：

- `dist/`：輸出結果
- `build/`：中間構建產物

## 開發啟動

1. 建立並啟用虛擬環境
2. 安裝依賴
3. 執行 GUI

PowerShell 範例：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl pyyaml
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
- 正式發佈可生成帶版本號 zip

## 資料說明

- 預設資料夾可在程式內設定
- `.delr` 檔案以表格儲存，欄位如下：
  - `date, amount, item, unit, payment, merchant, category`
- `settings.json` 用於儲存應用偏好（語言、上次路徑、上次檔案等）

## 可選依賴

部分格式需要額外套件：

- `XLSX`：`openpyxl`
- `YAML`：`PyYAML`

若缺少依賴，程式在使用對應格式時會提示錯誤。

## 授權

MIT License，見 [LICENSE](LICENSE)。


