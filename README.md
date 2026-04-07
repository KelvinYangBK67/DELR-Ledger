# Expenses CLI

一個簡單的記賬工具，會把資料存成 `CSV`，並自動輸出每月 `Markdown`。

## 功能

- 新增一筆記錄
- 依月份統計
- 輸出格式為「日期 + 表格 + 小計 + 月總計」

## 使用方式

在專案根目錄執行：

```powershell
python .\expenses.py add --date 2026-04-07 --item "Dueruem" --amount 8.00 --payment "現金" --merchant "Memetello Grillhouse" --category "餐飲"
```

如果不帶 `--date`，會自動用今天日期。

查詢月統計：

```powershell
python .\expenses.py summary --month 2026-04
```

匯入既有月報（像你現在的 `2026/2026-04.md`）：

```powershell
python .\expenses.py import-md --month 2026-04 --md-path .\2026\2026-04.md
```

## 檔案結構

- 原始資料：`data/YYYY-MM.csv`
- 月報表：`YYYY/YYYY-MM.md`

例如 `2026-04`：

- `data/2026-04.csv`
- `2026/2026-04.md`
