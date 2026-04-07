#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.etree.ElementTree import Element, SubElement, ElementTree, parse as xml_parse

try:
    import yaml  # type: ignore
except Exception:
    yaml = None

try:
    from openpyxl import Workbook, load_workbook  # type: ignore
except Exception:
    Workbook = None
    load_workbook = None

CSV_HEADERS = ["date", "amount", "item", "unit", "payment", "merchant", "category"]
ISO_CURRENCIES = sorted({
    "AED","AFN","ALL","AMD","ANG","AOA","ARS","AUD","AWG","AZN","BAM","BBD","BDT","BGN","BHD","BIF",
    "BMD","BND","BOB","BRL","BSD","BTN","BWP","BYN","BZD","CAD","CDF","CHF","CLP","CNY","COP","CRC",
    "CUP","CVE","CZK","DJF","DKK","DOP","DZD","EGP","ERN","ETB","EUR","FJD","FKP","FOK","GBP","GEL",
    "GGP","GHS","GIP","GMD","GNF","GTQ","GYD","HKD","HNL","HRK","HTG","HUF","IDR","ILS","IMP","INR",
    "IQD","IRR","ISK","JEP","JMD","JOD","JPY","KES","KGS","KHR","KID","KMF","KRW","KWD","KYD","KZT",
    "LAK","LBP","LKR","LRD","LSL","LYD","MAD","MDL","MGA","MKD","MMK","MNT","MOP","MRU","MUR","MVR",
    "MWK","MXN","MYR","MZN","NAD","NGN","NIO","NOK","NPR","NZD","OMR","PAB","PEN","PGK","PHP","PKR",
    "PLN","PYG","QAR","RON","RSD","RUB","RWF","SAR","SBD","SCR","SDG","SEK","SGD","SHP","SLE","SLL",
    "SOS","SRD","SSP","STN","SYP","SZL","THB","TJS","TMT","TND","TOP","TRY","TTD","TVD","TWD","TZS",
    "UAH","UGX","USD","UYU","UZS","VED","VES","VND","VUV","WST","XAF","XCD","XCG","XDR","XOF","XPF",
    "YER","ZAR","ZMW","ZWL"
})

I18N = {
    "en": {
        "app_title": "DELR Ledger", "language": "Language", "data_folder": "Data Folder", "choose_folder": "Choose Folder", "current_file": "Current Ledger",
        "new": "New (.delr)", "open": "Open (.delr)", "export": "Export", "import": "Import", "type": "Type", "income": "Income", "expense": "Expense",
        "date": "Date", "item": "Item", "category": "Category", "amount": "Amount", "unit": "Unit", "payment": "Payment", "merchant": "Merchant",
        "add": "Add", "clear": "Clear", "delete": "Delete", "entries": "Ledger Entries", "total": "Total", "income_total": "Income", "expense_total": "Expense",
        "filter": "Filter", "from": "From", "to": "To", "invalid_item": "Item cannot be empty.", "invalid_amount": "Amount must be a number.",
        "invalid_date": "Invalid date.", "no_select": "Please select at least one row.", "apply": "Apply", "cancel": "Cancel", "contains": "Contains"
    },
    "zh-TW": {
        "app_title": "DELR 記賬本", "language": "語言", "data_folder": "資料夾", "choose_folder": "選擇資料夾", "current_file": "目前賬本",
        "new": "新建（.delr）", "open": "開啟（.delr）", "export": "導出", "import": "導入", "type": "類型", "income": "收入", "expense": "支出",
        "date": "日期", "item": "項目", "category": "分類", "amount": "金額", "unit": "單位", "payment": "支付方式", "merchant": "商家",
        "add": "新增", "clear": "清空", "delete": "刪除", "entries": "賬本條目", "total": "總計", "income_total": "收入", "expense_total": "支出",
        "filter": "篩選", "from": "起", "to": "迄", "invalid_item": "項目不可為空。", "invalid_amount": "金額必須是數字。",
        "invalid_date": "日期無效。", "no_select": "請至少選擇一行。", "apply": "套用", "cancel": "取消", "contains": "包含"
    },
    "de": {
        "app_title": "DELR Kassenbuch", "language": "Sprache", "data_folder": "Datenordner", "choose_folder": "Ordner wählen", "current_file": "Aktuelles Ledger",
        "new": "Neu (.delr)", "open": "Öffnen (.delr)", "export": "Exportieren", "import": "Importieren", "type": "Typ", "income": "Einnahme", "expense": "Ausgabe",
        "date": "Datum", "item": "Artikel", "category": "Kategorie", "amount": "Betrag", "unit": "Einheit", "payment": "Zahlung", "merchant": "Händler",
        "add": "Hinzufügen", "clear": "Leeren", "delete": "Löschen", "entries": "Ledger-Einträge", "total": "Summe", "income_total": "Einnahmen", "expense_total": "Ausgaben",
        "filter": "Filter", "from": "Von", "to": "Bis", "invalid_item": "Artikel darf nicht leer sein.", "invalid_amount": "Betrag muss eine Zahl sein.",
        "invalid_date": "Ungültiges Datum.", "no_select": "Bitte mindestens eine Zeile wählen.", "apply": "Anwenden", "cancel": "Abbrechen", "contains": "Enthält"
    },
}

LANG_DISPLAY_TO_CODE = {"繁體中文": "zh-TW", "English": "en", "Deutsch": "de"}
CASH_LABELS = {"zh-TW": "現金", "en": "Cash", "de": "Bar"}
ALL_TYPE_LABELS = {
    "income": {"income", "收入", "einnahme"},
    "expense": {"expense", "支出", "ausgabe"},
}


@dataclass
class Entry:
    date: str
    amount: float
    item: str
    unit: str
    payment: str
    merchant: str
    category: str


class DelrLedgerApp:
    def __init__(self, root: tk.Tk, app_dir: Path) -> None:
        self.root = root
        self.app_dir = app_dir
        self.user_dir = app_dir / "user"
        self.user_dir.mkdir(parents=True, exist_ok=True)
        self.config_dir = app_dir / "config"
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.settings_path = self.config_dir / "settings.json"

        now = datetime.now()
        self.lang_display_var = tk.StringVar(value="繁體中文")
        self.folder_var = tk.StringVar(value=str(self.user_dir))
        self.file_var = tk.StringVar(value="")
        self.type_var = tk.StringVar(value="expense")
        self.y_var = tk.StringVar(value=str(now.year))
        self.m_var = tk.StringVar(value=f"{now.month:02d}")
        self.d_var = tk.StringVar(value=f"{now.day:02d}")
        self.item_var = tk.StringVar()
        self.category_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.unit_var = tk.StringVar(value="EUR")
        self.payment_var = tk.StringVar(value="")
        self.merchant_var = tk.StringVar()

        self.entries: list[Entry] = []
        self.current_file: Path | None = None

        self.header_filters: dict[str, str] = {}
        self.header_range: dict[str, tuple[str, str]] = {}
        self.header_multi: dict[str, set[str]] = {}

        self.payment_values: set[str] = set()
        self.category_values: set[str] = set()
        self.merchant_values: set[str] = set()
        self.unit_values: set[str] = set(ISO_CURRENCIES)

        self._build_ui()
        self._load_settings()
        self.apply_language()
        self._open_last()
        self._update_days()

    def code(self) -> str:
        return LANG_DISPLAY_TO_CODE.get(self.lang_display_var.get(), "zh-TW")

    def tr(self, key: str) -> str:
        return I18N[self.code()][key]

    def cash(self) -> str:
        return CASH_LABELS[self.code()]

    def ui_date_fmt(self) -> str:
        return "%m-%d-%Y" if self.code() == "en" else ("%d-%m-%Y" if self.code() == "de" else "%Y-%m-%d")

    def fmt_ui_date(self, iso: str) -> str:
        return datetime.strptime(iso, "%Y-%m-%d").strftime(self.ui_date_fmt())
    def _build_ui(self) -> None:
        self.root.title("DELR Ledger")
        self.root.geometry("1280x760")
        self.root.minsize(1120, 620)

        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=(10, 6))
        self.lang_label = ttk.Label(top)
        self.lang_label.pack(side="left")
        self.lang_combo = ttk.Combobox(top, textvariable=self.lang_display_var, values=["繁體中文", "English", "Deutsch"], state="readonly", width=12)
        self.lang_combo.pack(side="left", padx=(6, 14))
        self.lang_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_lang())
        self.folder_label = ttk.Label(top)
        self.folder_label.pack(side="left")
        ttk.Entry(top, textvariable=self.folder_var, width=42).pack(side="left", padx=6)
        self.choose_btn = ttk.Button(top, command=self.choose_folder)
        self.choose_btn.pack(side="left")

        file_row = ttk.Frame(self.root)
        file_row.pack(fill="x", padx=10, pady=(0, 6))
        self.file_label = ttk.Label(file_row)
        self.file_label.pack(side="left")
        ttk.Entry(file_row, textvariable=self.file_var).pack(side="left", fill="x", expand=True, padx=6)
        self.new_btn = ttk.Button(file_row, command=self.new_ledger)
        self.new_btn.pack(side="left")
        self.open_btn = ttk.Button(file_row, command=self.open_ledger)
        self.open_btn.pack(side="left", padx=(6, 0))
        self.export_btn = ttk.Button(file_row, command=self.export_ledger)
        self.export_btn.pack(side="left", padx=(6, 0))
        self.import_btn = ttk.Button(file_row, command=self.import_merge)
        self.import_btn.pack(side="left", padx=(6, 0))

        frm = ttk.LabelFrame(self.root)
        frm.pack(fill="x", padx=10, pady=(0, 6))
        self.form_frame = frm

        r1 = ttk.Frame(frm)
        r1.pack(fill="x", padx=8, pady=(6, 2))
        self.type_label = ttk.Label(r1)
        self.type_label.pack(side="left")
        self.type_combo = ttk.Combobox(r1, textvariable=self.type_var, state="readonly", width=10)
        self.type_combo.pack(side="left", padx=(4, 14))

        self.date_label = ttk.Label(r1)
        self.date_label.pack(side="left")
        self.date_wrap = ttk.Frame(r1)
        self.date_wrap.pack(side="left", padx=(4, 14))
        self.y_combo = ttk.Combobox(self.date_wrap, textvariable=self.y_var, values=[str(y) for y in range(2000, 2201)], width=6, state="readonly")
        self.y_u = ttk.Label(self.date_wrap)
        self.m_combo = ttk.Combobox(self.date_wrap, textvariable=self.m_var, values=[f"{m:02d}" for m in range(1, 13)], width=4, state="readonly")
        self.m_u = ttk.Label(self.date_wrap)
        self.d_combo = ttk.Combobox(self.date_wrap, textvariable=self.d_var, values=[f"{d:02d}" for d in range(1, 32)], width=4, state="readonly")
        self.d_u = ttk.Label(self.date_wrap)
        self.y_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_days())
        self.m_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_days())

        self.item_label = ttk.Label(r1)
        self.item_label.pack(side="left")
        ttk.Entry(r1, textvariable=self.item_var, width=20).pack(side="left", padx=(4, 14))

        self.category_label = ttk.Label(r1)
        self.category_label.pack(side="left")
        self.category_combo = ttk.Combobox(r1, textvariable=self.category_var, values=[], width=16)
        self.category_combo.pack(side="left", padx=(4, 8))

        r2 = ttk.Frame(frm)
        r2.pack(fill="x", padx=8, pady=(2, 6))
        self.amount_label = ttk.Label(r2)
        self.amount_label.pack(side="left")
        ttk.Entry(r2, textvariable=self.amount_var, width=12).pack(side="left", padx=(4, 14))

        self.unit_label = ttk.Label(r2)
        self.unit_label.pack(side="left")
        self.unit_combo = ttk.Combobox(r2, textvariable=self.unit_var, values=sorted(self.unit_values), width=8)
        self.unit_combo.pack(side="left", padx=(4, 14))
        self.unit_combo.bind("<KeyRelease>", lambda _e: self._filter_combo_values(self.unit_combo, sorted(self.unit_values)))
        self.unit_combo.bind("<FocusOut>", lambda _e: self._enforce_combo_legal(self.unit_combo, sorted(self.unit_values), "EUR"))

        self.payment_label = ttk.Label(r2)
        self.payment_label.pack(side="left")
        self.payment_combo = ttk.Combobox(r2, textvariable=self.payment_var, values=[], width=16)
        self.payment_combo.pack(side="left", padx=(4, 14))

        self.merchant_label = ttk.Label(r2)
        self.merchant_label.pack(side="left")
        self.merchant_combo = ttk.Combobox(r2, textvariable=self.merchant_var, values=[], width=22)
        self.merchant_combo.pack(side="left", padx=(4, 8))

        ttk.Frame(r2).pack(side="left", fill="x", expand=True)
        self.add_btn = ttk.Button(r2, command=self.add_entry)
        self.add_btn.pack(side="right")
        self.clear_btn = ttk.Button(r2, command=self.clear_form)
        self.clear_btn.pack(side="right", padx=(0, 6))

        table_box = ttk.LabelFrame(self.root)
        table_box.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.table_box = table_box
        hdr = ttk.Frame(table_box)
        hdr.pack(fill="x", padx=8, pady=(6, 4))
        ttk.Frame(hdr).pack(side="left", fill="x", expand=True)
        self.del_btn = ttk.Button(hdr, command=self.delete_selected)
        self.del_btn.pack(side="right")

        tw = ttk.Frame(table_box)
        tw.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("date", "type", "item", "amount", "unit", "payment", "merchant", "category")
        self.tree = ttk.Treeview(tw, columns=cols, show="headings", selectmode="extended")
        widths = {"date": 120, "type": 90, "item": 220, "amount": 100, "unit": 70, "payment": 130, "merchant": 180, "category": 130}
        for c in cols:
            self.tree.column(c, width=widths[c], anchor="w")
            self.tree.heading(c, text=c, command=lambda col=c: self.on_header_click(col))
        y_scroll = ttk.Scrollbar(tw, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="left", fill="y")

        st = ttk.Frame(self.root)
        st.pack(side="bottom", fill="x", padx=10, pady=(0, 10))
        self.total_label = ttk.Label(st)
        self.total_label.pack(anchor="w")
        self.income_label = ttk.Label(st)
        self.income_label.pack(anchor="w")
        self.expense_label = ttk.Label(st)
        self.expense_label.pack(anchor="w")

    def _layout_date(self) -> None:
        for child in self.date_wrap.winfo_children():
            child.pack_forget()
        if self.code() == "en":
            items = [(self.m_combo, "M"), (self.d_combo, "D"), (self.y_combo, "Y")]
        elif self.code() == "de":
            items = [(self.d_combo, "T"), (self.m_combo, "M"), (self.y_combo, "J")]
        else:
            items = [(self.y_combo, "年"), (self.m_combo, "月"), (self.d_combo, "日")]
        labels = [self.y_u, self.m_u, self.d_u]
        for (combo, text), lbl in zip(items, labels):
            combo.pack(side="left", padx=(0, 1))
            lbl.config(text=text)
            lbl.pack(side="left", padx=(0, 2))

    def _load_settings(self) -> None:
        if not self.settings_path.exists():
            return
        try:
            settings = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except Exception:
            return
        if settings.get("lang_display") in LANG_DISPLAY_TO_CODE:
            self.lang_display_var.set(settings["lang_display"])
        if isinstance(settings.get("data_folder"), str) and settings["data_folder"].strip():
            self.folder_var.set(settings["data_folder"])
        if isinstance(settings.get("last_file"), str) and settings["last_file"].strip():
            self.file_var.set(settings["last_file"])

    def _save_settings(self) -> None:
        payload = {"lang_display": self.lang_display_var.get(), "data_folder": self.folder_var.get(), "last_file": str(self.current_file) if self.current_file else ""}
        self.settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _type_from_ui(self, value: str) -> str:
        s = (value or "").strip().lower()
        if s in ALL_TYPE_LABELS["income"]:
            return "income"
        return "expense"

    def _filter_combo_values(self, combo: ttk.Combobox, source: list[str]) -> None:
        key = combo.get().strip().upper()
        combo["values"] = source if not key else [v for v in source if key in v.upper()]

    def _enforce_combo_legal(self, combo: ttk.Combobox, source: list[str], fallback: str = "") -> None:
        value = combo.get().strip().upper()
        legal = {x.upper(): x for x in source}
        combo.set(legal[value] if value in legal else fallback)

    def on_lang(self) -> None:
        self.apply_language()
        self.refresh_table()
        self._save_settings()

    def apply_language(self) -> None:
        self.root.title(self.tr("app_title"))
        self.lang_label.config(text=f"{self.tr('language')}: ")
        self.folder_label.config(text=f"{self.tr('data_folder')}: ")
        self.choose_btn.config(text=self.tr("choose_folder"))
        self.file_label.config(text=f"{self.tr('current_file')}: ")
        self.new_btn.config(text=self.tr("new"))
        self.open_btn.config(text=self.tr("open"))
        self.export_btn.config(text=self.tr("export"))
        self.import_btn.config(text=self.tr("import"))
        self.form_frame.config(text=self.tr("entries"))
        self.type_label.config(text=self.tr("type")); self.date_label.config(text=self.tr("date")); self.item_label.config(text=self.tr("item")); self.category_label.config(text=self.tr("category"))
        self.amount_label.config(text=self.tr("amount")); self.unit_label.config(text=self.tr("unit")); self.payment_label.config(text=self.tr("payment")); self.merchant_label.config(text=self.tr("merchant"))
        self.add_btn.config(text=self.tr("add")); self.clear_btn.config(text=self.tr("clear")); self.del_btn.config(text=self.tr("delete"))
        current_type = self._type_from_ui(self.type_var.get())
        self.type_combo["values"] = [self.tr("expense"), self.tr("income")]
        self.type_var.set(self.tr("income") if current_type == "income" else self.tr("expense"))
        current_payment = self.payment_var.get().strip()
        if current_payment and current_payment in set(CASH_LABELS.values()):
            self.payment_var.set(self.cash())
        self._layout_date()
        for col, key in [("date", "date"), ("type", "type"), ("item", "item"), ("amount", "amount"), ("unit", "unit"), ("payment", "payment"), ("merchant", "merchant"), ("category", "category")]:
            self.tree.heading(col, text=f"{self.tr(key)}  [F]", command=lambda c=col: self.on_header_click(c))
        self._rebuild_options()

    def choose_folder(self) -> None:
        picked = filedialog.askdirectory(initialdir=self.folder_var.get())
        if picked:
            self.folder_var.set(picked)
            self._save_settings()

    def _ensure_current(self) -> None:
        if self.current_file is not None:
            return
        folder = Path(self.folder_var.get().strip() or str(self.user_dir))
        folder.mkdir(parents=True, exist_ok=True)
        self.current_file = folder / datetime.now().strftime("ledger-%Y%m%d-%H%M%S.delr")
        self.file_var.set(str(self.current_file))
    def _open_last(self) -> None:
        path = self.file_var.get().strip()
        if path and Path(path).exists() and Path(path).suffix.lower() == ".delr":
            self.current_file = Path(path)
            self.entries = self.read_entries_from_path(self.current_file)
            self._rebuild_options()
            self.refresh_table()

    def _read_rows_csv_like(self, path: Path, delimiter: str = ",") -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f, delimiter=delimiter):
                rows.append({k: (row.get(k, "") or "") for k in CSV_HEADERS})
        return rows

    def _read_rows_json(self, path: Path) -> list[dict[str, str]]:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        rows: list[dict[str, str]] = []
        for it in data:
            if isinstance(it, dict):
                rows.append({k: str(it.get(k, "") or "") for k in CSV_HEADERS})
        return rows

    def _read_rows_xml(self, path: Path) -> list[dict[str, str]]:
        root = xml_parse(path).getroot()
        rows: list[dict[str, str]] = []
        for node in root.findall("entry"):
            rows.append({k: (node.findtext(k, default="") or "") for k in CSV_HEADERS})
        return rows

    def _read_rows_yaml(self, path: Path) -> list[dict[str, str]]:
        if yaml is None:
            raise RuntimeError("PyYAML is not installed")
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        rows: list[dict[str, str]] = []
        for it in data:
            if isinstance(it, dict):
                rows.append({k: str(it.get(k, "") or "") for k in CSV_HEADERS})
        return rows

    def _read_rows_xlsx(self, path: Path) -> list[dict[str, str]]:
        if load_workbook is None:
            raise RuntimeError("openpyxl is not installed")
        wb = load_workbook(filename=path, data_only=True)
        ws = wb.active
        header = [str(c.value or "").strip() for c in ws[1]]
        idx = {h: i for i, h in enumerate(header)}
        rows: list[dict[str, str]] = []
        for r in ws.iter_rows(min_row=2):
            row = {}
            for k in CSV_HEADERS:
                i = idx.get(k)
                row[k] = "" if i is None else str(r[i].value or "")
            rows.append(row)
        return rows

    def _rows_to_entries(self, rows: list[dict[str, str]]) -> list[Entry]:
        out: list[Entry] = []
        for r in rows:
            try:
                amount = float(str(r.get("amount", "0") or "0"))
            except ValueError:
                amount = 0.0
            out.append(Entry(date=str(r.get("date", "1970-01-01") or "1970-01-01"), amount=amount, item=str(r.get("item", "") or ""), unit=str(r.get("unit", "EUR") or "EUR").upper(), payment=str(r.get("payment", "") or ""), merchant=str(r.get("merchant", "") or ""), category=str(r.get("category", "") or "")))
        return out

    def read_entries_from_path(self, path: Path) -> list[Entry]:
        ext = path.suffix.lower()
        if ext in {".delr", ".csv"}:
            rows = self._read_rows_csv_like(path, ",")
        elif ext == ".tsv":
            rows = self._read_rows_csv_like(path, "\t")
        elif ext == ".json":
            rows = self._read_rows_json(path)
        elif ext == ".xml":
            rows = self._read_rows_xml(path)
        elif ext in {".yaml", ".yml"}:
            rows = self._read_rows_yaml(path)
        elif ext == ".xlsx":
            rows = self._read_rows_xlsx(path)
        else:
            raise RuntimeError(f"Unsupported format: {ext}")
        return self._rows_to_entries(rows)

    def _rows(self) -> list[dict[str, str]]:
        return [{"date": e.date, "amount": f"{e.amount:.2f}", "item": e.item, "unit": e.unit, "payment": e.payment, "merchant": e.merchant, "category": e.category} for e in sorted(self.entries, key=lambda x: (x.date, x.item.lower()))]

    def _write_csv_like(self, path: Path, delimiter: str = ",") -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=CSV_HEADERS, delimiter=delimiter)
            w.writeheader()
            for row in self._rows():
                w.writerow(row)

    def _write_json(self, path: Path) -> None:
        path.write_text(json.dumps(self._rows(), ensure_ascii=False, indent=2), encoding="utf-8")

    def _write_xml(self, path: Path) -> None:
        root = Element("ledger")
        for row in self._rows():
            node = SubElement(root, "entry")
            for k, v in row.items():
                c = SubElement(node, k)
                c.text = str(v)
        ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)

    def _write_yaml(self, path: Path) -> None:
        if yaml is None:
            raise RuntimeError("PyYAML is not installed")
        path.write_text(yaml.safe_dump(self._rows(), allow_unicode=True, sort_keys=False), encoding="utf-8")

    def _write_xlsx(self, path: Path) -> None:
        if Workbook is None:
            raise RuntimeError("openpyxl is not installed")
        wb = Workbook()
        ws = wb.active
        ws.append(CSV_HEADERS)
        for row in self._rows():
            ws.append([row[k] for k in CSV_HEADERS])
        wb.save(path)

    def write_current_delr(self) -> None:
        if self.current_file is not None:
            self._write_csv_like(self.current_file, ",")

    def new_ledger(self) -> None:
        folder = Path(self.folder_var.get().strip() or str(self.user_dir))
        folder.mkdir(parents=True, exist_ok=True)
        target = filedialog.asksaveasfilename(initialdir=str(folder), defaultextension=".delr", filetypes=[("DELR Ledger", "*.delr")], title=self.tr("new"))
        if not target:
            return
        self.current_file = Path(target).with_suffix(".delr")
        self.file_var.set(str(self.current_file))
        self.entries = []
        self.write_current_delr()
        self._save_settings()
        self.refresh_table()

    def open_ledger(self) -> None:
        folder = Path(self.folder_var.get().strip() or str(self.user_dir))
        target = filedialog.askopenfilename(initialdir=str(folder), filetypes=[("DELR Ledger", "*.delr")], title=self.tr("open"))
        if not target:
            return
        self.current_file = Path(target)
        self.file_var.set(str(self.current_file))
        self.entries = self.read_entries_from_path(self.current_file)
        self._rebuild_options()
        self._save_settings()
        self.refresh_table()

    def import_merge(self) -> None:
        self._ensure_current()
        folder = Path(self.folder_var.get().strip() or str(self.user_dir))
        target = filedialog.askopenfilename(initialdir=str(folder), filetypes=[("All Supported", "*.delr;*.csv;*.tsv;*.xlsx;*.json;*.xml;*.yaml;*.yml"), ("DELR", "*.delr"), ("CSV", "*.csv"), ("TSV", "*.tsv"), ("XLSX", "*.xlsx"), ("JSON", "*.json"), ("XML", "*.xml"), ("YAML", "*.yaml;*.yml")], title=self.tr("import"))
        if not target:
            return
        try:
            incoming = self.read_entries_from_path(Path(target))
        except Exception as exc:
            messagebox.showerror(self.tr("app_title"), str(exc))
            return
        self.entries.extend(incoming)
        self._rebuild_options()
        self.write_current_delr()
        self._save_settings()
        self.refresh_table()

    def export_ledger(self) -> None:
        self._ensure_current()
        folder = Path(self.folder_var.get().strip() or str(self.user_dir))
        folder.mkdir(parents=True, exist_ok=True)
        target = filedialog.asksaveasfilename(initialdir=str(folder), defaultextension=".delr", filetypes=[("DELR", "*.delr"), ("CSV", "*.csv"), ("TSV", "*.tsv"), ("XLSX", "*.xlsx"), ("JSON", "*.json"), ("XML", "*.xml"), ("YAML", "*.yaml")], title=self.tr("export"))
        if not target:
            return
        path = Path(target)
        ext = path.suffix.lower()
        try:
            if ext == ".delr":
                self._write_csv_like(path, ",")
            elif ext == ".csv":
                self._write_csv_like(path, ",")
            elif ext == ".tsv":
                self._write_csv_like(path, "\t")
            elif ext == ".xlsx":
                self._write_xlsx(path)
            elif ext == ".json":
                self._write_json(path)
            elif ext == ".xml":
                self._write_xml(path)
            elif ext in {".yaml", ".yml"}:
                self._write_yaml(path)
            else:
                self._write_csv_like(path.with_suffix(".delr"), ",")
        except Exception as exc:
            messagebox.showerror(self.tr("app_title"), str(exc))

    def _rebuild_options(self) -> None:
        self.payment_values = {self.cash()} | {e.payment for e in self.entries if e.payment and e.payment != "-"}
        self.category_values = {e.category for e in self.entries if e.category and e.category != "-"}
        self.merchant_values = {e.merchant for e in self.entries if e.merchant and e.merchant != "-"}
        self.unit_values = set(ISO_CURRENCIES) | {e.unit for e in self.entries if e.unit}
        self.payment_combo["values"] = sorted(self.payment_values)
        self.category_combo["values"] = sorted(self.category_values)
        self.merchant_combo["values"] = sorted(self.merchant_values)
        self.unit_combo["values"] = sorted(self.unit_values)

    def _update_days(self) -> None:
        y = int(self.y_var.get())
        m = int(self.m_var.get())
        md = 29 if (m == 2 and ((y % 4 == 0 and y % 100 != 0) or (y % 400 == 0))) else (28 if m == 2 else (30 if m in {4, 6, 9, 11} else 31))
        vals = [f"{d:02d}" for d in range(1, md + 1)]
        self.d_combo["values"] = vals
        if self.d_var.get() not in vals:
            self.d_var.set(vals[-1])

    def add_entry(self) -> None:
        if not self.item_var.get().strip():
            messagebox.showerror(self.tr("app_title"), self.tr("invalid_item"))
            return
        try:
            raw = float(self.amount_var.get().strip())
        except ValueError:
            messagebox.showerror(self.tr("app_title"), self.tr("invalid_amount"))
            return
        selected_type = self._type_from_ui(self.type_var.get())
        if raw < 0:
            amount = abs(raw)
            final_type = "expense"
        else:
            final_type = selected_type
            amount = -abs(raw) if final_type == "income" else abs(raw)
        self.type_var.set(self.tr("income") if final_type == "income" else self.tr("expense"))
        unit = self.unit_var.get().strip().upper()
        if unit not in self.unit_values:
            messagebox.showerror(self.tr("app_title"), self.tr("invalid_amount"))
            return
        self._ensure_current()
        self.entries.append(Entry(date=f"{self.y_var.get()}-{self.m_var.get()}-{self.d_var.get()}", amount=amount, item=self.item_var.get().strip(), unit=unit, payment=(self.payment_var.get().strip() or "-"), merchant=(self.merchant_var.get().strip() or "-"), category=(self.category_var.get().strip() or "-")))
        self._rebuild_options()
        self.write_current_delr()
        self._save_settings()
        self.refresh_table()

    def clear_form(self) -> None:
        self.item_var.set("")
        self.amount_var.set("")
        self.payment_var.set("")
        self.merchant_var.set("")
    def _create_date_picker(self, parent: tk.Widget, date_vars: tuple[tk.StringVar, tk.StringVar, tk.StringVar]) -> None:
        yv, mv, dv = date_vars
        wrap = ttk.Frame(parent)
        wrap.pack(side="left", padx=(4, 0))
        yb = ttk.Combobox(wrap, textvariable=yv, values=[str(y) for y in range(2000, 2201)], width=6, state="readonly")
        mb = ttk.Combobox(wrap, textvariable=mv, values=[f"{m:02d}" for m in range(1, 13)], width=4, state="readonly")
        db = ttk.Combobox(wrap, textvariable=dv, values=[f"{d:02d}" for d in range(1, 32)], width=4, state="readonly")

        def sync_days(*_args: object) -> None:
            y = int(yv.get())
            m = int(mv.get())
            md = 29 if (m == 2 and ((y % 4 == 0 and y % 100 != 0) or (y % 400 == 0))) else (28 if m == 2 else (30 if m in {4, 6, 9, 11} else 31))
            vals = [f"{d:02d}" for d in range(1, md + 1)]
            db["values"] = vals
            if dv.get() not in vals:
                dv.set(vals[-1])

        yb.bind("<<ComboboxSelected>>", sync_days)
        mb.bind("<<ComboboxSelected>>", sync_days)
        sync_days()

        order = [(yb, "年"), (mb, "月"), (db, "日")]
        if self.code() == "en":
            order = [(mb, "M"), (db, "D"), (yb, "Y")]
        elif self.code() == "de":
            order = [(db, "T"), (mb, "M"), (yb, "J")]

        for box, label in order:
            box.pack(side="left", padx=(0, 2))
            ttk.Label(wrap, text=label).pack(side="left", padx=(0, 4))

    def _toggle_or_apply_filter(self, col: str, apply_fn) -> None:
        has_filter = col in self.header_filters or col in self.header_range or col in self.header_multi
        if has_filter:
            self.header_filters.pop(col, None)
            self.header_range.pop(col, None)
            self.header_multi.pop(col, None)
            if col in {"amount", "unit"}:
                self.header_filters.pop("unit", None)
                self.header_range.pop("amount", None)
            self.refresh_table()
            return
        apply_fn()

    def on_header_click(self, col: str) -> None:
        if col == "date":
            self._toggle_or_apply_filter(col, lambda: self._open_date_filter(col))
        elif col == "type":
            self._toggle_or_apply_filter(col, lambda: self._open_type_filter(col))
        elif col == "item":
            self._toggle_or_apply_filter(col, lambda: self._open_text_filter(col))
        elif col in {"payment", "merchant", "category"}:
            self._toggle_or_apply_filter(col, lambda c=col: self._open_multi_filter(c))
        elif col in {"amount", "unit"}:
            self._toggle_or_apply_filter("amount", self._open_amount_unit_filter)
        else:
            self._toggle_or_apply_filter(col, lambda: self._open_text_filter(col))

    def _open_filter_window(self, title: str) -> tk.Toplevel:
        w = tk.Toplevel(self.root)
        w.title(title)
        w.geometry("560x280")
        w.minsize(560, 260)
        w.transient(self.root)
        w.grab_set()
        return w

    def _open_date_filter(self, col: str) -> None:
        w = self._open_filter_window(self.tr("filter"))
        now = datetime.now()
        fy, fm, fd = tk.StringVar(value=str(now.year)), tk.StringVar(value=f"{now.month:02d}"), tk.StringVar(value=f"{now.day:02d}")
        ty, tm, td = tk.StringVar(value=str(now.year)), tk.StringVar(value=f"{now.month:02d}"), tk.StringVar(value=f"{now.day:02d}")
        r1 = ttk.Frame(w); r1.pack(fill="x", padx=10, pady=(12, 6))
        ttk.Label(r1, text=f"{self.tr('from')}:").pack(side="left")
        self._create_date_picker(r1, (fy, fm, fd))
        r2 = ttk.Frame(w); r2.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Label(r2, text=f"{self.tr('to')}:").pack(side="left")
        self._create_date_picker(r2, (ty, tm, td))
        btn = ttk.Frame(w); btn.pack(side="bottom", fill="x", padx=10, pady=10)
        def apply() -> None:
            self.header_range[col] = (f"{fy.get()}-{fm.get()}-{fd.get()}", f"{ty.get()}-{tm.get()}-{td.get()}")
            w.destroy(); self.refresh_table()
        ttk.Button(btn, text=self.tr("apply"), command=apply).pack(side="right")
        ttk.Button(btn, text=self.tr("cancel"), command=w.destroy).pack(side="right", padx=(0, 6))

    def _open_type_filter(self, col: str) -> None:
        w = self._open_filter_window(self.tr("filter"))
        values = [self.tr("expense"), self.tr("income")]
        var = tk.StringVar(value=values[0])
        row = ttk.Frame(w); row.pack(fill="x", padx=10, pady=14)
        ttk.Label(row, text=f"{self.tr('type')}:").pack(side="left")
        ttk.Combobox(row, textvariable=var, values=values, state="readonly", width=20).pack(side="left", padx=(6, 0))
        btn = ttk.Frame(w); btn.pack(side="bottom", fill="x", padx=10, pady=10)
        def apply() -> None:
            self.header_filters[col] = self._type_from_ui(var.get())
            w.destroy(); self.refresh_table()
        ttk.Button(btn, text=self.tr("apply"), command=apply).pack(side="right")
        ttk.Button(btn, text=self.tr("cancel"), command=w.destroy).pack(side="right", padx=(0, 6))

    def _open_text_filter(self, col: str) -> None:
        w = self._open_filter_window(self.tr("filter"))
        var = tk.StringVar(value="")
        row = ttk.Frame(w); row.pack(fill="x", padx=10, pady=14)
        ttk.Label(row, text=f"{self.tr('contains')}: ").pack(side="left")
        ttk.Entry(row, textvariable=var, width=40).pack(side="left", padx=(6, 0))
        btn = ttk.Frame(w); btn.pack(side="bottom", fill="x", padx=10, pady=10)
        def apply() -> None:
            value = var.get().strip()
            if value:
                self.header_filters[col] = value
            w.destroy(); self.refresh_table()
        ttk.Button(btn, text=self.tr("apply"), command=apply).pack(side="right")
        ttk.Button(btn, text=self.tr("cancel"), command=w.destroy).pack(side="right", padx=(0, 6))

    def _open_multi_filter(self, col: str) -> None:
        w = self._open_filter_window(self.tr("filter"))
        values = sorted({"payment": self.payment_values, "merchant": self.merchant_values, "category": self.category_values}[col])
        row = ttk.Frame(w); row.pack(fill="both", expand=True, padx=10, pady=10)
        listbox = tk.Listbox(row, selectmode="multiple", exportselection=False)
        listbox.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(row, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scroll.set)
        scroll.pack(side="left", fill="y")
        for v in values:
            listbox.insert("end", v)
        btn = ttk.Frame(w); btn.pack(side="bottom", fill="x", padx=10, pady=10)
        def apply() -> None:
            picks = {listbox.get(i) for i in listbox.curselection()}
            if picks:
                self.header_multi[col] = picks
            w.destroy(); self.refresh_table()
        ttk.Button(btn, text=self.tr("apply"), command=apply).pack(side="right")
        ttk.Button(btn, text=self.tr("cancel"), command=w.destroy).pack(side="right", padx=(0, 6))

    def _open_amount_unit_filter(self) -> None:
        w = self._open_filter_window(self.tr("filter"))
        unit_var = tk.StringVar(value="EUR")
        lo_var = tk.StringVar(value="")
        hi_var = tk.StringVar(value="")
        r1 = ttk.Frame(w); r1.pack(fill="x", padx=10, pady=(12, 6))
        ttk.Label(r1, text=f"{self.tr('unit')}:").pack(side="left")
        u_combo = ttk.Combobox(r1, textvariable=unit_var, values=sorted(self.unit_values), width=10)
        u_combo.pack(side="left", padx=(6, 0))
        u_combo.bind("<KeyRelease>", lambda _e: self._filter_combo_values(u_combo, sorted(self.unit_values)))
        u_combo.bind("<FocusOut>", lambda _e: self._enforce_combo_legal(u_combo, sorted(self.unit_values), "EUR"))
        r2 = ttk.Frame(w); r2.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(r2, text=f"{self.tr('from')}:").pack(side="left")
        ttk.Entry(r2, textvariable=lo_var, width=14).pack(side="left", padx=(6, 14))
        ttk.Label(r2, text=f"{self.tr('to')}:").pack(side="left")
        ttk.Entry(r2, textvariable=hi_var, width=14).pack(side="left", padx=(6, 0))
        btn = ttk.Frame(w); btn.pack(side="bottom", fill="x", padx=10, pady=10)
        def apply() -> None:
            unit = unit_var.get().strip().upper()
            if unit not in self.unit_values:
                return
            self.header_filters["unit"] = unit
            self.header_range["amount"] = (lo_var.get().strip(), hi_var.get().strip())
            w.destroy(); self.refresh_table()
        ttk.Button(btn, text=self.tr("apply"), command=apply).pack(side="right")
        ttk.Button(btn, text=self.tr("cancel"), command=w.destroy).pack(side="right", padx=(0, 6))

    def _filtered(self) -> list[tuple[int, Entry]]:
        out: list[tuple[int, Entry]] = []
        for i, e in enumerate(self.entries):
            ftype = self.header_filters.get("type", "")
            if ftype == "income" and e.amount >= 0:
                continue
            if ftype == "expense" and e.amount < 0:
                continue
            if self.header_filters.get("item") and self.header_filters["item"].lower() not in e.item.lower():
                continue
            if self.header_filters.get("unit") and self.header_filters["unit"].upper() != e.unit.upper():
                continue
            if "payment" in self.header_multi and e.payment not in self.header_multi["payment"]:
                continue
            if "merchant" in self.header_multi and e.merchant not in self.header_multi["merchant"]:
                continue
            if "category" in self.header_multi and e.category not in self.header_multi["category"]:
                continue
            aa = abs(e.amount)
            if "amount" in self.header_range:
                lo, hi = self.header_range["amount"]
                if lo:
                    try:
                        if aa < float(lo):
                            continue
                    except ValueError:
                        pass
                if hi:
                    try:
                        if aa > float(hi):
                            continue
                    except ValueError:
                        pass
            if "date" in self.header_range:
                d1, d2 = self.header_range["date"]
                if d1 and e.date < d1:
                    continue
                if d2 and e.date > d2:
                    continue
            out.append((i, e))
        return out

    def delete_selected(self) -> None:
        ids = list(self.tree.selection())
        if not ids:
            messagebox.showerror(self.tr("app_title"), self.tr("no_select"))
            return
        for i in sorted([int(x) for x in ids], reverse=True):
            if 0 <= i < len(self.entries):
                del self.entries[i]
        self.write_current_delr(); self._save_settings(); self.refresh_table()

    def _fmt(self, money: dict[str, float]) -> str:
        return " | ".join([f"{k}: {v:.2f}" for k, v in sorted(money.items())]) if money else "0.00"

    def refresh_table(self) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        try:
            rows = self._filtered()
        except Exception:
            rows = list(enumerate(self.entries))
        total: dict[str, float] = {}
        inc: dict[str, float] = {}
        exp: dict[str, float] = {}
        for idx, e in sorted(rows, key=lambda x: (x[1].date, x[1].item.lower(), x[0])):
            u = (e.unit or "EUR").upper()
            a = abs(e.amount)
            total[u] = total.get(u, 0.0) + a
            if e.amount < 0:
                inc[u] = inc.get(u, 0.0) + a
            else:
                exp[u] = exp.get(u, 0.0) + a
            self.tree.insert("", "end", iid=str(idx), values=(self.fmt_ui_date(e.date), self.tr("income") if e.amount < 0 else self.tr("expense"), e.item, f"{a:.2f}", u, e.payment, e.merchant, e.category))
        self.total_label.config(text=f"{self.tr('total')}: {self._fmt(total)}")
        self.income_label.config(text=f"{self.tr('income_total')}: {self._fmt(inc)}")
        self.expense_label.config(text=f"{self.tr('expense_total')}: {self._fmt(exp)}")


def runtime_app_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def main() -> None:
    root = tk.Tk()
    DelrLedgerApp(root, runtime_app_dir())
    root.mainloop()


if __name__ == "__main__":
    main()
