#!/usr/bin/env python3
from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from expenses import Entry, build_markdown, month_csv_path, month_md_path, read_entries, write_entries


class DelrLedgerApp:
    def __init__(self, root: tk.Tk, workspace: Path) -> None:
        self.root = root
        self.workspace = workspace
        self.root.title("DELR Ledger")
        self.root.geometry("760x560")
        self.root.minsize(700, 520)

        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.item_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.payment_var = tk.StringVar(value="Cash")
        self.merchant_var = tk.StringVar()
        self.category_var = tk.StringVar(value="Food")
        self.month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))

        self._build_ui()
        self.refresh_summary()

    def _build_ui(self) -> None:
        form = ttk.LabelFrame(self.root, text="Add Expense")
        form.pack(fill="x", padx=12, pady=12)

        labels = [
            ("Date (YYYY-MM-DD)", self.date_var),
            ("Item", self.item_var),
            ("Amount (EUR)", self.amount_var),
            ("Payment", self.payment_var),
            ("Merchant", self.merchant_var),
            ("Category", self.category_var),
        ]
        for i, (text, var) in enumerate(labels):
            ttk.Label(form, text=text).grid(row=i, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(form, textvariable=var, width=42).grid(row=i, column=1, sticky="ew", padx=8, pady=6)

        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(form)
        actions.grid(row=len(labels), column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        ttk.Button(actions, text="Add Entry", command=self.add_entry).pack(side="left")
        ttk.Button(actions, text="Clear", command=self.clear_form).pack(side="left", padx=(8, 0))

        summary = ttk.LabelFrame(self.root, text="Monthly Summary")
        summary.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        top = ttk.Frame(summary)
        top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="Month (YYYY-MM)").pack(side="left")
        ttk.Entry(top, textvariable=self.month_var, width=10).pack(side="left", padx=8)
        ttk.Button(top, text="Refresh", command=self.refresh_summary).pack(side="left")

        self.summary_text = tk.Text(summary, height=20, wrap="word")
        self.summary_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def clear_form(self) -> None:
        self.item_var.set("")
        self.amount_var.set("")
        self.merchant_var.set("")

    def add_entry(self) -> None:
        try:
            d = datetime.strptime(self.date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Invalid Date", "Date must be YYYY-MM-DD.")
            return

        item = self.item_var.get().strip()
        if not item:
            messagebox.showerror("Invalid Item", "Item cannot be empty.")
            return

        try:
            amount = float(self.amount_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Amount", "Amount must be a number.")
            return

        payment = self.payment_var.get().strip() or "Cash"
        merchant = self.merchant_var.get().strip() or "-"
        category = self.category_var.get().strip() or "Other"

        ym = d.strftime("%Y-%m")
        csv_path = month_csv_path(self.workspace, ym)
        entries = read_entries(csv_path)
        entries.append(
            Entry(
                date=d,
                item=item,
                amount=amount,
                payment=payment,
                merchant=merchant,
                category=category,
            )
        )
        write_entries(csv_path, entries)

        md_path = month_md_path(self.workspace, ym)
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text(build_markdown(ym, entries), encoding="utf-8")

        self.month_var.set(ym)
        self.refresh_summary()
        messagebox.showinfo("Saved", f"Saved entry to {csv_path.name}")

    def refresh_summary(self) -> None:
        ym = self.month_var.get().strip()
        csv_path = month_csv_path(self.workspace, ym)
        entries = read_entries(csv_path)
        out: list[str] = []
        out.append(f"Workspace: {self.workspace}")
        out.append(f"Month: {ym}")
        out.append(f"CSV: {csv_path}")
        out.append("")
        if not entries:
            out.append("No data.")
        else:
            total = sum(e.amount for e in entries)
            by_category: dict[str, float] = defaultdict(float)
            by_day: dict[str, float] = defaultdict(float)
            for e in entries:
                by_category[e.category] += e.amount
                by_day[e.date.isoformat()] += e.amount

            out.append(f"Total: {total:.2f} EUR")
            out.append("")
            out.append("By category:")
            for cat, amount in sorted(by_category.items(), key=lambda x: x[1], reverse=True):
                out.append(f"- {cat}: {amount:.2f} EUR")
            out.append("")
            out.append("By day:")
            for day, amount in sorted(by_day.items()):
                out.append(f"- {day}: {amount:.2f} EUR")

        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", "\n".join(out))


def main() -> None:
    workspace = Path(__file__).resolve().parent
    root = tk.Tk()
    DelrLedgerApp(root, workspace)
    root.mainloop()


if __name__ == "__main__":
    main()
