#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable


@dataclass
class Entry:
    date: date
    item: str
    amount: float
    payment: str
    merchant: str
    category: str


CSV_HEADERS = ["date", "item", "amount", "payment", "merchant", "category"]


def parse_date(raw: str) -> date:
    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Invalid date '{raw}'. Use YYYY-MM-DD."
        ) from exc


def month_key(d: date) -> str:
    return d.strftime("%Y-%m")


def month_csv_path(root: Path, ym: str) -> Path:
    return root / "data" / f"{ym}.csv"


def month_md_path(root: Path, ym: str) -> Path:
    year = ym.split("-")[0]
    return root / year / f"{ym}.md"


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def read_entries(csv_path: Path) -> list[Entry]:
    if not csv_path.exists():
        return []

    out: list[Entry] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            out.append(
                Entry(
                    date=parse_date(row["date"]),
                    item=row["item"],
                    amount=float(row["amount"]),
                    payment=row["payment"],
                    merchant=row["merchant"],
                    category=row["category"],
                )
            )
    return out


def write_entries(csv_path: Path, entries: Iterable[Entry]) -> None:
    ensure_parent(csv_path)
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writeheader()
        for e in sorted(entries, key=lambda x: (x.date, x.item.lower())):
            writer.writerow(
                {
                    "date": e.date.isoformat(),
                    "item": e.item,
                    "amount": f"{e.amount:.2f}",
                    "payment": e.payment,
                    "merchant": e.merchant,
                    "category": e.category,
                }
            )


def build_markdown(ym: str, entries: list[Entry]) -> str:
    lines: list[str] = [f"# Expenses {ym}", ""]
    grouped: dict[date, list[Entry]] = defaultdict(list)
    for e in sorted(entries, key=lambda x: (x.date, x.item.lower())):
        grouped[e.date].append(e)

    for day in sorted(grouped):
        lines.append(f"## {day.strftime('%m-%d')}")
        lines.append("")
        lines.append("| 商品名稱 | 金額 | 支付方式 | 消費地點 | 分類 |")
        lines.append("| --- | ---: | --- | --- | --- |")
        subtotal = 0.0
        for e in grouped[day]:
            subtotal += e.amount
            lines.append(
                f"| {e.item} | {e.amount:.2f} EUR | {e.payment} | {e.merchant} | {e.category} |"
            )
        lines.append("")
        lines.append(f"**小計：{subtotal:.2f} EUR**")
        lines.append("")

    total = sum(e.amount for e in entries)
    lines.append("---")
    lines.append("")
    lines.append(f"**月總計：{total:.2f} EUR**")
    lines.append("")
    return "\n".join(lines)


def cmd_add(args: argparse.Namespace) -> None:
    root = Path(args.root).resolve()
    d = args.date if args.date else datetime.now().date()
    ym = month_key(d)

    csv_path = month_csv_path(root, ym)
    entries = read_entries(csv_path)
    entries.append(
        Entry(
            date=d,
            item=args.item,
            amount=args.amount,
            payment=args.payment,
            merchant=args.merchant,
            category=args.category,
        )
    )
    write_entries(csv_path, entries)

    md_path = month_md_path(root, ym)
    ensure_parent(md_path)
    md_path.write_text(build_markdown(ym, entries), encoding="utf-8")
    print(f"Added 1 entry to {csv_path}")
    print(f"Updated markdown: {md_path}")


def cmd_summary(args: argparse.Namespace) -> None:
    root = Path(args.root).resolve()
    ym = args.month
    csv_path = month_csv_path(root, ym)
    entries = read_entries(csv_path)
    if not entries:
        print(f"No data found for {ym} ({csv_path})")
        return

    total = sum(e.amount for e in entries)
    by_category: dict[str, float] = defaultdict(float)
    by_day: dict[date, float] = defaultdict(float)
    for e in entries:
        by_category[e.category] += e.amount
        by_day[e.date] += e.amount

    print(f"Month: {ym}")
    print(f"Total: {total:.2f} EUR")
    print("")
    print("By category:")
    for cat, amount in sorted(by_category.items(), key=lambda x: x[1], reverse=True):
        print(f"  - {cat}: {amount:.2f} EUR")
    print("")
    print("By day:")
    for d, amount in sorted(by_day.items(), key=lambda x: x[0]):
        print(f"  - {d.isoformat()}: {amount:.2f} EUR")


def _parse_amount(raw: str) -> float:
    match = re.search(r"-?\d+(?:\.\d+)?", raw)
    if not match:
        raise ValueError(f"Cannot parse amount from '{raw}'")
    return float(match.group(0))


def cmd_import_md(args: argparse.Namespace) -> None:
    root = Path(args.root).resolve()
    md_path = Path(args.md_path).resolve()
    if not md_path.exists():
        raise FileNotFoundError(f"Markdown not found: {md_path}")

    ym = args.month
    year, _ = ym.split("-")
    text = md_path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()

    current_day: str | None = None
    imported: list[Entry] = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("## "):
            current_day = stripped.replace("## ", "").strip()
            continue

        if not stripped.startswith("|") or stripped.count("|") < 6:
            continue
        if "---" in stripped:
            continue

        cells = [c.strip() for c in stripped.strip("|").split("|")]
        if len(cells) != 5:
            continue

        if current_day is None:
            continue
        if not re.fullmatch(r"\d{2}-\d{2}", current_day):
            continue

        d = parse_date(f"{year}-{current_day}")
        item, amount_raw, payment, merchant, category = cells
        try:
            amount = _parse_amount(amount_raw)
        except ValueError:
            continue
        if not item or item.lower() in {"商品名稱", "item"}:
            continue

        imported.append(
            Entry(
                date=d,
                item=item,
                amount=amount,
                payment=payment,
                merchant=merchant,
                category=category,
            )
        )

    if not imported:
        print("No entries parsed from markdown.")
        return

    csv_path = month_csv_path(root, ym)
    write_entries(csv_path, imported)
    print(f"Imported {len(imported)} entries into {csv_path}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Simple CLI bookkeeping tool with Markdown export."
    )
    p.add_argument(
        "--root",
        default=".",
        help="Workspace root (default: current directory)",
    )

    sub = p.add_subparsers(dest="command", required=True)

    add = sub.add_parser("add", help="Add one expense entry")
    add.add_argument("--date", type=parse_date, help="Date in YYYY-MM-DD")
    add.add_argument("--item", required=True, help="Item name")
    add.add_argument("--amount", required=True, type=float, help="Amount in EUR")
    add.add_argument("--payment", required=True, help="Payment method")
    add.add_argument("--merchant", required=True, help="Merchant or location")
    add.add_argument("--category", required=True, help="Category")
    add.set_defaults(func=cmd_add)

    summary = sub.add_parser("summary", help="Show monthly summary")
    summary.add_argument("--month", required=True, help="Month in YYYY-MM")
    summary.set_defaults(func=cmd_summary)

    import_md = sub.add_parser("import-md", help="Import monthly markdown into CSV")
    import_md.add_argument("--month", required=True, help="Month in YYYY-MM")
    import_md.add_argument("--md-path", required=True, help="Path to the monthly markdown")
    import_md.set_defaults(func=cmd_import_md)
    return p


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
