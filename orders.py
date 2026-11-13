#!/usr/bin/env python3
"""Store orders in an Excel workbook.

Usage:
    python orders.py init                      # create orders.xlsx with headers
    python orders.py add "Alice" "Widget" 3 9.99 [--status Pending]
    python orders.py list

Requires: pip install openpyxl
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

DEFAULT_FILE = "orders.xlsx"
SHEET_NAME = "Orders"
HEADERS = ["Order ID", "Order Date", "Customer", "Product", "Quantity", "Unit Price", "Total", "Status"]
COL_WIDTHS = [12, 20, 24, 24, 10, 12, 12, 12]
STATUSES = ["Pending", "Shipped", "Delivered", "Cancelled"]

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def create_workbook(path: Path) -> None:
    """Create a new workbook with a formatted Orders sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    ws.append(HEADERS)
    for col, width in enumerate(COL_WIDTHS, start=1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H1"

    # Dropdown for Status column (rows 2..10000)
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("H2:H10000")

    wb.save(path)
    print(f"Created {path}")


def next_order_id(ws) -> str:
    last = ws.max_row
    if last < 2:
        return "ORD-0001"
    prev = str(ws.cell(row=last, column=1).value or "ORD-0000")
    try:
        num = int(prev.split("-")[1]) + 1
    except (IndexError, ValueError):
        num = last
    return f"ORD-{num:04d}"


def add_order(path: Path, customer: str, product: str, qty: int, price: float, status: str) -> None:
    if not path.exists():
        create_workbook(path)
    wb = load_workbook(path)
    ws = wb[SHEET_NAME]

    order_id = next_order_id(ws)
    row = ws.max_row + 1
    values = [order_id, datetime.now(), customer, product, qty, price, f"=E{row}*F{row}", status]
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.border = BORDER

    ws.cell(row=row, column=2).number_format = "yyyy-mm-dd hh:mm"
    ws.cell(row=row, column=6).number_format = "$#,##0.00"
    ws.cell(row=row, column=7).number_format = "$#,##0.00"
    ws.auto_filter.ref = f"A1:H{row}"

    wb.save(path)
    print(f"Added {order_id}: {qty} x {product} for {customer} @ {price:.2f} ({status})")


def list_orders(path: Path) -> None:
    if not path.exists():
        sys.exit(f"{path} not found. Run 'init' or 'add' first.")
    ws = load_workbook(path)[SHEET_NAME]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows:
        print("No orders yet.")
        return
    print(f"{'ID':<10} {'Date':<17} {'Customer':<18} {'Product':<18} {'Qty':>4} {'Price':>9} {'Total':>10}  Status")
    grand = 0.0
    for oid, date, cust, prod, qty, price, _total, status in rows:
        total = (qty or 0) * (price or 0)
        grand += total
        date_s = date.strftime("%Y-%m-%d %H:%M") if isinstance(date, datetime) else str(date)
        print(f"{oid:<10} {date_s:<17} {cust:<18} {prod:<18} {qty:>4} {price:>9.2f} {total:>10.2f}  {status}")
    print(f"\n{len(rows)} order(s), grand total: {grand:.2f}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Manage orders in an Excel sheet.")
    parser.add_argument("--file", default=DEFAULT_FILE, help=f"Excel file (default: {DEFAULT_FILE})")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_init = sub.add_parser("init", help="Create a new orders workbook")
    p_init.add_argument("--force", action="store_true", help="Overwrite if it exists")

    p_add = sub.add_parser("add", help="Add an order")
    p_add.add_argument("customer")
    p_add.add_argument("product")
    p_add.add_argument("quantity", type=int)
    p_add.add_argument("price", type=float)
    p_add.add_argument("--status", default="Pending", choices=STATUSES)

    sub.add_parser("list", help="List all orders")

    args = parser.parse_args()
    path = Path(args.file)

    if args.cmd == "init":
        if path.exists() and not args.force:
            sys.exit(f"{path} already exists. Use --force to overwrite.")
        create_workbook(path)
    elif args.cmd == "add":
        if args.quantity <= 0 or args.price < 0:
            sys.exit("Quantity must be > 0 and price must be >= 0.")
        add_order(path, args.customer, args.product, args.quantity, args.price, args.status)
    elif args.cmd == "list":
        list_orders(path)


if __name__ == "__main__":
    main()
