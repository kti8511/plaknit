import argparse
import json
from pathlib import Path


DATA_FILE = Path("data.json")
MANUAL_FILE = Path("manual_sales_updates.json")
RETAILER = "29cm"


DAILY_ITEMS = [
    {
        "name": "ICE LITE 시그니처 조거",
        "color": "블랙",
        "size": "L",
        "qty": 2,
        "payment": 117300,
    }
]


def as_int(value):
    if value in (None, ""):
        return 0
    return int(round(float(str(value).replace(",", ""))))


def standard_name(item):
    return f"{item['name']}_{item['color']}-{item['size']}"


def remove_existing_day(rows, day):
    for row in rows:
        if row.get("retailer") != RETAILER:
            continue
        kept = []
        removed_qty = removed_gross = removed_payment = removed_orders = 0
        for daily in row.get("daily", []):
            if daily.get("date") == day:
                removed_qty += as_int(daily.get("qty"))
                removed_gross += as_int(daily.get("gross"))
                removed_payment += as_int(daily.get("payment"))
                removed_orders += as_int(daily.get("orders"))
            else:
                kept.append(daily)
        if len(kept) != len(row.get("daily", [])):
            row["daily"] = kept
            row["qty"] = as_int(row.get("qty")) - removed_qty
            row["gross"] = as_int(row.get("gross")) - removed_gross
            row["payment"] = as_int(row.get("payment")) - removed_payment
            row["orders"] = as_int(row.get("orders")) - removed_orders
            row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0


def build_index(rows):
    index = {}
    for idx, row in enumerate(rows):
        if row.get("retailer") != RETAILER:
            continue
        key = (row.get("name", ""), row.get("color", ""), row.get("size", ""))
        index.setdefault(key, idx)
    return index


def append_daily(row, day, item):
    row.setdefault("daily", []).append(
        {
            "date": day,
            "qty": item["qty"],
            "gross": item["payment"],
            "payment": item["payment"],
            "orders": 1,
        }
    )
    row["qty"] = as_int(row.get("qty")) + item["qty"]
    row["gross"] = as_int(row.get("gross")) + item["payment"]
    row["payment"] = as_int(row.get("payment")) + item["payment"]
    row["orders"] = as_int(row.get("orders")) + 1
    row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0


def update_manual(day, items):
    manual = json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    by_date = {entry["date"]: entry for entry in manual}
    entry = by_date.setdefault(day, {"date": day, "retailers": []})
    retailers = entry.setdefault("retailers", [])
    target = next((r for r in retailers if r.get("retailer") == RETAILER), None)
    if target is None:
        target = {"retailer": RETAILER, "items": []}
        retailers.append(target)
    target["payment"] = sum(item["payment"] for item in items)
    target["qty"] = sum(item["qty"] for item in items)
    target["orders"] = len(items)
    target["items"] = [{"name": standard_name(item), "qty": item["qty"], "payment": item["payment"]} for item in items]
    return [by_date[date] for date in sorted(by_date)]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default="2026-05-07")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    rows = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    remove_existing_day(rows, args.date)
    index = build_index(rows)
    created = []

    for item in DAILY_ITEMS:
        key = (item["name"], item["color"], item["size"])
        idx = index.get(key)
        if idx is None:
            row = {
                "retailer": RETAILER,
                "mall_no": "",
                "name": item["name"],
                "color": item["color"],
                "size": item["size"],
                "qty": 0,
                "gross": 0,
                "payment": 0,
                "avg_unit": 0,
                "orders": 0,
                "source_type": "상품",
                "daily": [],
                "match_status": "매칭완료",
                "match_sku": "",
                "standard_name": standard_name(item),
                "received_qty": 0,
                "stock_qty": 0,
            }
            rows.append(row)
            idx = len(rows) - 1
            created.append(row["standard_name"])
        append_daily(rows[idx], args.date, item)

    manual = update_manual(args.date, DAILY_ITEMS)
    summary = {
        "date": args.date,
        "payment": sum(item["payment"] for item in DAILY_ITEMS),
        "qty": sum(item["qty"] for item in DAILY_ITEMS),
        "created_rows": created,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if args.dry_run:
        return
    DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    MANUAL_FILE.write_text(json.dumps(manual, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
