import argparse
import json
from pathlib import Path


DATA_FILE = Path("data.json")
MANUAL_FILE = Path("manual_sales_updates.json")


def as_int(value):
    return int(round(float(str(value).replace(",", ""))))


def append_or_merge_daily(row, day, qty, gross, payment, orders):
    for daily in row.setdefault("daily", []):
        if daily.get("date") == day:
            daily["qty"] = as_int(daily.get("qty", 0)) + qty
            daily["gross"] = as_int(daily.get("gross", 0)) + gross
            daily["payment"] = as_int(daily.get("payment", 0)) + payment
            daily["orders"] = as_int(daily.get("orders", 0)) + orders
            return
    row["daily"].append(
        {"date": day, "qty": qty, "gross": gross, "payment": payment, "orders": orders}
    )


def update_manual(manual, day, retailer, item_name, qty, payment):
    by_date = {entry["date"]: entry for entry in manual}
    entry = by_date.setdefault(day, {"date": day, "retailers": []})
    retailers = entry.setdefault("retailers", [])
    target = next((item for item in retailers if item.get("retailer") == retailer), None)
    if target is None:
        target = {"retailer": retailer, "items": [], "payment": 0, "qty": 0, "orders": 0}
        retailers.append(target)

    items = target.setdefault("items", [])
    existing = next((item for item in items if item.get("name") == item_name), None)
    if existing is None:
        items.append({"name": item_name, "qty": qty, "payment": payment})
    else:
        existing["qty"] = as_int(existing.get("qty", 0)) + qty
        existing["payment"] = as_int(existing.get("payment", 0)) + payment

    target["payment"] = as_int(target.get("payment", 0)) + payment
    target["qty"] = as_int(target.get("qty", 0)) + qty
    target["orders"] = as_int(target.get("orders", 0)) + max(qty, 0)
    return [by_date[date] for date in sorted(by_date)]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", required=True)
    parser.add_argument("--retailer", required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--color", required=True)
    parser.add_argument("--size", required=True)
    parser.add_argument("--standard-name", required=True)
    parser.add_argument("--qty", required=True, type=int)
    parser.add_argument("--payment", required=True, type=int)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    rows = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    matches = [
        row
        for row in rows
        if row.get("retailer") == args.retailer
        and row.get("standard_name") == args.standard_name
    ]
    if len(matches) != 1:
        raise SystemExit(f"expected 1 match, found {len(matches)} for {args.standard_name}")

    row = matches[0]
    row["name"] = args.name
    row["color"] = args.color
    row["size"] = args.size
    row["match_status"] = "매칭완료"
    row["qty"] = as_int(row.get("qty", 0)) + args.qty
    row["gross"] = as_int(row.get("gross", 0)) + args.payment
    row["payment"] = as_int(row.get("payment", 0)) + args.payment
    row["orders"] = as_int(row.get("orders", 0)) + max(args.qty, 0)
    row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0
    append_or_merge_daily(row, args.date, args.qty, args.payment, args.payment, max(args.qty, 0))

    manual = json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    manual = update_manual(
        manual,
        args.date,
        args.retailer,
        args.standard_name,
        args.qty,
        args.payment,
    )

    summary = {
        "date": args.date,
        "retailer": args.retailer,
        "standard_name": args.standard_name,
        "qty": args.qty,
        "payment": args.payment,
        "row_qty": row["qty"],
        "row_payment": row["payment"],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if not args.dry_run:
        DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        MANUAL_FILE.write_text(json.dumps(manual, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
