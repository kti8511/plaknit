import csv
import json
import re
from pathlib import Path


DATA_FILE = Path("data.json")
MANUAL_FILE = Path("manual_sales_updates.json")

CSV_UPDATES = {
    "2026-04-24": Path(r"C:\Users\user\Desktop\0424.csv"),
    "2026-04-25": Path(r"C:\Users\user\Desktop\0425.csv"),
    "2026-04-26": Path(r"C:\Users\user\Desktop\0426.csv"),
}

DAILY_TOTALS = {
    "2026-04-24": {"payment": 462400, "orders": 8, "qty": 13},
    "2026-04-25": {"payment": 360680, "orders": 4, "qty": 9},
    "2026-04-26": {"payment": 344200, "orders": 5, "qty": 10},
}


def normalize_name(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r"^\*", "", name).strip()
    name = name.replace("[채코제에디션]", "").strip()
    return re.sub(r"\s+", " ", name)


def parse_option(option_text: str):
    color = ""
    size = ""
    for line in (option_text or "").splitlines():
        if ":" not in line:
            continue
        key, value = [part.strip() for part in line.split(":", 1)]
        if key == "색상":
            color = value
        elif key == "사이즈":
            size = value
    return color, size


def load_csv_rows(path: Path):
    rows = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        for r in reader:
            if not r:
                continue
            rows.append(
                {
                    "name": normalize_name(r[3]),
                    "raw_name": r[3],
                    "option": r[4],
                    "gross": int(float(r[10] or 0)),
                    "qty": int(float(r[9] or 0)),
                    "price": int(float(r[5] or 0)),
                }
            )
    return rows


def allocate_scaled_payments(items, target_payment):
    gross_total = sum(item["gross"] for item in items)
    if gross_total == 0:
        for item in items:
            item["payment"] = 0
        return
    scaled = []
    running = 0
    for item in items:
        raw = item["gross"] * target_payment / gross_total
        pay = int(round(raw))
        scaled.append(pay)
        running += pay
    diff = target_payment - running
    if scaled:
        scaled[-1] += diff
    for item, pay in zip(items, scaled):
        item["payment"] = pay


def build_match_index(rows):
    index = {}
    for i, row in enumerate(rows):
        if row.get("retailer") != "자사몰":
            continue
        key = (normalize_name(row.get("name", "")), row.get("color", ""), row.get("size", ""))
        index.setdefault(key, []).append(i)
    return index


def find_match(rows, index, item_name, color, size):
    keys = [
        (item_name, color, size),
        (item_name, "", size),
        (item_name, color, ""),
        (item_name, "", ""),
    ]
    for key in keys:
        hits = index.get(key) or []
        if hits:
            return hits[0]

    for i, row in enumerate(rows):
        if row.get("retailer") != "자사몰":
            continue
        if normalize_name(row.get("name", "")) != item_name:
            continue
        row_color = row.get("color", "")
        row_size = row.get("size", "")
        if color and row_color and color != row_color:
            continue
        if size and row_size and size != row_size:
            continue
        return i
    return None


def update_manual_totals(manual_updates):
    by_date = {entry["date"]: entry for entry in manual_updates}
    for day, totals in DAILY_TOTALS.items():
        entry = by_date.setdefault(day, {"date": day, "retailers": []})
        retailer_entry = None
        for retailer in entry["retailers"]:
            if retailer.get("retailer") == "자사몰":
                retailer_entry = retailer
                break
        if retailer_entry is None:
            retailer_entry = {"retailer": "자사몰", "payment": 0, "items": []}
            entry["retailers"].append(retailer_entry)
        retailer_entry["payment"] = 0
        retailer_entry["orders"] = totals["orders"]
        retailer_entry["qty"] = totals["qty"]
        retailer_entry["items"] = []
    return sorted(manual_updates, key=lambda x: x["date"])


def main():
    rows = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    manual_updates = json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    index = build_match_index(rows)

    for day, path in CSV_UPDATES.items():
        items = load_csv_rows(path)
        allocate_scaled_payments(items, DAILY_TOTALS[day]["payment"])
        for item in items:
            color, size = parse_option(item["option"])
            item_name = item["name"]
            match_idx = find_match(rows, index, item_name, color, size)

            if match_idx is None:
                new_row = {
                    "retailer": "자사몰",
                    "mall_no": "",
                    "name": item_name,
                    "color": color,
                    "size": size,
                    "qty": item["qty"],
                    "gross": item["gross"],
                    "payment": item["payment"],
                    "avg_unit": int(round(item["payment"] / item["qty"])) if item["qty"] else 0,
                    "orders": max(item["qty"], 0),
                    "source_type": "단품",
                    "daily": [
                        {
                            "date": day,
                            "qty": item["qty"],
                            "gross": item["gross"],
                            "payment": item["payment"],
                            "orders": max(item["qty"], 0),
                        }
                    ],
                    "match_status": "매칭됨",
                    "match_sku": "",
                    "standard_name": f"{item_name}_{color}-{size}" if color else f"{item_name} {size}".strip(),
                    "received_qty": 0,
                    "stock_qty": 0,
                }
                rows.append(new_row)
                index.setdefault((item_name, color, size), []).append(len(rows) - 1)
                continue

            row = rows[match_idx]
            row["qty"] = int(row.get("qty", 0)) + item["qty"]
            row["gross"] = int(row.get("gross", 0)) + item["gross"]
            row["payment"] = int(row.get("payment", 0)) + item["payment"]
            row["orders"] = int(row.get("orders", 0)) + max(item["qty"], 0)
            row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0

            existing_daily = None
            for daily in row.get("daily", []):
                if daily.get("date") == day:
                    existing_daily = daily
                    break
            if existing_daily is None:
                row.setdefault("daily", []).append(
                    {
                        "date": day,
                        "qty": item["qty"],
                        "gross": item["gross"],
                        "payment": item["payment"],
                        "orders": max(item["qty"], 0),
                    }
                )
            else:
                existing_daily["qty"] = int(existing_daily.get("qty", 0)) + item["qty"]
                existing_daily["gross"] = int(existing_daily.get("gross", 0)) + item["gross"]
                existing_daily["payment"] = int(existing_daily.get("payment", 0)) + item["payment"]
                existing_daily["orders"] = int(existing_daily.get("orders", 0)) + max(item["qty"], 0)

    manual_updates = update_manual_totals(manual_updates)

    DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    MANUAL_FILE.write_text(json.dumps(manual_updates, ensure_ascii=False, indent=2), encoding="utf-8")
    print("updated data.json and manual_sales_updates.json")


if __name__ == "__main__":
    main()
