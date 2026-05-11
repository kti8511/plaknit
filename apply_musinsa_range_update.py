import argparse
import json
import re
from pathlib import Path


DATA_FILE = Path("data.json")
MANUAL_FILE = Path("manual_sales_updates.json")
RETAILER = "무신사"


UPDATES = {
    "2026-05-08": [
        {"name": "ICE LITE 벌룬 조거", "qty": 2, "payment": 77450},
        {"name": "[2PACK] TECH SWEAT 7인치 카고 반바지", "qty": 1, "payment": 67280},
        {"name": "TECH SWEAT 피그먼트 스웻쇼츠 레드", "qty": 1, "payment": 58780},
        {"name": "[2PACK] 퀵드라이 세미오버핏 아치로고 티셔츠", "qty": 1, "payment": 46960},
        {"name": "ICE LITE 와이드 카고팬츠", "qty": 1, "payment": 44100},
        {"name": "ICE LITE 세미오버핏 맨투맨", "qty": 1, "payment": 39900},
        {"name": "퀵드라이 피그먼트 세미오버핏 반팔티", "qty": 1, "payment": 36620},
        {"name": "TECH SWEAT 벌룬 쇼츠", "qty": 1, "payment": 34380},
        {"name": "퀵드라이 퍼포먼스 숏슬리브 블랙", "qty": -1, "payment": -51380},
        {"name": "[헤비쮸리] TECH SWEAT 세미오버핏 후드", "qty": -1, "payment": -54900},
        {"name": "[2PACK] 초냉감 트래블러 티셔츠", "qty": -1, "payment": -58710},
    ],
    "2026-05-09": [
        {"name": "ICE LITE 초냉감 오버핏 티셔츠", "qty": 2, "payment": 65050},
        {"name": "ICE LITE 와이드 카고팬츠", "qty": 1, "payment": 58030},
        {"name": "ICE LITE 카고 조거", "qty": 1, "payment": 48780},
        {"name": "ICE LITE 벌크업 쇼츠", "qty": 1, "payment": 42550},
        {"name": "ICE LITE 벌룬 조거", "qty": 1, "payment": 38310},
        {"name": "TECH SWEAT 릴렉스드 반바지", "qty": 1, "payment": 33430},
        {"name": "ICE LITE 초냉감 러닝 나시 (싱글렛)", "qty": 1, "payment": 29380},
    ],
    "2026-05-10": [
        {"name": "퀵드라이 세미 오버핏 티셔츠", "qty": 5, "payment": 158590},
        {"name": "ICE LITE 루즈핏카고쇼츠", "qty": 2, "payment": 98340},
        {"name": "[2PACK] 퀵드라이 세미오버핏 아치로고 티셔츠", "qty": 2, "payment": 87350},
        {"name": "ICE LITE 와이드 카고팬츠", "qty": 1, "payment": 54290},
        {"name": "ICE LITE 벌룬 조거", "qty": 1, "payment": 39900},
        {"name": "퀵드라이 피그먼트 세미오버핏 반팔티", "qty": 1, "payment": 36070},
        {"name": "퀵드라이 백레터링 세미오버핏 반팔티", "qty": 1, "payment": 35890},
        {"name": "TECH SWEAT 피그먼트 스웻쇼츠 레드", "qty": -1, "payment": -58780},
    ],
}


NAME_ALIASES = {
    "퀵드라이 세미 오버핏 티셔츠": "퀵드라이 세미오버핏 티셔츠",
    "[2PACK] 초냉감 트래블러 티셔츠": "[2PACK] 채코제에디션 초냉감 트래블러티셔츠",
    "초냉감 트래블러 티셔츠": "채코제에디션 초냉감 트래블러티셔츠",
    "[2PACK] TECH SWEAT 7인치 카고 반바지": '[2PACK] TECH SWEAT 7" 카고 반바지',
}


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def normalize_name(name):
    text = clean_text(name)
    text = text.replace("퀵 드라이", "퀵드라이")
    text = text.replace("세미 오버핏", "세미오버핏")
    text = text.replace("7인치", '7"')
    text = NAME_ALIASES.get(text, text)
    return text


def as_int(value):
    if value in (None, ""):
        return 0
    return int(round(float(str(value).replace(",", ""))))


def row_key(row):
    return normalize_name(row.get("name"))


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
        index.setdefault(row_key(row), idx)
    return index


def append_daily(row, day, item):
    gross = item["payment"]
    order_count = 1 if item["qty"] > 0 else -1
    row.setdefault("daily", []).append(
        {
            "date": day,
            "qty": item["qty"],
            "gross": gross,
            "payment": item["payment"],
            "orders": order_count,
        }
    )
    row["qty"] = as_int(row.get("qty")) + item["qty"]
    row["gross"] = as_int(row.get("gross")) + gross
    row["payment"] = as_int(row.get("payment")) + item["payment"]
    row["orders"] = as_int(row.get("orders")) + order_count
    row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0


def update_manual(manual, day, items):
    by_date = {entry["date"]: entry for entry in manual}
    entry = by_date.setdefault(day, {"date": day, "retailers": []})
    retailers = entry.setdefault("retailers", [])
    target = next((retailer for retailer in retailers if retailer.get("retailer") == RETAILER), None)
    if target is None:
        target = {"retailer": RETAILER, "items": []}
        retailers.append(target)
    target["payment"] = sum(item["payment"] for item in items)
    target["qty"] = sum(item["qty"] for item in items)
    target["orders"] = sum(1 if item["qty"] > 0 else -1 for item in items)
    target["items"] = [
        {"name": item["name"], "qty": item["qty"], "payment": item["payment"]}
        for item in items
    ]
    return [by_date[date] for date in sorted(by_date)]


def apply_updates(rows, day, items):
    remove_existing_day(rows, day)
    index = build_index(rows)
    created = []
    normalized_items = [{**item, "name": normalize_name(item["name"])} for item in items]
    for item in normalized_items:
        idx = index.get(item["name"])
        if idx is None:
            new_row = {
                "retailer": RETAILER,
                "mall_no": "",
                "name": item["name"],
                "color": "",
                "size": "",
                "qty": 0,
                "gross": 0,
                "payment": 0,
                "avg_unit": 0,
                "orders": 0,
                "source_type": "상품",
                "daily": [],
                "match_status": "매칭완료",
                "match_sku": "",
                "standard_name": item["name"],
                "received_qty": 0,
                "stock_qty": 0,
            }
            rows.append(new_row)
            idx = len(rows) - 1
            index[item["name"]] = idx
            created.append(item["name"])
        row = rows[idx]
        row["name"] = item["name"]
        row["match_status"] = "매칭완료"
        if not row.get("standard_name"):
            row["standard_name"] = item["name"]
        append_daily(row, day, item)
    return normalized_items, created


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    rows = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    manual = json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    all_created = {}
    summaries = {}

    for day, items in UPDATES.items():
        normalized_items, created = apply_updates(rows, day, items)
        manual = update_manual(manual, day, normalized_items)
        all_created[day] = created
        summaries[day] = {
            "line_items": len(items),
            "qty": sum(item["qty"] for item in normalized_items),
            "payment": sum(item["payment"] for item in normalized_items),
            "created_rows": created,
        }

    print(json.dumps(summaries, ensure_ascii=False, indent=2))
    if args.dry_run:
        return
    DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    MANUAL_FILE.write_text(json.dumps(manual, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
