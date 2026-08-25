import argparse
import json
import re
from pathlib import Path


DATA_FILE = Path("data.json")
MANUAL_FILE = Path("manual_sales_updates.json")
RETAILER = "무신사"


DAILY_ITEMS = [
    {"name": "ICE LITE 오버핏 맨투맨", "qty": 3, "payment": 106310},
    {"name": "ICE LITE 루즈핏 카고쇼츠", "qty": 2, "payment": 96640},
    {"name": "KINTERRA 퀵드라이 스포츠 나시", "qty": 2, "payment": 64820},
    {"name": "퀵드라이 세미오버핏 백로고 티셔츠", "qty": 2, "payment": 57150},
    {"name": "퀵드라이 메쉬 러닝싱글렛 화이트", "qty": 1, "payment": 54930},
    {"name": "ICE LITE 카고 조거", "qty": 1, "payment": 54190},
    {"name": "퀵드라이 셔링 티셔츠 블랙", "qty": 1, "payment": 51260},
    {"name": "퀵드라이 셔링 티셔츠 세이지", "qty": 1, "payment": 51260},
    {"name": "[2PACK] 퀵드라이 세미오버핏 아치로고 티셔츠", "qty": 1, "payment": 47410},
    {"name": "[2PACK] 퀵드라이 세미오버핏 백로고 티셔츠", "qty": 1, "payment": 44910},
    {"name": "ICE LITE 벌룬 조거", "qty": 1, "payment": 38710},
    {"name": "ICE LITE 우먼즈 데일리쇼츠 아이보리", "qty": 1, "payment": 38520},
    {"name": "ICE LITE 데일리 쇼츠", "qty": 1, "payment": 36770},
    {"name": "ICE LITE 컴포트 쇼츠", "qty": 1, "payment": 36010},
    {"name": "퀵드라이 마운틴 세미오버핏 반팔티", "qty": 1, "payment": 33980},
    {"name": "MESH-ON 파워리프팅 롱티(롱 슬리브)", "qty": 1, "payment": 33390},
    {"name": "퀵드라이 세미 오버핏 티셔츠", "qty": 1, "payment": 31920},
    {"name": "스포츠 쿠션업 심볼 크루삭스", "qty": 2, "payment": 13530},
    {"name": "퀵드라이 피그먼트 오버핏 반팔티", "qty": -1, "payment": -36700},
]


NAME_ALIASES = {
    "[헤비쮸리]TECH SWEAT 피그먼트 빈티지 조거팬츠 (베이직)": "TECH SWEAT 피그먼트 소프트 빈티지 조거팬츠(베이직)",
    "[헤비쮸리] TECH SWEAT 피그먼트 빈티지 조거팬츠 (베이직)": "TECH SWEAT 피그먼트 소프트 빈티지 조거팬츠(베이직)",
    "KINTERRA 퀵드라이 스포츠 오버핏 반팔티": "KINTERRA 퀵드라이 오버핏 반팔티",
    "KINTERRA 퀵드라이 스포츠 머슬핏 반팔티": "KINTERRA 퀵드라이 스포츠 머슬핏 반팔티",
    "ICE LITE 초냉감 PLNT 세미오버핏 반팔티": "ICE LITE 초냉감 PLNT 세미오버핏 티셔츠",
    "초냉감 트래블러 티셔츠": "채코제에디션 초냉감 트래블러티셔츠",
}


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def normalize_name(name):
    text = clean_text(name)
    text = text.replace("퀵 드라이", "퀵드라이")
    text = text.replace("세미 오버핏", "세미오버핏")
    return NAME_ALIASES.get(text, text)


def as_int(value):
    if value in (None, ""):
        return 0
    return int(round(float(str(value).replace(",", ""))))


def row_name(row):
    return normalize_name(row.get("name"))


def standard_name(name, row=None):
    if row and row.get("standard_name"):
        return row["standard_name"]
    return name


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


def build_name_index(rows):
    index = {}
    for idx, row in enumerate(rows):
        if row.get("retailer") != RETAILER:
            continue
        index.setdefault(row_name(row), idx)
    return index


def append_daily(row, day, item):
    gross = abs(item["payment"])
    order_delta = 1 if item["qty"] > 0 else -1 if item["qty"] < 0 else 0
    row.setdefault("daily", []).append(
        {
            "date": day,
            "qty": item["qty"],
            "gross": gross if item["payment"] >= 0 else -gross,
            "payment": item["payment"],
            "orders": order_delta,
        }
    )
    row["qty"] = as_int(row.get("qty")) + item["qty"]
    row["gross"] = as_int(row.get("gross")) + (gross if item["payment"] >= 0 else -gross)
    row["payment"] = as_int(row.get("payment")) + item["payment"]
    row["orders"] = as_int(row.get("orders")) + order_delta
    row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0


def update_manual(day, items):
    manual = json.loads(MANUAL_FILE.read_text(encoding="utf-8-sig"))
    by_date = {entry["date"]: entry for entry in manual}
    entry = by_date.setdefault(day, {"date": day, "retailers": []})
    retailers = entry.setdefault("retailers", [])
    target = next((r for r in retailers if r.get("retailer") == RETAILER), None)
    if target is None:
        target = {"retailer": RETAILER, "items": []}
        retailers.append(target)
    target["payment"] = sum(item["payment"] for item in items)
    target["qty"] = sum(item["qty"] for item in items)
    target["orders"] = sum(1 if item["qty"] > 0 else -1 if item["qty"] < 0 else 0 for item in items)
    target["items"] = [{"name": item["name"], "qty": item["qty"], "payment": item["payment"]} for item in items]
    return [by_date[date] for date in sorted(by_date)]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default="2026-05-07")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    rows = json.loads(DATA_FILE.read_text(encoding="utf-8-sig"))
    items = [{**item, "name": normalize_name(item["name"])} for item in DAILY_ITEMS]
    remove_existing_day(rows, args.date)
    index = build_name_index(rows)
    created = []

    for item in items:
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
        row["standard_name"] = standard_name(item["name"], row)
        append_daily(row, args.date, item)

    manual = update_manual(args.date, items)
    summary = {
        "date": args.date,
        "line_items": len(items),
        "qty": sum(item["qty"] for item in items),
        "payment": sum(item["payment"] for item in items),
        "created_rows": created,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if args.dry_run:
        return
    DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    MANUAL_FILE.write_text(json.dumps(manual, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
