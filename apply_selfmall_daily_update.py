import argparse
import csv
import datetime as dt
import json
import re
from pathlib import Path


DATA_FILE = Path("data.json")
MANUAL_FILE = Path("manual_sales_updates.json")
RETAILER = "자사몰"


def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def norm_size(value):
    text = clean_text(value).upper()
    return {"XXL": "2XL", "XXXL": "3XL", "XXXXL": "4XL"}.get(text, text)


def norm_color(value):
    text = clean_text(value)
    aliases = {
        "BLACK": "블랙",
        "WHITE": "화이트",
        "CHARCOAL": "차콜",
        "DEEP CHARCOAL": "차콜",
        "LIGHT GRAY": "라이트그레이",
        "LIGHTGRAY": "라이트그레이",
        "차콜그레이": "차콜",
    }
    return aliases.get(text.upper(), text.replace("라이트 그레이", "라이트그레이"))


def normalize_name(name):
    text = clean_text(name)
    text = re.sub(r"^\*", "", text).strip()
    text = re.sub(r"\s*\((블랙|화이트|라이트그레이|그레이|차콜|네이비|카키|브라운)\)$", "", text)
    text = text.replace("퀵 드라이", "퀵드라이")
    text = text.replace("세미 오버핏", "세미오버핏")
    text = text.replace("세비오버핏", "세미오버핏")
    text = text.replace("러닝 팬츠", "러닝팬츠")
    text = text.replace("채코제 에디션", "채코제에디션")
    text = text.replace("차코제에디션", "채코제에디션")
    return clean_text(text)


def split_name_color(name, color):
    name = normalize_name(name)
    color = norm_color(color)
    if name == "KINTERRA 벤치프레스 빈티지 그래픽 나시":
        color = {
            "블랙": "BLACK",
            "화이트": "WHITE",
            "라이트그레이": "LIGHT GREY",
            "그레이": "GREY",
        }.get(color, color)
    if color:
        return name, color
    suffixes = [
        (" 라이트그레이", "라이트그레이"),
        (" 멜란지그레이", "멜란지그레이"),
        (" 카키그레이", "카키그레이"),
        (" 차콜", "차콜"),
        (" 블랙", "블랙"),
        (" 화이트", "화이트"),
        (" 그레이", "그레이"),
        (" 세이지", "세이지"),
        (" 올리브", "올리브"),
    ]
    for suffix, suffix_color in suffixes:
        if name.endswith(suffix):
            return name[: -len(suffix)].strip(), suffix_color
    return name, color


def parse_option(option):
    color = ""
    size = ""
    for line in str(option or "").splitlines():
        if ":" not in line:
            continue
        key, value = [part.strip() for part in line.split(":", 1)]
        if key in {"색상", "컬러"}:
            color = norm_color(value)
        elif key == "사이즈":
            size = norm_size(value)
    return color, size


def as_int(value):
    if value in (None, ""):
        return 0
    return int(round(float(str(value).replace(",", ""))))


def infer_date(path):
    match = re.search(r"(20\d{6})", path.name)
    if not match:
        raise ValueError("파일명에서 생성일자를 찾을 수 없습니다. --date 값을 지정해 주세요.")
    generated = dt.datetime.strptime(match.group(1), "%Y%m%d").date()
    return (generated - dt.timedelta(days=1)).isoformat()


def load_items(path):
    items = []
    with path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            qty = as_int(row.get("판매수량"))
            gross = as_int(row.get("판매합계"))
            if qty == 0 and gross == 0:
                continue
            color, size = parse_option(row.get("옵션"))
            name, color = split_name_color(row.get("상품명"), color)
            items.append(
                {
                    "name": name,
                    "color": color,
                    "size": size,
                    "qty": qty,
                    "gross": gross,
                    "payment": 0,
                }
            )
    return items


def allocate_payments(items, target_payment):
    gross_total = sum(item["gross"] for item in items)
    if gross_total == 0:
        return
    running = 0
    for item in items:
        payment = int(round(item["gross"] * target_payment / gross_total))
        item["payment"] = payment
        running += payment
    if items:
        items[-1]["payment"] += target_payment - running


def standard_name(name, color, size):
    if color and size:
        return f"{name}_{color}-{size}"
    if color:
        return f"{name}_{color}"
    if size:
        return f"{name}-{size}"
    return name


def row_key(row):
    return (
        normalize_name(row.get("name")),
        norm_color(row.get("color")),
        norm_size(row.get("size")),
    )


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
    name_index = {}
    for i, row in enumerate(rows):
        if row.get("retailer") != RETAILER:
            continue
        key = row_key(row)
        index.setdefault(key, i)
        name_index.setdefault(key[0], i)
    return index, name_index


def append_daily(row, day, item):
    row.setdefault("daily", []).append(
        {
            "date": day,
            "qty": item["qty"],
            "gross": item["gross"],
            "payment": item["payment"],
            "orders": max(item["qty"], 0),
        }
    )
    row["qty"] = as_int(row.get("qty")) + item["qty"]
    row["gross"] = as_int(row.get("gross")) + item["gross"]
    row["payment"] = as_int(row.get("payment")) + item["payment"]
    row["orders"] = as_int(row.get("orders")) + max(item["qty"], 0)
    row["avg_unit"] = int(round(row["payment"] / row["qty"])) if row["qty"] else 0


def update_manual(day, items, target_payment):
    manual = json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    by_date = {entry["date"]: entry for entry in manual}
    entry = by_date.setdefault(day, {"date": day, "retailers": []})
    retailers = entry.setdefault("retailers", [])
    target = next((r for r in retailers if r.get("retailer") == RETAILER), None)
    if target is None:
        target = {"retailer": RETAILER, "items": []}
        retailers.append(target)
    target["payment"] = target_payment
    target["qty"] = sum(item["qty"] for item in items)
    target["orders"] = sum(max(item["qty"], 0) for item in items)
    target["items"] = []
    return [by_date[date] for date in sorted(by_date)]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_path")
    parser.add_argument("--date")
    parser.add_argument("--payment", type=int, required=True)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    csv_path = Path(args.csv_path)
    day = args.date or infer_date(csv_path)
    items = load_items(csv_path)
    allocate_payments(items, args.payment)

    rows = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    remove_existing_day(rows, day)
    index, name_index = build_index(rows)
    created = []

    for item in items:
        key = (item["name"], item["color"], item["size"])
        idx = index.get(key) or name_index.get(item["name"])
        if idx is None:
            new_row = {
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
                "standard_name": standard_name(item["name"], item["color"], item["size"]),
                "received_qty": 0,
                "stock_qty": 0,
            }
            rows.append(new_row)
            idx = len(rows) - 1
            index[key] = idx
            name_index[item["name"]] = idx
            created.append(new_row["standard_name"])
        row = rows[idx]
        row["name"] = item["name"]
        row["color"] = item["color"]
        row["size"] = item["size"]
        row["standard_name"] = standard_name(item["name"], item["color"], item["size"])
        row["match_status"] = "매칭완료"
        append_daily(row, day, item)

    manual = update_manual(day, items, args.payment)
    summary = {
        "date": day,
        "line_items": len(items),
        "qty": sum(item["qty"] for item in items),
        "gross": sum(item["gross"] for item in items),
        "payment": sum(item["payment"] for item in items),
        "created_rows": len(created),
        "created_preview": created,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if args.dry_run:
        return
    DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    MANUAL_FILE.write_text(json.dumps(manual, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
