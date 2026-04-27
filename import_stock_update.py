import json
import re
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "data.json"
UNMATCHED_FILE = ROOT / "stock_unmatched.json"
DESKTOP = Path.home() / "Desktop"


def normalize(value):
    return re.sub(r"\s+", "", str(value or "")).upper()


def normalize_name(value):
    text = normalize(value)
    text = re.sub(r"^\[[^\]]+\]", "", text)
    return re.sub(r"[-_()（）\[\]/·,]", "", text)


def latest_stock_file():
    files = [
        p
        for p in DESKTOP.glob("*2026-04-27_151321.xlsx")
        if not p.name.startswith("~$")
    ]
    if not files:
        files = [p for p in DESKTOP.glob("*.xlsx") if "재고조회" in p.name and not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError("재고조회 엑셀 파일을 찾지 못했습니다.")
    return max(files, key=lambda p: p.stat().st_mtime)


def load_stock_rows(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    headers = [str(v).strip() if v is not None else "" for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {h: i for i, h in enumerate(headers)}
    required = ["출고상품명", "바코드", "총재고"]
    missing = [h for h in required if h not in idx]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}")

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        barcode = str(row[idx["바코드"]] or "").strip()
        outbound_name = str(row[idx["출고상품명"]] or "").strip()
        stock_raw = row[idx["총재고"]]
        try:
            stock_qty = int(float(str(stock_raw or 0).replace(",", "")))
        except ValueError:
            stock_qty = 0
        if not barcode and not outbound_name:
            continue
        rows.append(
            {
                "barcode": barcode,
                "outbound_name": outbound_name,
                "stock_qty": stock_qty,
            }
        )
    return rows


def main():
    stock_file = latest_stock_file()
    data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    stock_rows = load_stock_rows(stock_file)

    by_barcode = {normalize(r["barcode"]): r for r in stock_rows if r["barcode"]}
    by_name = {normalize(r["outbound_name"]): r for r in stock_rows if r["outbound_name"]}
    stock_name_pairs = [(normalize_name(r["outbound_name"]), r) for r in stock_rows if r["outbound_name"]]

    matched = 0
    unmatched_data = []
    used_stock_keys = set()

    for item in data:
        candidates = [
            normalize(item.get("match_sku")),
            normalize(item.get("standard_name")),
            normalize(item.get("name")),
        ]
        stock_row = None
        match_key = None
        for key in candidates:
            if key and key in by_barcode:
                stock_row = by_barcode[key]
                match_key = key
                break
        if stock_row is None:
            for key in candidates:
                if key and key in by_name:
                    stock_row = by_name[key]
                    match_key = key
                    break
        if stock_row is None:
            name_candidates = [
                normalize_name(item.get("standard_name")),
                normalize_name(item.get("name")),
            ]
            name_candidates = [v for v in name_candidates if len(v) >= 6]
            matched_by_name = []
            for candidate in name_candidates:
                matched_by_name = [r for stock_name, r in stock_name_pairs if candidate in stock_name or stock_name in candidate]
                if matched_by_name:
                    break
            if matched_by_name:
                stock_row = {
                    "barcode": "",
                    "outbound_name": matched_by_name[0]["outbound_name"],
                    "stock_qty": sum(r["stock_qty"] for r in matched_by_name),
                }
                match_key = normalize_name(stock_row["outbound_name"])

        if stock_row:
            item["stock_qty"] = stock_row["stock_qty"]
            item["stock_barcode"] = stock_row["barcode"]
            item["stock_name"] = stock_row["outbound_name"]
            matched += 1
            if match_key:
                used_stock_keys.add(match_key)
        else:
            unmatched_data.append(
                {
                    "retailer": item.get("retailer"),
                    "match_sku": item.get("match_sku"),
                    "standard_name": item.get("standard_name"),
                    "name": item.get("name"),
                    "color": item.get("color"),
                    "size": item.get("size"),
                }
            )

    unmatched_stock = [
        r
        for r in stock_rows
        if normalize(r["barcode"]) not in used_stock_keys and normalize(r["outbound_name"]) not in used_stock_keys
    ]

    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    UNMATCHED_FILE.write_text(
        json.dumps(
            {
                "stock_file": str(stock_file),
                "data_rows": len(data),
                "stock_rows": len(stock_rows),
                "matched_data_rows": matched,
                "unmatched_data_rows": unmatched_data,
                "unmatched_stock_rows": unmatched_stock,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "stock_file": str(stock_file),
                "data_rows": len(data),
                "stock_rows": len(stock_rows),
                "matched_data_rows": matched,
                "unmatched_data_rows": len(unmatched_data),
                "unmatched_stock_rows": len(unmatched_stock),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
