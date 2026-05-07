import json
from pathlib import Path


DATA_FILE = Path("data.json")
RETAILER = "자사몰"


COLOR_SUFFIXES = [
    (" 라이트그레이", "라이트그레이"),
    (" 멜란지그레이", "멜란지그레이"),
    (" 카키그레이", "카키그레이"),
    (" 차콜", "차콜"),
    (" 블랙", "블랙"),
    (" 화이트", "화이트"),
    (" 네이비", "네이비"),
]


DIRECT_NAMES = {
    "[플래니트X박영감] 빈티지 워싱 스웻셔츠 멜란지그레이": (
        "[플래니트X박영감] 빈티지 워싱 스웻셔츠",
        "멜란지그레이",
    ),
    "퀵드라이 퍼포먼스 롱슬리브 블랙": ("퀵드라이 퍼포먼스 롱슬리브", "블랙"),
    "퀵드라이 퍼포먼스 숏슬리브 라이트블루": ("퀵드라이 퍼포먼스 숏슬리브", "라이트블루"),
    "퀵드라이 메쉬 러닝싱글렛 블랙": ("퀵드라이 메쉬 러닝싱글렛", "블랙"),
    "퀵드라이 메쉬 러닝싱글렛 라이트그레이": ("퀵드라이 메쉬 러닝싱글렛", "라이트그레이"),
    "TECH SWEAT 우먼즈 피그먼트 스웻쇼츠 차콜": (
        "TECH SWEAT 우먼즈 피그먼트 스웻쇼츠",
        "차콜",
    ),
}


def canonicalize(row):
    name = str(row.get("name") or "").strip()
    color = str(row.get("color") or "").strip()
    size = str(row.get("size") or "").strip()

    if name in DIRECT_NAMES:
        name, color = DIRECT_NAMES[name]
    elif not color:
        for suffix, suffix_color in COLOR_SUFFIXES:
            if name.endswith(suffix):
                name = name[: -len(suffix)].strip()
                color = suffix_color
                break

    size = {"XXL": "2XL", "XXXL": "3XL"}.get(size.upper(), size)
    color = {
        "BLACK": "블랙",
        "WHITE": "화이트",
        "CHARCOAL": "차콜",
        "LIGHTGRAY": "라이트그레이",
    }.get(color.upper(), color)

    return name, color, size


def standard_name(name, color, size):
    if color and size:
        return f"{name}_{color}-{size}"
    if color:
        return f"{name}_{color}"
    if size:
        return f"{name}-{size}"
    return name


def april_payment(row):
    return sum(
        int(daily.get("payment") or 0)
        for daily in row.get("daily", [])
        if str(daily.get("date", "")).startswith("2026-04-")
    )


def main():
    rows = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    repaired = []

    for row in rows:
        if row.get("retailer") != RETAILER:
            continue
        if row.get("match_status") != "미매칭":
            continue
        if april_payment(row) <= 0:
            continue

        name, color, size = canonicalize(row)
        row["name"] = name
        row["color"] = color
        row["size"] = size
        row["standard_name"] = standard_name(name, color, size)
        row["match_status"] = "매칭완료"
        row["source_type"] = row.get("source_type") or "상품"
        row["avg_unit"] = int(round(int(row.get("payment") or 0) / int(row.get("qty") or 1))) if int(row.get("qty") or 0) else 0
        repaired.append(row["standard_name"])

    DATA_FILE.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"repaired": repaired, "count": len(repaired)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
