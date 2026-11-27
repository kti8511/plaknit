"""Export the selected WMS workbook as a dated, SKU-level stock snapshot."""
import argparse
import json
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent


def export_snapshot(source):
    source = Path(source)
    book = load_workbook(source, read_only=True, data_only=True)
    records = book.active.iter_rows(values_only=True)
    headers = [str(v or '').strip() for v in next(records)]
    required = ['상품코드', '출고상품명', '바코드', '총재고', '출고가능']
    if not all(k in headers for k in required):
        raise ValueError('WMS stock columns missing')
    products = []
    seen = set()
    for values in records:
        r = dict(zip(headers, values))
        if not r.get('출고상품명') or str(r.get('회사명') or '').strip() == '합계':
            continue
        code = str(r.get('상품코드') or '').strip()
        barcode = str(r.get('바코드') or '').strip()
        key = 'code:' + code if code else barcode
        if key in seen:
            raise ValueError('Duplicate SKU in WMS stock export: ' + key)
        seen.add(key)
        def number(k):
            return int(str(r.get(k) or 0).replace(',', ''))
        products.append(dict(key=key, barcode=barcode, code=code,
            name=str(r['출고상품명']), stock=number('총재고'),
            available=number('출고가능'), instructed=number('출고지시'),
            processing=number('출고작업중'), incoming=number('입고작업중'),
            customer_return=number('고객반품'), defective=number('불량'),
            last_in=str(r.get('최근입고') or ''), last_out=str(r.get('최근출고') or '')))
    book.close()
    stamp = re.search(r'(\d{4}-\d{2}-\d{2})_(\d{6})', source.stem)
    captured = datetime.strptime(''.join(stamp.groups()), '%Y-%m-%d%H%M%S').replace(tzinfo=ZoneInfo('Asia/Seoul')) if stamp else datetime.fromtimestamp(source.stat().st_mtime,ZoneInfo('Asia/Seoul'))
    payload = dict(source=source.name, basis='총재고',
        captured_at=captured.isoformat(timespec='seconds'),
        products=products)
    (ROOT/'inventory_snapshot.json').write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf8')
    print(json.dumps(dict(products=len(products), source=source.name, basis='총재고'), ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('source')
    export_snapshot(parser.parse_args().source)
