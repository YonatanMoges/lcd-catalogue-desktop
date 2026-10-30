"""Build the existing LCD catalogue JSON from the three business Excel workbooks.
Read-only for Excel; an invalid conversion never replaces the working JSON.
"""
import json
import os
import tempfile
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

COMPAT_SHEETS = ('传音', 'VIVO', 'OPPO', '红米丶小米', '三星', 'MOTO', '华为')
PRICE_SHEETS = ('CY', 'SM', 'OP', 'VI', 'HM', 'XM', 'HW', 'SM WF')
INVENTORY_SHEET = 'Office available models '


def clean(value):
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip() if not value.strip() else value
    return value


def numeric(value):
    if value is None or value == '':
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def build_catalogue(compat_path, price_path, inventory_path):
    paths = [Path(p) for p in (compat_path, price_path, inventory_path)]
    for p in paths:
        if not p.is_file():
            raise ValueError(f'Excel file not found: {p}')
    workbooks = [load_workbook(p, read_only=True, data_only=True) for p in paths]
    try:
        compat_wb, price_wb, inventory_wb = workbooks
        for wb, required in [(compat_wb, COMPAT_SHEETS), (price_wb, PRICE_SHEETS),
                             (inventory_wb, (INVENTORY_SHEET,))]:
            for sheet in required:
                if sheet not in wb.sheetnames:
                    raise ValueError(f'Missing expected worksheet: {sheet}')

        compatibility, prices, inventory = [], [], []
        sheets = {}
        for sheet in COMPAT_SHEETS:
            ws = compat_wb[sheet]
            accepted = 0
            for rownum, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if rownum == 1:
                    continue
                fields = list(row[:5]) + [None] * max(0, 5 - len(row))
                if not fields[2] or str(fields[2]).strip() == '名称':
                    continue
                if sheet == '传音':
                    record = {'row_number': rownum, '内部规格型号': clean(fields[0]),
                              'col_2': clean(fields[1]), '名称': clean(fields[2]),
                              '型号': clean(fields[3]), '备注': clean(fields[4])}
                else:
                    record = {'row_number': rownum, '内部型号': clean(fields[0]),
                              '系列' if sheet in ('VIVO','OPPO','MOTO','华为') else 'col_2': clean(fields[1]),
                              '名称': clean(fields[2]), '型号': clean(fields[3]),
                              '备注': clean(fields[4])}
                record['_source_sheet'] = sheet
                compatibility.append(record)
                accepted += 1
            sheets['compat/' + sheet] = {'accepted_rows': accepted}

        for sheet in PRICE_SHEETS:
            ws = price_wb[sheet]
            accepted = 0
            for rownum, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if rownum == 1:
                    continue
                values = list(row) + [None] * max(0, 7 - len(row))
                if not values[1] or str(values[1]).strip() == 'Model':
                    continue
                if sheet == 'SM WF':
                    record = {'row_number': rownum, 'Model': clean(values[1]),
                              'Price': clean(values[2]), 'Compatible models': '',
                              'Remark': '', 'col_5': None,
                              '_main_selling_price': None}
                else:
                    selling = numeric(values[5])
                    record = {'row_number': rownum, 'Model': clean(values[1]),
                              'Compatible models': clean(values[2]),
                              'Price': clean(values[3]),
                              'Remark': clean(values[4]) if sheet not in ('HM',) else '',
                              'col_5': clean(values[5]) if values[5] is not None else None,
                              '_main_selling_price': selling if selling is not None and selling > 0 else None}
                    if sheet == 'CY' and values[6] is not None:
                        record['col_6'] = clean(values[6])
                    if sheet == 'HM':
                        record['Sale price'] = clean(values[4])
                record['_source_sheet'] = sheet
                prices.append(record)
                accepted += 1
            sheets['prices/' + sheet] = {'accepted_rows': accepted}

        ws = inventory_wb[INVENTORY_SHEET]
        accepted = 0
        for rownum, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if rownum < 4:
                continue
            vals = list(row[:24]) + [None] * max(0, 24 - len(row))
            if not vals[2] or not isinstance(vals[0], (int, float)):
                continue
            stock_total = numeric(vals[18])
            if stock_total is None:
                raise ValueError(f'Invalid Total stock in row {rownum}; keeping old catalogue')
            named = {1:'TYPE', 2:'MODEL', 3:'Selling Price', 4:'Stock', 5:'Restock',
                     6:'SOLD OUT', 12:'Sold out', 18:'Total'}
            record = {'row_number': rownum}
            for idx in range(24):
                key = named.get(idx, f'col_{idx + 1}')
                record[key] = clean(vals[idx])
            record['col_2'] = clean(vals[1])
            record['col_3'] = clean(vals[2])
            record['col_4'] = clean(vals[3])
            record['col_19'] = stock_total
            record['_source_sheet'] = INVENTORY_SHEET.strip()
            inventory.append(record)
            accepted += 1
        sheets['inventory/' + INVENTORY_SHEET.strip()] = {'accepted_rows': accepted}

        if len(compatibility) < 100 or len(prices) < 20 or len(inventory) < 5:
            raise ValueError('Too few catalogue records; refusing to replace last good data')
        # Require a valid majority of prices to prevent accepting an unrelated spreadsheet.
        if sum(bool(p.get('Model')) for p in prices) != len(prices):
            raise ValueError('Invalid price workbook: missing LCD product identifiers')
        return {'schema_version': 1, 'generated_at': datetime.now().astimezone().isoformat(),
                'counts': {'compatibility': len(compatibility), 'prices': len(prices),
                           'inventory': len(inventory)}, 'sheets': sheets,
                'compatibility': compatibility, 'prices': prices, 'inventory': inventory}
    finally:
        for wb in workbooks:
            wb.close()


def save_catalogue_atomic(data, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    tmp = None
    try:
        with tempfile.NamedTemporaryFile('w', encoding='utf-8', suffix='.tmp',
                                         prefix='.catalogue-', dir=path.parent, delete=False) as f:
            tmp = Path(f.name)
            f.write(serialized)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if tmp and tmp.exists():
            tmp.unlink()
