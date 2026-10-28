"""Offline LCD order lookup. Port of the validated n8n Local Excel V3 matching rules.
No remote/network services. Never use purchase cost as a selling price.
"""
import re


def norm(value):
    return re.sub(r'[\s_-]+', '', str(value or '').upper().replace('\u00a0', ' ')).strip()


def model_code(value):
    return re.sub(r'^(INFINIX|TECNO|ITEL)', '', norm(value))


def product_code(value):
    return re.sub(r'^(MS|JK)\s*-\s*', '', str(value or '').upper()).strip()


def number(value):
    if value is None or isinstance(value, bool) or str(value).strip() == '':
        return None
    try:
        n = float(value)
        return round(n, 2) if n == n and abs(n) != float('inf') else None
    except (TypeError, ValueError):
        return None


def contains_token(text, term):
    source = str(text or '').upper().replace('\u00a0', ' ')
    target = str(term or '').upper().replace('\u00a0', ' ').strip()
    return bool(target and re.search(r'(^|[^A-Z0-9])' + re.escape(target) + r'(?=$|[^A-Z0-9])', source))


def lookup(catalogue, query, quantity=1):
    """Return V3-compatible result. Quantities and stock refer only to exact SKUs."""
    query = str(query or '').strip()
    key = norm(query)
    if not key:
        raise ValueError('Enter a phone model or LCD SKU.')
    qty = number(quantity)
    if qty is not None and qty <= 0:
        qty = None

    compatible = [r for r in catalogue['compatibility'] if (
        norm(r.get('名称')) == key or norm(r.get('型号')) == key or
        norm(r.get(' ')) == key or norm(r.get('系列')) == key or
        model_code(r.get('型号')) == key
    )]

    codes = {key}
    names = {key}
    for row in compatible:
        if row.get('型号'):
            codes.add(model_code(row['型号']))
        if row.get(' '):
            codes.add(model_code(row[' ']))
        if row.get('系列'):
            names.add(norm(row['系列']))
        if row.get('名称'):
            names.add(norm(row['名称']))

    products = []
    seen = set()
    for row in catalogue['prices']:
        model_raw = row.get('Model')
        if not model_raw:
            continue
        identity = (row.get('_source_sheet', ''), model_raw,
                    row.get('Compatible models', ''), row.get('row_number'))
        if identity in seen:
            continue
        seen.add(identity)
        model = product_code(model_raw)
        compatible_text = str(row.get('Compatible models') or '')
        fragments = [norm(f) for f in compatible_text.split('/')]
        matches = (
            norm(model) == key or
            any(len(c) >= 3 and (contains_token(model, c) or contains_token(compatible_text, c)) for c in codes) or
            any(len(n) >= 6 and n in fragments for n in names)
        )
        if not matches:
            continue

        store_rows = [s for s in catalogue['inventory']
                      if norm(product_code(model_raw)) == norm(s.get('col_3'))]
        stock = sum(max(0, number(s.get('col_19')) or 0) for s in store_rows)
        sufficient = (stock >= qty) if qty is not None else None
        selling = number(row.get('_main_selling_price'))
        selling = selling if selling is not None and selling > 0 else None
        in_store = stock > 0
        reasons = []
        if selling is None:
            reasons.append('SELLING_PRICE_MISSING')
        if not in_store:
            reasons.append('NOT_IN_LOCAL_STOCK')
        if in_store and sufficient is False:
            reasons.append('INSUFFICIENT_LOCAL_STOCK')
        if sufficient is None:
            reasons.append('QUANTITY_REVIEW')

        products.append({
            'product': model_raw,
            'source_sheet': row.get('_source_sheet', 'UNKNOWN'),
            'main_selling_price': selling,
            'price_confirmed': selling is not None,
            'customer_quote_allowed': selling is not None,
            'store_quantity': stock,
            'source': 'IN_STORE' if in_store else 'BIG_WAREHOUSE_INDICATED',
            'sufficient_store_stock': sufficient,
            'warehouse_confirmation_required': not in_store or sufficient is not True,
            'inventory_rows': [s.get('row_number') for s in store_rows],
            'price_row': row.get('row_number'),
            'review_reasons': reasons,
        })

    products.sort(key=lambda p: norm(product_code(p['product'])) == key, reverse=True)
    any_missing = any(not p['price_confirmed'] for p in products)
    any_warehouse = any(p['warehouse_confirmation_required'] for p in products)
    if not products:
        status = 'NOT_FOUND_IN_LOADED_SHEETS'
    elif len(products) > 1:
        status = 'REVIEW_MULTIPLE_MATCHES'
    elif any_missing:
        status = 'SELLING_PRICE_REVIEW'
    elif any_warehouse:
        status = 'WAREHOUSE_OR_SHORTAGE_REVIEW'
    else:
        status = 'MATCH_FOUND'
    needs_review = len(products) != 1 or any_missing or any_warehouse or qty is None
    return {
        'requested_model': query,
        'requested_quantity': qty,
        'compatible_codes': sorted(codes),
        'products': products,
        'match_count': len(products),
        'status': status,
        'needs_review': needs_review,
        'automatic_quote_allowed': len(products) == 1 and not needs_review and products[0]['price_confirmed'],
        'catalogue_source': 'LOCAL_EXCEL_JSON',
        'catalogue_generated_at': catalogue.get('generated_at'),
        'catalogue_counts': catalogue.get('counts', {}),
    }
