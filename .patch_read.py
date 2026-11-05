import io
p = 'db.py'
s = io.open(p, encoding='utf-8', newline='').read()

def rep(old, new):
    global s
    assert s.count(old) == 1, ('MISS', s.count(old), old[:70])
    s = s.replace(old, new)
    print('ok', old.splitlines()[0][:60])

# get_range_stats: принимаем и ISO, и ДД.ММ.ГГГГ
rep('''    start_date = str(start_iso)[:10]\n    end_date = str(end_iso)[:10]\n    if start_date > end_date:''',
'''    start_date = normalize_tx_date(start_iso) or str(start_iso)[:10]\n    end_date = normalize_tx_date(end_iso) or str(end_iso)[:10]\n    if start_date > end_date:''')

# get_monthly_stats: нормализуем формат года-месяца
rep('''    stats = _aggregate_period_stats(workspace_id, start_date, end_date)\n    stats[\"year_month\"] = year_month''',
'''    stats = _aggregate_period_stats(workspace_id, start_date, end_date)\n    stats[\"year_month\"] = year_month\n    stats[\"is_range\"] = False''')

# export: те же правила для периода
rep('''        where_sql += \" AND tx_date >= ?\"\n        params.append(str(date_from)[:10])''',
'''        where_sql += \" AND tx_date >= ?\"\n        params.append(normalize_tx_date(date_from) or str(date_from)[:10])''')

rep('''        where_sql += \" AND tx_date <= ?\"\n        params.append(str(date_to)[:10])''',
'''        where_sql += \" AND tx_date <= ?\"\n        params.append(normalize_tx_date(date_to) or str(date_to)[:10])''')

io.open(p, 'w', encoding='utf-8', newline='').write(s)
print('read-side patched')
