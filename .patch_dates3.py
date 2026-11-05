import io
p = 'db.py'
s = io.open(p, encoding='utf-8', newline='').read()

def rep(old, new, count=1):
    global s
    assert s.count(old) == count, ('MISS', s.count(old), old[:70])
    s = s.replace(old, new)
    print('ok', old.splitlines()[0][:60])

if 'normalize_tx_date' not in s:
    rep('def looks_like_fuel(description: str) -> bool:',
'''_DATE_INPUT_FORMATS: Tuple[str, ...] = (\n    \"%Y-%m-%d\",\n    \"%d.%m.%Y\",\n    \"%d.%m.%y\",\n    \"%d/%m/%Y\",\n    \"%d/%m/%y\",\n    \"%d-%m-%Y\",\n    \"%d-%m-%y\",\n    \"%Y/%m/%d\",\n    \"%Y.%m.%d\",\n)\n\n\ndef normalize_tx_date(raw: Any, default: str = \"\") -> str:\n    \"\"\"Приводит любую дату к ISO YYYY-MM-DD для корректного сравнения в SQL.\n\n    Понимает ДД.ММ.ГГГГ, ДД.ММ.ГГ, ДД/ММ/ГГГГ, YYYY-MM-DD и дату-время.\n    На нераспознанном мусоре возвращает default (обычно сегодня), а не падает.\n    \"\"\"\n    text = str(raw or \"\").strip()\n    if not text:\n        return default\n\n    head = text.replace(\"T\", \" \").split(\" \", 1)[0].strip()\n    if len(head) < 6:\n        return default\n\n    for fmt in _DATE_INPUT_FORMATS:\n        try:\n            return datetime.strptime(head, fmt).date().isoformat()\n        except ValueError:\n            continue\n\n    return default\n\n\ndef looks_like_fuel(description: str) -> bool:''')

old = '        tx_date = (item.get(\"tx_date\") or \"\").strip() or today_str\n'
new = '        tx_date = normalize_tx_date(item.get(\"tx_date\"), today_str)\n'
rep(old, new, count=2)

io.open(p, 'w', encoding='utf-8', newline='').write(s)
print('dates patched')
