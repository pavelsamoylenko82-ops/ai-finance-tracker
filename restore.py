import pandas as pd, sqlite3, os, uuid

p1 = r'D:\my_projects\personal_finance (3).xlsx'
if not os.path.exists(p1): p1 = r'D:\my_projects\personal_finance (6).xlsx'
p2 = r'D:\my_projects\family_finance_2.xlsx'
if not os.path.exists(p2): p2 = r'D:\my_projects\family_finance.xlsx'

d1 = pd.read_excel(p1, sheet_name='Транзакции')
d1['author_name'] = 'Паша'
d2 = pd.read_excel(p2, sheet_name=0)
d2['author_name'] = 'Семья'

df = pd.concat([d1, d2], ignore_index=True)
df = df.rename(columns={'Дата':'tx_date', 'Тип':'tx_type', 'Категория':'category', 'Описание':'description', 'Сумма':'amount', 'Валюта':'currency'})
df['tx_type'] = df['tx_type'].astype(str).str.lower().str.strip().apply(lambda x: 'income' if x in ['доход', 'income', '+'] else 'expense')

df['amount_byn'] = df['amount']
df['project_tag'] = ''
df['description'] = df['description'].fillna('')
df['batch_id'] = str(uuid.uuid4())

# Исправленный парсинг дат (теперь 2026-11-01 останется ноябрем)
df['tx_date'] = pd.to_datetime(df['tx_date']).dt.strftime('%Y-%m-%d %H:%M:%S')

with sqlite3.connect('data/finance_tracker.db') as conn:
    real_id = conn.execute('SELECT tg_user_id FROM user_sessions LIMIT 1').fetchone()[0]
    df['user_id'] = real_id
    
    # Очищаем таблицу от старых кривых записей перед чистой заливкой
    conn.execute('DELETE FROM transactions') 
    
    df = df[['tx_date', 'tx_type', 'category', 'description', 'amount', 'amount_byn', 'currency', 'author_name', 'project_tag', 'user_id', 'batch_id']]
    df.to_sql('transactions', conn, if_exists='append', index=False)
    
    print('--- Сессии пользователей ---')
    print(pd.read_sql('SELECT * FROM user_sessions', conn).to_string())
    print('\n--- Колонки таблицы transactions ---')
    print(', '.join(pd.read_sql('PRAGMA table_info(transactions)', conn)['name']))

print('Данные залиты с правильными датами!')
