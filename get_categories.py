import sqlite3, json

with sqlite3.connect('data/finance_tracker.db') as conn:
    cursor = conn.execute("SELECT DISTINCT category FROM transactions WHERE category != '' AND category IS NOT NULL ORDER BY category")
    categories = [row[0] for row in cursor.fetchall()]
    
    print('Отправь папе это сообщение:\n')
    print('Привет! Вот текущий массив категорий расходов, которые зафиксированы в базе бота:\n')
    
    print(json.dumps(categories, ensure_ascii=False, indent=2))
    
    print('\n*Примечание: сейчас бот использует ИИ для распределения трат, поэтому жесткой привязки к словам-триггерам нет. Жду от тебя спецификацию с ключевыми словами, чтобы я мог жестко зашить их в логику парсера!*')
