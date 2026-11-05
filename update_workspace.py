import sqlite3
with sqlite3.connect('data/finance_tracker.db') as conn:
    conn.execute("UPDATE transactions SET user_id = -100 WHERE author_name = 'Семья'")
    conn.commit()
    print('Готово! Семейные транзакции перенесены на общий workspace_id (-100).')
