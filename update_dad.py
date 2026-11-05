import sqlite3, pandas as pd

with sqlite3.connect('data/finance_tracker.db') as conn:
    sessions = pd.read_sql('SELECT * FROM user_sessions', conn)
    print('--- Все пользователи в базе ---')
    print(sessions[['tg_user_id', 'display_name', 'mode']].to_string(index=False))
    
    dad_session = sessions[sessions['tg_user_id'] != 853218288]
    
    if not dad_session.empty:
        dad_id = int(dad_session.iloc[0]['tg_user_id'])
        
        conn.execute("UPDATE transactions SET user_id = ? WHERE author_name = 'Семья'", (dad_id,))
        conn.commit()
        print(f'\nУспех! Все семейные транзакции привязаны к аккаунту папы (ID: {dad_id}).')
        print('Теперь он может открыть бота и увидеть всю статистику за октябрь и ноябрь!')
    else:
        print('\nНовый аккаунт пока не найден. Возможно, бот еще не успел записать его в базу.')
