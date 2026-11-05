import io
import json
import os
import sqlite3
import tempfile
import uuid
from contextlib import contextmanager
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from config import (
    CHILD_CATEGORIES,
    DB_PATH,
    DEFAULT_CURRENCY,
    EXPENSE_CATEGORIES,
    FAMILY_WORKSPACE_ID,
    INCOME_CATEGORIES,
    MAX_FAMILY_MEMBERS,
    now_minsk,
)
from currency import convert_to_byn

# In-memory кэш активных сессий для мгновенной работы Middleware без нагрузки на SQLite
_SESSION_CACHE: Dict[int, Dict[str, Any]] = {}

# --- Справочник категорий: единая точка нормализации для БД, бота и Excel ---
CANONICAL_CATEGORIES: Dict[str, str] = {
    c.lower(): c for c in (*EXPENSE_CATEGORIES, *INCOME_CATEGORIES)
}

# Старые/произвольные названия, которые могли прийти от ИИ или из прошлых версий
LEGACY_CATEGORY_ALIASES: Dict[str, str] = {
    "авто и топливо": "Авто",
    "авто и расходники": "Авто",
    "автомобиль": "Авто",
    "продукты и супермаркеты": "Продукты",
    "супермаркеты": "Продукты",
    "кафе": "Кафе и фастфуд",
    "фастфуд": "Кафе и фастфуд",
    "здоровье": "Гигиена и здоровье",
    "аптека": "Гигиена и здоровье",
    "развлечение": "Развлечения",
    "подарки": "Развлечения",
    "прочее": "Прочие расходы",
}

# Ключевые слова топлива для безопасного сплита старых записей «Авто и топливо»
_FUEL_MARKERS: Tuple[str, ...] = (
    "бензин", "топливо", "заправк", "азс", "дизель", "солярк", "дт",
    "аи-92", "аи-95", "аи-98", "аи92", "аи95", "аи98", "евро-95",
    "евро95", "белнефтепродукт", "газпром", "лукойл", "литров",
)


_DATE_INPUT_FORMATS: Tuple[str, ...] = (
    "%Y-%m-%d",
    "%d.%m.%Y",
    "%d.%m.%y",
    "%d/%m/%Y",
    "%d/%m/%y",
    "%d-%m-%Y",
    "%d-%m-%y",
    "%Y/%m/%d",
    "%Y.%m.%d",
)


def normalize_tx_date(raw: Any, default: str = "") -> str:
    """Приводит любую дату к ISO YYYY-MM-DD для корректного сравнения в SQL.

    Понимает ДД.ММ.ГГГГ, ДД.ММ.ГГ, ДД/ММ/ГГГГ, YYYY-MM-DD и дату-время.
    На нераспознанном мусоре возвращает default (обычно сегодня), а не падает.
    """
    text = str(raw or "").strip()
    if not text:
        return default

    head = text.replace("T", " ").split(" ", 1)[0].strip()
    if len(head) < 6:
        return default

    for fmt in _DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(head, fmt).date().isoformat()
        except ValueError:
            continue

    return default


def looks_like_fuel(description: str) -> bool:
    """Грубая, но безопасная эвристика: относить ли старую запись к «Топливу»."""
    text = (description or "").lower()
    return any(marker in text for marker in _FUEL_MARKERS)


def normalize_category(category: str, tx_type: str = "expense") -> str:
    """Приводит категорию к каноническому виду из справочника config.py.

    Никогда не бросает исключений: неизвестная категория возвращается как есть,
    чтобы не терять пользовательские данные.
    """
    fallback = "Прочие доходы" if tx_type == "income" else "Прочие расходы"
    raw = (category or "").strip()
    if not raw:
        return fallback

    low = raw.lower()
    if low in CANONICAL_CATEGORIES:
        return CANONICAL_CATEGORIES[low]
    if low in LEGACY_CATEGORY_ALIASES:
        return LEGACY_CATEGORY_ALIASES[low]

    # Мягкое префиксное совпадение: «Кафе» -> «Кафе и фастфуд», «Продукт» -> «Продукты»
    for canon_low, canon in CANONICAL_CATEGORIES.items():
        if len(low) >= 3 and canon_low.startswith(low):
            return canon
    return raw


def _split_fuel_category(
    category: str, tx_type: str, description: str
) -> str:
    """Страховка от модели: «бензин» в категории «Авто» уводим в «Топливо»."""
    if tx_type == "expense" and category == "Авто" and looks_like_fuel(description):
        return "Топливо"
    return category


@contextmanager
def get_connection():
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=15000;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    with get_connection() as conn:
        conn.execute("PRAGMA journal_mode=WAL;")

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS user_sessions (
                tg_user_id INTEGER PRIMARY KEY,
                workspace_id INTEGER NOT NULL,
                mode TEXT NOT NULL CHECK(mode IN ('family', 'personal')),
                display_name TEXT NOT NULL DEFAULT '',
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        # Постоянный реестр участников семейного счета (защита слотов 2 из 2 навсегда)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS family_members (
                tg_user_id INTEGER PRIMARY KEY,
                display_name TEXT NOT NULL DEFAULT '',
                joined_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        # Переносим уже активных семейных пользователей в постоянный реестр при старте
        existing_family = conn.execute(
            "SELECT tg_user_id, display_name FROM user_sessions WHERE workspace_id = ?",
            (FAMILY_WORKSPACE_ID,),
        ).fetchall()
        for r in existing_family:
            conn.execute(
                "INSERT OR IGNORE INTO family_members (tg_user_id, display_name) VALUES (?, ?)",
                (int(r["tg_user_id"]), r["display_name"]),
            )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                batch_id TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                tx_date TEXT NOT NULL DEFAULT '',
                tx_type TEXT NOT NULL CHECK(tx_type IN ('income', 'expense')),
                amount REAL NOT NULL CHECK(amount > 0),
                currency TEXT NOT NULL,
                amount_byn REAL NOT NULL DEFAULT 0,
                category TEXT NOT NULL,
                description TEXT NOT NULL,
                project_tag TEXT NOT NULL DEFAULT '',
                author_name TEXT NOT NULL DEFAULT ''
            )
            """
        )

        existing_cols = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(transactions);").fetchall()
        }
        if "tx_date" not in existing_cols:
            conn.execute("ALTER TABLE transactions ADD COLUMN tx_date TEXT NOT NULL DEFAULT '';")
        if "amount_byn" not in existing_cols:
            conn.execute("ALTER TABLE transactions ADD COLUMN amount_byn REAL NOT NULL DEFAULT 0;")
        if "project_tag" not in existing_cols:
            conn.execute("ALTER TABLE transactions ADD COLUMN project_tag TEXT NOT NULL DEFAULT '';")
        if "author_name" not in existing_cols:
            conn.execute("ALTER TABLE transactions ADD COLUMN author_name TEXT NOT NULL DEFAULT '';")

        _migrate_categories(conn)

        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_tx_user_date ON transactions (user_id, tx_date);"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_tx_user_type_date ON transactions (user_id, tx_type, tx_date);"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_tx_batch ON transactions (batch_id, user_id);"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_tx_project ON transactions (user_id, project_tag);"
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS budgets (
                user_id INTEGER NOT NULL,
                category TEXT NOT NULL,
                monthly_limit REAL NOT NULL CHECK(monthly_limit > 0),
                PRIMARY KEY (user_id, category)
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS pending_imports (
                import_id TEXT PRIMARY KEY,
                user_id INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                payload_json TEXT NOT NULL
            )
            """
        )
        # Очистка старых неподтвержденных выписок (старше 24 часов)
        conn.execute(
            "DELETE FROM pending_imports WHERE created_at < datetime('now', '-1 day');"
        )


def _migrate_categories(conn: sqlite3.Connection) -> None:
    """Идемпотентная миграция: старые категории -> новые, «Авто и топливо» -> «Авто»/«Топливо».

    Работает только с затронутыми строками, чтобы не сканировать всю таблицу на каждом старте.
    Никогда не роняет приложение: любые ошибки логируются и игнорируются.
    """
    try:
        canonical_lower = tuple(CANONICAL_CATEGORIES.keys())
        placeholders = ",".join("?" for _ in canonical_lower)

        legacy_rows = conn.execute(
            f"""
            SELECT id, tx_type, category, description
            FROM transactions
            WHERE LOWER(TRIM(category)) NOT IN ({placeholders})
            """,
            canonical_lower,
        ).fetchall()

        updates: List[Tuple[str, int]] = []
        for row in legacy_rows:
            fixed = normalize_category(row["category"], row["tx_type"])
            fixed = _split_fuel_category(fixed, row["tx_type"], row["description"])
            if fixed and fixed != row["category"]:
                updates.append((fixed, int(row["id"])))

        # Спilt старых «Авто»-записей, где явно фигурирует топливо
        auto_rows = conn.execute(
            """
            SELECT id, description
            FROM transactions
            WHERE tx_type = 'expense' AND category = 'Авто'
            """
        ).fetchall()
        for row in auto_rows:
            if looks_like_fuel(row["description"]):
                updates.append(("Топливо", int(row["id"])))

        if updates:
            conn.executemany("UPDATE transactions SET category = ? WHERE id = ?", updates)

        # Бюджеты: старую общую статью переносим на «Топливо» (именно её обычно и лимитировали)
        conn.execute(
            "UPDATE OR REPLACE budgets SET category = 'Топливо' "
            "WHERE LOWER(TRIM(category)) IN ('авто и топливо', 'топливо')"
        )
        conn.execute(
            "UPDATE OR REPLACE budgets SET category = 'Авто' "
            "WHERE LOWER(TRIM(category)) IN ('авто и расходники', 'автомобиль')"
        )
    except Exception:
        # Миграция не должна мешать запуску бота
        pass


def get_user_session(tg_user_id: int) -> Optional[Dict[str, Any]]:
    cached = _SESSION_CACHE.get(tg_user_id)
    if cached is not None:
        return cached

    with get_connection() as conn:
        row = conn.execute(
            "SELECT tg_user_id, workspace_id, mode, display_name FROM user_sessions WHERE tg_user_id = ?",
            (tg_user_id,),
        ).fetchone()
        if row:
            data = dict(row)
            _SESSION_CACHE[tg_user_id] = data
            return data
        return None


def try_register_family_member(tg_user_id: int, display_name: str) -> bool:
    """
    Проверяет и навсегда закрепляет пользователя в семейном бюджете (максимум 2 человека).
    Даже если участник семьи нажмет /switch, его место останется за ним.
    """
    with get_connection() as conn:
        rows = conn.execute("SELECT tg_user_id FROM family_members").fetchall()
        member_ids = {int(r["tg_user_id"]) for r in rows}

        if tg_user_id in member_ids:
            return True

        if len(member_ids) >= MAX_FAMILY_MEMBERS:
            return False

        conn.execute(
            "INSERT INTO family_members (tg_user_id, display_name) VALUES (?, ?)",
            (tg_user_id, display_name.strip() or "Участник"),
        )
        return True


def set_user_session(tg_user_id: int, workspace_id: int, mode: str, display_name: str) -> None:
    clean_name = display_name.strip() or "Пользователь"
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO user_sessions (tg_user_id, workspace_id, mode, display_name, updated_at)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(tg_user_id) DO UPDATE SET
                workspace_id = excluded.workspace_id,
                mode = excluded.mode,
                display_name = excluded.display_name,
                updated_at = CURRENT_TIMESTAMP
            """,
            (tg_user_id, workspace_id, mode, clean_name),
        )
    _SESSION_CACHE[tg_user_id] = {
        "tg_user_id": tg_user_id,
        "workspace_id": workspace_id,
        "mode": mode,
        "display_name": clean_name,
    }


def clear_user_session(tg_user_id: int) -> None:
    _SESSION_CACHE.pop(tg_user_id, None)
    with get_connection() as conn:
        conn.execute("DELETE FROM user_sessions WHERE tg_user_id = ?", (tg_user_id,))


def get_family_member_ids(exclude_tg_id: int) -> List[int]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT tg_user_id FROM family_members WHERE tg_user_id != ?",
            (exclude_tg_id,),
        ).fetchall()
        return [int(r["tg_user_id"]) for r in rows]


def add_transactions(
    workspace_id: int,
    items: List[Dict[str, Any]],
    default_currency: str,
    author_name: str = "",
    batch_id: Optional[str] = None,
) -> str:
    batch_id = batch_id or uuid.uuid4().hex[:12]
    now_dt = now_minsk()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    today_str = now_dt.strftime("%Y-%m-%d")

    # Вычисляем курсы валют ДО открытия транзакции БД, чтобы не блокировать SQLite
    prepared_rows = []
    for item in items:
        raw_amount = round(float(item["amount"]), 2)
        cur = (item.get("currency") or default_currency).upper().strip()
        amount_byn = convert_to_byn(raw_amount, cur)
        tx_date = normalize_tx_date(item.get("tx_date"), today_str)
        tag = (item.get("project_tag") or "").strip().lstrip("#").lower()
        tx_type = item["tx_type"]
        category = normalize_category(item.get("category") or "", tx_type)
        category = _split_fuel_category(category, tx_type, item.get("description") or "")
        prepared_rows.append(
            (
                workspace_id,
                batch_id,
                now_str,
                tx_date,
                tx_type,
                raw_amount,
                cur,
                amount_byn,
                category,
                item["description"],
                tag,
                author_name,
            )
        )

    with get_connection() as conn:
        conn.executemany(
            """
            INSERT INTO transactions (
                user_id, batch_id, created_at, tx_date, tx_type,
                amount, currency, amount_byn, category, description, project_tag, author_name
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            prepared_rows,
        )
    return batch_id


def get_batch_stats(workspace_id: int, batch_id: str) -> Optional[Dict[str, Any]]:
    """Сводка по партии — для идемпотентного приёма расходов из других ботов."""
    if not batch_id:
        return None
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT COUNT(*) AS cnt, COALESCE(SUM(amount_byn), 0) AS total
            FROM transactions
            WHERE user_id = ? AND batch_id = ?
            """,
            (workspace_id, batch_id),
        ).fetchone()
    if not row or int(row["cnt"] or 0) == 0:
        return None
    return {"count": int(row["cnt"] or 0), "total": round(float(row["total"] or 0), 2)}


def analyze_and_stage_import(
    workspace_id: int, items: List[Dict[str, Any]], default_currency: str
) -> Tuple[str, List[Dict[str, Any]], List[Dict[str, Any]]]:
    today_str = now_minsk().strftime("%Y-%m-%d")
    new_items: List[Dict[str, Any]] = []
    duplicate_items: List[Dict[str, Any]] = []
    used_db_ids = set()

    # Предварительно конвертируем валюты вне транзакции SQLite
    pre_normalized = []
    for item in items:
        raw_amount = round(float(item["amount"]), 2)
        cur = (item.get("currency") or default_currency).upper().strip()
        tx_date = normalize_tx_date(item.get("tx_date"), today_str)
        norm = dict(item)
        norm["amount"] = raw_amount
        norm["currency"] = cur
        norm["tx_date"] = tx_date
        norm["amount_byn"] = convert_to_byn(raw_amount, cur)
        pre_normalized.append(norm)

    with get_connection() as conn:
        conn.execute("DELETE FROM pending_imports WHERE created_at < datetime('now', '-1 day');")

        for norm in pre_normalized:
            candidates = conn.execute(
                """
                SELECT id, tx_date, amount, currency, description
                FROM transactions
                WHERE user_id = ?
                  AND tx_type = ?
                  AND UPPER(currency) = ?
                  AND ABS(amount - ?) < 0.015
                  AND ABS(julianday(tx_date) - julianday(?)) <= 1.0
                ORDER BY id DESC
                """,
                (workspace_id, norm["tx_type"], norm["currency"], norm["amount"], norm["tx_date"]),
            ).fetchall()

            matched_row = None
            for cand in candidates:
                if cand["id"] not in used_db_ids:
                    matched_row = cand
                    used_db_ids.add(cand["id"])
                    break

            if matched_row is not None:
                norm["matched_desc"] = matched_row["description"]
                duplicate_items.append(norm)
            else:
                new_items.append(norm)

        import_id = uuid.uuid4().hex[:12]
        if new_items:
            conn.execute(
                "INSERT INTO pending_imports (import_id, user_id, payload_json) VALUES (?, ?, ?)",
                (import_id, workspace_id, json.dumps(new_items, ensure_ascii=False)),
            )

    return import_id, new_items, duplicate_items


def commit_staged_import(
    workspace_id: int, import_id: str, author_name: str = ""
) -> Optional[Tuple[str, int, float]]:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT payload_json FROM pending_imports WHERE import_id = ? AND user_id = ?",
            (import_id, workspace_id),
        ).fetchone()
        if not row:
            return None
        conn.execute(
            "DELETE FROM pending_imports WHERE import_id = ? AND user_id = ?",
            (import_id, workspace_id),
        )

    items = json.loads(row["payload_json"])
    batch_id = add_transactions(workspace_id, items, DEFAULT_CURRENCY, author_name=author_name)
    total_byn = sum(float(x.get("amount_byn") or 0.0) for x in items)
    return batch_id, len(items), round(total_byn, 2)


def cancel_staged_import(workspace_id: int, import_id: str) -> bool:
    with get_connection() as conn:
        cur = conn.execute(
            "DELETE FROM pending_imports WHERE import_id = ? AND user_id = ?",
            (import_id, workspace_id),
        )
        return cur.rowcount > 0


def delete_batch(workspace_id: int, batch_id: str) -> int:
    with get_connection() as conn:
        cursor = conn.execute(
            "DELETE FROM transactions WHERE user_id = ? AND batch_id = ?",
            (workspace_id, batch_id),
        )
        return cursor.rowcount


def delete_transaction_by_id(workspace_id: int, tx_id: int) -> bool:
    with get_connection() as conn:
        cursor = conn.execute(
            "DELETE FROM transactions WHERE user_id = ? AND id = ?",
            (workspace_id, tx_id),
        )
        return cursor.rowcount > 0


def _aggregate_period_stats(
    workspace_id: int,
    start_date: str,
    end_date: str,
    include_budgets: bool = True,
) -> Dict[str, Any]:
    """Единый агрегатор статистики по диапазону дат YYYY-MM-DD (включительно)."""
    with get_connection() as conn:
        totals_rows = conn.execute(
            """
            SELECT tx_type, SUM(amount_byn) as total
            FROM transactions
            WHERE user_id = ? AND tx_date >= ? AND tx_date <= ?
            GROUP BY tx_type
            """,
            (workspace_id, start_date, end_date),
        ).fetchall()

        totals = {"income": 0.0, "expense": 0.0}
        for row in totals_rows:
            totals[row["tx_type"]] = round(float(row["total"] or 0.0), 2)

        expense_by_cat = conn.execute(
            """
            SELECT category, SUM(amount_byn) as total, COUNT(*) as cnt
            FROM transactions
            WHERE user_id = ? AND tx_type = 'expense' AND tx_date >= ? AND tx_date <= ?
            GROUP BY category
            ORDER BY total DESC
            """,
            (workspace_id, start_date, end_date),
        ).fetchall()

        income_by_cat = conn.execute(
            """
            SELECT category, SUM(amount_byn) as total, COUNT(*) as cnt
            FROM transactions
            WHERE user_id = ? AND tx_type = 'income' AND tx_date >= ? AND tx_date <= ?
            GROUP BY category
            ORDER BY total DESC
            """,
            (workspace_id, start_date, end_date),
        ).fetchall()

        count_row = conn.execute(
            """
            SELECT COUNT(*) as cnt
            FROM transactions
            WHERE user_id = ? AND tx_date >= ? AND tx_date <= ?
            """,
            (workspace_id, start_date, end_date),
        ).fetchone()

        budgets: Dict[str, float] = {}
        if include_budgets:
            budgets_rows = conn.execute(
                "SELECT category, monthly_limit FROM budgets WHERE user_id = ?",
                (workspace_id,),
            ).fetchall()
            budgets = {r["category"]: float(r["monthly_limit"]) for r in budgets_rows}

    return {
        "start_date": start_date,
        "end_date": end_date,
        "tx_count": int(count_row["cnt"] or 0) if count_row else 0,
        "income": totals["income"],
        "expense": totals["expense"],
        "net": round(totals["income"] - totals["expense"], 2),
        "expense_categories": [dict(r) for r in expense_by_cat],
        "income_categories": [dict(r) for r in income_by_cat],
        "budgets": budgets,
    }


def get_monthly_stats(workspace_id: int, year_month: Optional[str] = None) -> Dict[str, Any]:
    if not year_month:
        year_month = now_minsk().strftime("%Y-%m")

    # Диапазонный фильтр вместо strftime() задействует индекс idx_tx_user_date
    start_date = f"{year_month}-01"
    end_date = f"{year_month}-31"

    stats = _aggregate_period_stats(workspace_id, start_date, end_date)
    stats["year_month"] = year_month
    stats["is_range"] = False
    return stats


def _iso_to_ru(iso_date: str) -> str:
    """YYYY-MM-DD -> DD.MM.YYYY. Без исключений: при мусоре возвращает вход как есть."""
    try:
        year, month, day = str(iso_date)[:10].split("-")
        return f"{day}.{month}.{year}"
    except Exception:
        return str(iso_date)


def get_range_stats(workspace_id: int, start_iso: str, end_iso: str) -> Dict[str, Any]:
    """Статистика за произвольный период (ISO YYYY-MM-DD, включительно).

    Месячные лимиты бюджета к произвольному периоду не применяются.
    """
    start_date = normalize_tx_date(start_iso) or str(start_iso)[:10]
    end_date = normalize_tx_date(end_iso) or str(end_iso)[:10]
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    stats = _aggregate_period_stats(workspace_id, start_date, end_date, include_budgets=False)
    stats["is_range"] = True
    if start_date == end_date:
        stats["label"] = _iso_to_ru(start_date)
    else:
        stats["label"] = f"{_iso_to_ru(start_date)} — {_iso_to_ru(end_date)}"
    return stats


def get_projects_stats(workspace_id: int) -> List[Dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT
                project_tag,
                SUM(CASE WHEN tx_type = 'income' THEN amount_byn ELSE 0 END) as income_byn,
                SUM(CASE WHEN tx_type = 'expense' THEN amount_byn ELSE 0 END) as expense_byn,
                COUNT(*) as tx_count
            FROM transactions
            WHERE user_id = ? AND project_tag != ''
            GROUP BY project_tag
            ORDER BY (income_byn + expense_byn) DESC
            """,
            (workspace_id,),
        ).fetchall()
        return [dict(r) for r in rows]


def get_today_summary(workspace_id: int) -> Dict[str, Any]:
    today_str = now_minsk().strftime("%Y-%m-%d")
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, tx_type, amount, currency, amount_byn, category, description,
                   project_tag, author_name, strftime('%H:%M', created_at) as time_str
            FROM transactions
            WHERE user_id = ? AND tx_date = ?
            ORDER BY id DESC
            """,
            (workspace_id, today_str),
        ).fetchall()

    income = sum(float(r["amount_byn"]) for r in rows if r["tx_type"] == "income")
    expense = sum(float(r["amount_byn"]) for r in rows if r["tx_type"] == "expense")
    return {
        "date": today_str,
        "income": round(income, 2),
        "expense": round(expense, 2),
        "items": [dict(r) for r in rows],
    }


def get_recent_transactions(workspace_id: int, limit: int = 10) -> List[Dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, tx_type, amount, currency, amount_byn, category, description,
                   project_tag, author_name, tx_date
            FROM transactions
            WHERE user_id = ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (workspace_id, limit),
        ).fetchall()
        return [dict(r) for r in rows]


def set_budget(workspace_id: int, category: str, limit_amount: float) -> None:
    with get_connection() as conn:
        if limit_amount <= 0:
            conn.execute(
                "DELETE FROM budgets WHERE user_id = ? AND LOWER(category) = LOWER(?)",
                (workspace_id, category),
            )
        else:
            conn.execute(
                """
                INSERT INTO budgets (user_id, category, monthly_limit)
                VALUES (?, ?, ?)
                ON CONFLICT(user_id, category) DO UPDATE SET monthly_limit = excluded.monthly_limit
                """,
                (workspace_id, category, round(limit_amount, 2)),
            )


def check_budget_warnings(workspace_id: int, categories_affected: List[str]) -> List[str]:
    year_month = now_minsk().strftime("%Y-%m")
    start_date = f"{year_month}-01"
    end_date = f"{year_month}-31"
    warnings = []
    with get_connection() as conn:
        for cat in set(categories_affected):
            b_row = conn.execute(
                "SELECT monthly_limit FROM budgets WHERE user_id = ? AND LOWER(category) = LOWER(?)",
                (workspace_id, cat),
            ).fetchone()
            if not b_row:
                continue
            limit_val = float(b_row["monthly_limit"])
            spent_row = conn.execute(
                """
                SELECT SUM(amount_byn) as spent
                FROM transactions
                WHERE user_id = ? AND tx_type = 'expense'
                  AND LOWER(category) = LOWER(?)
                  AND tx_date >= ? AND tx_date <= ?
                """,
                (workspace_id, cat, start_date, end_date),
            ).fetchone()
            spent_val = float(spent_row["spent"] or 0.0)
            if spent_val > limit_val:
                warnings.append(
                    f"⚠️ Лимит «{cat}» превышен: {spent_val:.2f} из {limit_val:.2f} {DEFAULT_CURRENCY}!"
                )
            elif spent_val >= 0.85 * limit_val:
                warnings.append(
                    f"🔔 Внимание: израсходовано {spent_val:.2f} из {limit_val:.2f} {DEFAULT_CURRENCY} в «{cat}»."
                )
    return warnings


def create_isolated_workspace_backup(workspace_id: int) -> bytes:
    """
    Создает 100% изолированную резервную копию SQLite (.db), содержащую
    ТОЛЬКО транзакции и лимиты текущего workspace_id.
    Чужие данные в этот файл физически не попадают.
    """
    with get_connection() as conn:
        tx_rows = conn.execute(
            """
            SELECT id, user_id, batch_id, created_at, tx_date, tx_type,
                   amount, currency, amount_byn, category, description, project_tag, author_name
            FROM transactions
            WHERE user_id = ?
            ORDER BY id ASC
            """,
            (workspace_id,),
        ).fetchall()

        budget_rows = conn.execute(
            "SELECT user_id, category, monthly_limit FROM budgets WHERE user_id = ?",
            (workspace_id,),
        ).fetchall()

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".db")
    os.close(tmp_fd)
    try:
        b_conn = sqlite3.connect(tmp_path)
        with b_conn:
            b_conn.execute(
                """
                CREATE TABLE transactions (
                    id INTEGER PRIMARY KEY,
                    user_id INTEGER NOT NULL,
                    batch_id TEXT NOT NULL,
                    created_at TIMESTAMP,
                    tx_date TEXT NOT NULL,
                    tx_type TEXT NOT NULL,
                    amount REAL NOT NULL,
                    currency TEXT NOT NULL,
                    amount_byn REAL NOT NULL,
                    category TEXT NOT NULL,
                    description TEXT NOT NULL,
                    project_tag TEXT NOT NULL,
                    author_name TEXT NOT NULL
                )
                """
            )
            b_conn.execute(
                """
                CREATE TABLE budgets (
                    user_id INTEGER NOT NULL,
                    category TEXT NOT NULL,
                    monthly_limit REAL NOT NULL,
                    PRIMARY KEY (user_id, category)
                )
                """
            )
            if tx_rows:
                b_conn.executemany(
                    "INSERT INTO transactions VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    [tuple(r) for r in tx_rows],
                )
            if budget_rows:
                b_conn.executemany(
                    "INSERT INTO budgets VALUES (?, ?, ?)",
                    [tuple(r) for r in budget_rows],
                )
        b_conn.close()

        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def export_transactions_xlsx(
    workspace_id: int,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
) -> bytes:
    is_family = (workspace_id == FAMILY_WORKSPACE_ID)

    where_sql = "user_id = ?"
    params: List[Any] = [workspace_id]
    if date_from:
        where_sql += " AND tx_date >= ?"
        params.append(normalize_tx_date(date_from) or str(date_from)[:10])
    if date_to:
        where_sql += " AND tx_date <= ?"
        params.append(normalize_tx_date(date_to) or str(date_to)[:10])

    with get_connection() as conn:
        rows = conn.execute(
            f"""
            SELECT id, tx_date, tx_type, amount, currency, amount_byn,
                   category, description, project_tag, author_name
            FROM transactions
            WHERE {where_sql}
            ORDER BY tx_date DESC, id DESC
            """,
            params,
        ).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Семейный бюджет" if is_family else "Транзакции"

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)

    income_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    income_font = Font(color="166534", name="Calibri", size=11)
    expense_fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
    expense_font = Font(color="991B1B", name="Calibri", size=11)

    text_font = Font(name="Calibri", size=11, color="1E293B")
    muted_font = Font(name="Calibri", size=11, color="94A3B8")
    bold_text_font = Font(name="Calibri", size=11, color="0F172A", bold=True)

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    if is_family:
        headers = ["Дата", "Тип", "Категория", "Описание", "Кто внёс", "Проект", "Сумма", "Валюта"]
    else:
        headers = ["Дата", "Тип", "Категория", "Описание", "Проект", "Сумма", "Валюта"]

    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    ws.row_dimensions[1].height = 25

    for row_idx, r in enumerate(rows, start=2):
        is_income = r["tx_type"] == "income"
        tx_label = "Доход" if is_income else "Расход"
        tag_str = f"#{r['project_tag']}" if r["project_tag"] else "—"
        cur = r["currency"].upper()
        numeric_amount = round(float(r["amount"]), 2)
        author_str = r["author_name"] or "—"

        if is_family:
            row_data = [
                r["tx_date"], tx_label, r["category"], r["description"],
                author_str, tag_str, numeric_amount, cur
            ]
        else:
            row_data = [
                r["tx_date"], tx_label, r["category"], r["description"],
                tag_str, numeric_amount, cur
            ]

        ws.append(row_data)

        fill_to_use = income_fill if is_income else expense_fill
        type_font = income_font if is_income else expense_font
        cells = [ws.cell(row=row_idx, column=c) for c in range(1, len(row_data) + 1)]

        for cell in cells:
            cell.fill = fill_to_use
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")

        cells[0].font = bold_text_font
        cells[0].alignment = Alignment(horizontal="center", vertical="center")
        cells[1].font = type_font
        cells[1].alignment = Alignment(horizontal="center", vertical="center")
        cells[2].font = text_font
        cells[3].font = text_font

        if is_family:
            cells[4].font = text_font
            cells[4].alignment = Alignment(horizontal="center", vertical="center")
            cells[5].font = muted_font if tag_str == "—" else bold_text_font
            cells[5].alignment = Alignment(horizontal="center", vertical="center")
            cells[6].font = bold_text_font
            cells[6].number_format = "#,##0.00"
            cells[6].alignment = Alignment(horizontal="right", vertical="center")
            cells[7].font = text_font
            cells[7].alignment = Alignment(horizontal="center", vertical="center")
        else:
            cells[4].font = muted_font if tag_str == "—" else bold_text_font
            cells[4].alignment = Alignment(horizontal="center", vertical="center")
            cells[5].font = bold_text_font
            cells[5].number_format = "#,##0.00"
            cells[5].alignment = Alignment(horizontal="right", vertical="center")
            cells[6].font = text_font
            cells[6].alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = f"{cell.value:.2f}" if isinstance(cell.value, float) else str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    # --- Лист «Категории»: полный справочник, включая новые статьи ---
    summary = wb.create_sheet("Категории")
    summary.append(["Категория", "Тип", "Сумма, BYN", "Операций"])

    totals: Dict[Tuple[str, str], Dict[str, float]] = {}
    for r in rows:
        key = (r["category"], r["tx_type"])
        acc = totals.setdefault(key, {"total": 0.0, "cnt": 0.0})
        acc["total"] += float(r["amount_byn"] or 0.0)
        acc["cnt"] += 1.0

    summary_rows: List[Tuple[str, str, float, int]] = []
    for cat in EXPENSE_CATEGORIES:
        acc = totals.get((cat, "expense"))
        summary_rows.append(
            (cat, "Расход", round(acc["total"], 2), int(acc["cnt"]))
            if acc else (cat, "Расход", 0.0, 0)
        )
    for cat in INCOME_CATEGORIES:
        acc = totals.get((cat, "income"))
        summary_rows.append(
            (cat, "Доход", round(acc["total"], 2), int(acc["cnt"]))
            if acc else (cat, "Доход", 0.0, 0)
        )

    # Категории вне справочника (исторические/произвольные) не теряем
    for (cat, tx_type), acc in totals.items():
        known = EXPENSE_CATEGORIES if tx_type == "expense" else INCOME_CATEGORIES
        if cat in known:
            continue
        summary_rows.append(
            (
                cat,
                "Расход" if tx_type == "expense" else "Доход",
                round(acc["total"], 2),
                int(acc["cnt"]),
            )
        )

    for row_data in summary_rows:
        summary.append(list(row_data))

    for col_idx in range(1, 5):
        cell = summary.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row_idx in range(2, len(summary_rows) + 2):
        is_income_row = summary.cell(row=row_idx, column=2).value == "Доход"
        for col_idx in range(1, 5):
            cell = summary.cell(row=row_idx, column=col_idx)
            cell.fill = income_fill if is_income_row else expense_fill
            cell.border = thin_border
            cell.font = income_font if is_income_row else expense_font
        summary.cell(row=row_idx, column=3).number_format = "#,##0.00"

    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 10
    summary.column_dimensions["C"].width = 16
    summary.column_dimensions["D"].width = 12

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()