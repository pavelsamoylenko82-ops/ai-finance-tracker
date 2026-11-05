"""HTTP-приёмник расходов для связи с ботом-напоминатором.

Запускается рядом с polling внутри того же процесса и слушает внутренний
Docker-порт (по умолчанию 8081). Наружу порт публиковать не обязательно.

Пример запроса:
POST /api/expenses
X-API-Token: <FINANCE_API_TOKEN>
{
  "source": "family_reminder_bot",
  "author": "Паша (Сын)",
  "external_id": "shop-abc123",
  "items": [
    {"description": "Молоко", "amount": 2.5, "currency": "BYN",
     "category": "Продукты", "tx_type": "expense"}
  ]
}
"""

from __future__ import annotations

import asyncio
import hmac
import logging
import os
import re
from typing import Any, Dict, List, Optional

from aiohttp import web

from config import DEFAULT_CURRENCY, FAMILY_WORKSPACE_ID
from db import add_transactions, get_batch_stats

logger = logging.getLogger(__name__)

# SQLite-проверка и вставка должны быть одной критической секцией, иначе два
# одновременных запроса с одним external_id могут оба пройти проверку.
_ingest_lock = asyncio.Lock()

API_HOST = (os.getenv("FINANCE_API_HOST") or "0.0.0.0").strip()
API_PORT = int(os.getenv("FINANCE_API_PORT", "8081") or 8081)
API_TOKEN = (os.getenv("FINANCE_API_TOKEN") or "").strip()
MAX_ITEMS = 100
_EXTERNAL_ID_RE = re.compile(r"^[A-Za-z0-9_.:-]{1,80}$")
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


class ApiError(Exception):
    def __init__(self, message: str, status: int = 400) -> None:
        super().__init__(message)
        self.status = status


def _authorized(request: web.Request) -> bool:
    if not API_TOKEN:
        # Внутренний Docker-порт без публикации наружу. Если токен задан —
        # проверка обязательна.
        return True
    supplied = request.headers.get("X-API-Token", "")
    return bool(supplied) and hmac.compare_digest(supplied, API_TOKEN)


def _parse_payload(payload: Any) -> tuple[List[Dict[str, Any]], str, str]:
    if not isinstance(payload, dict):
        raise ApiError("JSON должен быть объектом")
    raw_items = payload.get("items")
    if not isinstance(raw_items, list) or not raw_items:
        raise ApiError("Поле items должно быть непустым массивом")
    if len(raw_items) > MAX_ITEMS:
        raise ApiError(f"Слишком много позиций (максимум {MAX_ITEMS})")

    author = str(payload.get("author") or "Бот-напоминатор").strip()[:120]
    external_id = str(payload.get("external_id") or "").strip()
    if external_id and not _EXTERNAL_ID_RE.fullmatch(external_id):
        raise ApiError("Некорректный external_id")

    items: List[Dict[str, Any]] = []
    for index, raw in enumerate(raw_items, start=1):
        if not isinstance(raw, dict):
            raise ApiError(f"Позиция #{index}: ожидается объект")
        description = str(raw.get("description") or "").strip()
        if not description:
            raise ApiError(f"Позиция #{index}: пустое description")
        try:
            amount = round(float(raw.get("amount")), 2)
        except (TypeError, ValueError):
            raise ApiError(f"Позиция #{index}: некорректная сумма") from None
        if amount <= 0 or amount > 10_000_000:
            raise ApiError(f"Позиция #{index}: сумма должна быть больше нуля")
        currency = str(raw.get("currency") or DEFAULT_CURRENCY).upper().strip()[:8]
        category = str(raw.get("category") or "Продукты").strip()[:80]
        tx_date = str(raw.get("tx_date") or "").strip()
        if tx_date and not _DATE_RE.fullmatch(tx_date):
            raise ApiError(f"Позиция #{index}: дата должна быть YYYY-MM-DD")
        items.append(
            {
                "tx_type": "expense",
                "amount": amount,
                "currency": currency,
                "category": category,
                "description": description[:200],
                "tx_date": tx_date,
            }
        )
    return items, external_id, author


async def handle_health(request: web.Request) -> web.Response:
    return web.json_response(
        {
            "ok": True,
            "service": "finance_bot",
            "api_version": 1,
            "auth_required": bool(API_TOKEN),
        }
    )


async def handle_expenses(request: web.Request) -> web.Response:
    if not _authorized(request):
        raise web.HTTPUnauthorized(text="invalid API token")
    try:
        payload = await request.json()
    except ValueError:
        return web.json_response({"ok": False, "error": "invalid JSON"}, status=400)

    try:
        items, external_id, author = _parse_payload(payload)
    except ApiError as exc:
        return web.json_response({"ok": False, "error": str(exc)}, status=exc.status)

    async with _ingest_lock:
        if external_id:
            existing = await asyncio.to_thread(
                get_batch_stats, FAMILY_WORKSPACE_ID, external_id
            )
            if existing:
                return web.json_response(
                    {
                        "ok": True,
                        "duplicate": True,
                        "batch_id": external_id,
                        "count": existing["count"],
                        "total": existing["total"],
                    }
                )

        try:
            batch_id = await asyncio.to_thread(
                add_transactions,
                FAMILY_WORKSPACE_ID,
                items,
                DEFAULT_CURRENCY,
                author,
                external_id or None,
            )
        except Exception:
            logger.exception("Не удалось записать расходы из API")
            return web.json_response(
                {"ok": False, "error": "database error"},
                status=500,
            )

    total = round(sum(float(item["amount"]) for item in items), 2)
    logger.info(
        "API принял расходы: %s поз. на %s, author=%s, batch=%s",
        len(items),
        total,
        author,
        batch_id,
    )
    return web.json_response(
        {
            "ok": True,
            "duplicate": False,
            "batch_id": batch_id,
            "count": len(items),
            "total": total,
        }
    )


def create_app() -> web.Application:
    app = web.Application(client_max_size=1024 * 1024)
    app.router.add_get("/health", handle_health)
    app.router.add_post("/api/expenses", handle_expenses)
    return app


async def start_api() -> Optional[web.AppRunner]:
    if not API_TOKEN:
        logger.warning(
            "FINANCE_API_TOKEN не задан: API доступен без токена. "
            "Не публикуй порт %s наружу.",
            API_PORT,
        )
    runner = web.AppRunner(create_app(), access_log=None)
    await runner.setup()
    site = web.TCPSite(runner, API_HOST, API_PORT)
    try:
        await site.start()
    except OSError as exc:
        logger.error(
            "Не удалось запустить Finance API на %s:%s: %s",
            API_HOST,
            API_PORT,
            exc,
        )
        await runner.cleanup()
        return None
    logger.info(
        "Finance API слушает http://%s:%s/api/expenses (token=%s)",
        API_HOST,
        API_PORT,
        "on" if API_TOKEN else "off",
    )
    return runner


async def stop_api(runner: Optional[web.AppRunner]) -> None:
    if runner is not None:
        await runner.cleanup()
        logger.info("Finance API остановлен")
