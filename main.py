import asyncio
import html
import io
import logging
import re
from datetime import date
from typing import Any, Awaitable, Callable, Dict, Tuple

from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputMediaPhoto,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    TelegramObject,
)
from openpyxl import load_workbook

from charts import render_minimalist_stats_chart
from api import start_api, stop_api
from config import (
    BOT_TOKEN,
    DEFAULT_CURRENCY,
    EXPENSE_CATEGORIES,
    FAMILY_CODE,
    FAMILY_WORKSPACE_ID,
    PERSONAL_CODE,
    now_minsk,
)
from currency import convert_to_byn
from db import (
    add_transactions,
    analyze_and_stage_import,
    cancel_staged_import,
    check_budget_warnings,
    clear_user_session,
    commit_staged_import,
    create_isolated_workspace_backup,
    delete_batch,
    delete_transaction_by_id,
    export_transactions_xlsx,
    get_family_member_ids,
    get_monthly_stats,
    get_range_stats,
    get_projects_stats,
    get_recent_transactions,
    get_today_summary,
    get_user_session,
    init_db,
    set_budget,
    set_user_session,
    try_register_family_member,
)
from gemini_parser import (
    ParsedBatch,
    parse_document_statement,
    parse_receipt_image,
    parse_text_message,
    parse_voice_message,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("finance_bot")

router = Router()


def esc(val: Any) -> str:
    return html.escape(str(val or ""))


def main_menu_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [
                KeyboardButton(text="📊 Статистика"),
                KeyboardButton(text="📅 Сегодня"),
            ],
            [
                KeyboardButton(text="🏷 Проекты"),
                KeyboardButton(text="🧾 Последние записи"),
            ],
            [
                KeyboardButton(text="🎯 Лимиты"),
                KeyboardButton(text="📥 Экспорт Excel"),
            ],
            [
                KeyboardButton(text="💾 Бэкап БД"),
                KeyboardButton(text="🔄 Профиль"),
            ],
        ],
        resize_keyboard=True,
    )


class StatsPeriodForm(StatesGroup):
    waiting_period = State()


# --- Ручной период статистики ---------------------------------------------

PERIOD_FORMAT_HINT = "ДД.ММ.ГГГГ - ДД.ММ.ГГГГ"
MAX_PERIOD_DAYS = 3660  # ~10 лет, защита от абсурдных диапазонов

CANCEL_WORDS = {"отмена", "cancel", "❌ отмена", "/cancel"}

MAIN_MENU_LABELS = frozenset(
    {
        "📊 Статистика",
        "📅 Сегодня",
        "🏷 Проекты",
        "🧾 Последние записи",
        "🎯 Лимиты",
        "📥 Экспорт Excel",
        "💾 Бэкап БД",
        "🔄 Профиль",
    }
)


class PeriodInputError(ValueError):
    """Понятная пользователю ошибка разбора периода (бот не падает)."""


_DATE_PART = r"(\d{1,2})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{2,4})"
_SEPARATOR = r"(?:\s*(?:-|—|–)\s*|\s*\.{2,}\s*|\s+по\s+|\s+до\s+|\s+)"
_PERIOD_RE = re.compile(
    r"^\s*" + _DATE_PART + r"(?:" + _SEPARATOR + _DATE_PART + r")?\s*$"
)


def _safe_date(day_s: str, month_s: str, year_s: str) -> date:
    day, month, year = int(day_s), int(month_s), int(year_s)
    if year < 100:
        year += 2000
    try:
        return date(year, month, day)
    except ValueError:
        raise PeriodInputError(
            f"Такой даты не существует: {day:02d}.{month:02d}.{year}. "
            f"Введите в формате {PERIOD_FORMAT_HINT}"
        ) from None


def parse_period_input(text: str) -> Tuple[date, date]:
    """Парсит «ДД.ММ.ГГГГ - ДД.ММ.ГГГГ» или одну дату.

    Бросает только PeriodInputError с готовым текстом для пользователя.
    """
    raw = (text or "").strip()
    if not raw:
        raise PeriodInputError(
            f"Неверный формат даты, введите в формате {PERIOD_FORMAT_HINT}"
        )
    if len(raw) > 64:
        raise PeriodInputError(
            f"Слишком длинный ввод, введите в формате {PERIOD_FORMAT_HINT}"
        )

    match = _PERIOD_RE.match(raw)
    if not match:
        raise PeriodInputError(
            f"Неверный формат даты, введите в формате {PERIOD_FORMAT_HINT}"
        )

    start = _safe_date(*match.group(1, 2, 3))
    end = _safe_date(*match.group(4, 5, 6)) if match.group(4) else start

    if start > end:
        raise PeriodInputError(
            "Дата начала не может быть позже даты окончания. "
            f"Введите период в формате {PERIOD_FORMAT_HINT}"
        )
    if (end - start).days > MAX_PERIOD_DAYS:
        raise PeriodInputError(
            f"Слишком большой период: максимум {MAX_PERIOD_DAYS // 365} лет"
        )
    return start, end


def shift_year_month(ym: str, delta_months: int) -> str:
    year, month = map(int, ym.split("-"))
    total_months = (year * 12 + (month - 1)) + delta_months
    new_year = total_months // 12
    new_month = (total_months % 12) + 1
    return f"{new_year:04d}-{new_month:02d}"


def stats_nav_kb(year_month: str) -> InlineKeyboardMarkup:
    prev_ym = shift_year_month(year_month, -1)
    next_ym = shift_year_month(year_month, 1)
    curr_ym = now_minsk().strftime("%Y-%m")
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text=f"◀️ {prev_ym}", callback_data=f"stats_ym:{prev_ym}"),
                InlineKeyboardButton(text="🔄 Текущий", callback_data=f"stats_ym:{curr_ym}"),
                InlineKeyboardButton(text=f"{next_ym} ▶️", callback_data=f"stats_ym:{next_ym}"),
            ],
            [
                InlineKeyboardButton(
                    text="🗓 Свой период", callback_data="stats_custom"
                )
            ],
        ]
    )


def _sync_build_stats_view(
    workspace_id: int, workspace_mode: str, year_month: str
) -> Tuple[bytes, str, InlineKeyboardMarkup]:
    stats = get_monthly_stats(workspace_id, year_month=year_month)
    income = stats["income"]
    expense = stats["expense"]
    net = stats["net"]
    budgets = {k.lower(): v for k, v in stats["budgets"].items()}

    chart_bytes = render_minimalist_stats_chart(
        year_month=stats["year_month"],
        income_byn=income,
        expense_byn=expense,
        net_byn=net,
        expense_categories=stats["expense_categories"],
        currency=DEFAULT_CURRENCY,
    )

    mode_label = "Семейный бюджет" if workspace_mode == "family" else "Личный счет"
    net_sign = "+" if net >= 0 else ""
    lines = [
        f"📊 <b>{mode_label} • {esc(stats['year_month'])}</b>\n",
        f"🟢 Доходы: <b>+{income:.2f} {DEFAULT_CURRENCY}</b>",
        f"🔴 Расходы: <b>-{expense:.2f} {DEFAULT_CURRENCY}</b>",
        f"💰 Сальдо: <b>{net_sign}{net:.2f} {DEFAULT_CURRENCY}</b>\n",
    ]

    if stats["expense_categories"]:
        lines.append("📉 <b>Детализация расходов:</b>")
        for row in stats["expense_categories"]:
            cat = row["category"]
            tot = float(row["total"])
            pct = (tot / expense * 100.0) if expense > 0 else 0.0
            lim = budgets.get(cat.lower())
            budget_str = f" [Лимит: {tot:.0f}/{lim:.0f}]" if lim is not None else ""
            lines.append(
                f"• {esc(cat)}: <b>{tot:.2f} {DEFAULT_CURRENCY}</b> ({pct:.0f}%){budget_str}"
            )

    if stats["income_categories"]:
        lines.append("\n📈 <b>Источники дохода:</b>")
        for row in stats["income_categories"]:
            lines.append(
                f"• {esc(row['category'])}: <b>+{float(row['total']):.2f} {DEFAULT_CURRENCY}</b>"
            )

    return chart_bytes, "\n".join(lines), stats_nav_kb(stats["year_month"])

def period_result_kb(start_iso: str, end_iso: str) -> InlineKeyboardMarkup:
    curr_ym = now_minsk().strftime("%Y-%m")
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📥 Excel за этот период",
                    callback_data=f"stats_xlsx:{start_iso}:{end_iso}",
                )
            ],
            [
                InlineKeyboardButton(text="🗓 Другой период", callback_data="stats_custom"),
                InlineKeyboardButton(
                    text="🔄 Текущий месяц", callback_data=f"stats_ym:{curr_ym}"
                ),
            ],
        ]
    )


def _sync_build_range_stats_view(
    workspace_id: int, workspace_mode: str, start_iso: str, end_iso: str
) -> Tuple[bytes, str, InlineKeyboardMarkup]:
    stats = get_range_stats(workspace_id, start_iso, end_iso)
    label = stats["label"]
    income = stats["income"]
    expense = stats["expense"]
    net = stats["net"]

    chart_bytes = render_minimalist_stats_chart(
        year_month=label,
        income_byn=income,
        expense_byn=expense,
        net_byn=net,
        expense_categories=stats["expense_categories"],
        currency=DEFAULT_CURRENCY,
    )

    mode_label = "Семейный бюджет" if workspace_mode == "family" else "Личный счет"
    net_sign = "+" if net >= 0 else ""
    lines = [
        f"📊 <b>{mode_label} • период {esc(label)}</b>\n",
        f"🟢 Доходы: <b>+{income:.2f} {DEFAULT_CURRENCY}</b>",
        f"🔴 Расходы: <b>-{expense:.2f} {DEFAULT_CURRENCY}</b>",
        f"💰 Сальдо: <b>{net_sign}{net:.2f} {DEFAULT_CURRENCY}</b>\n",
    ]

    if stats["expense_categories"]:
        lines.append("📉 <b>Детализация расходов:</b>")
        for row in stats["expense_categories"]:
            tot = float(row["total"])
            pct = (tot / expense * 100.0) if expense > 0 else 0.0
            lines.append(
                f"• {esc(row['category'])}: <b>{tot:.2f} {DEFAULT_CURRENCY}</b> ({pct:.0f}%)"
            )

    if stats["income_categories"]:
        lines.append("\n📈 <b>Источники дохода:</b>")
        for row in stats["income_categories"]:
            lines.append(
                f"• {esc(row['category'])}: <b>+{float(row['total']):.2f} {DEFAULT_CURRENCY}</b>"
            )

    lines.append(
        f"\n<i>Операций за период: {stats['tx_count']}. "
        "Лимиты считаются только за календарный месяц.</i>"
    )

    return chart_bytes, "\n".join(lines), period_result_kb(
        stats["start_date"], stats["end_date"]
    )


class WorkspaceAuthMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        user = data.get("event_from_user")
        if not user:
            return await handler(event, data)

        if isinstance(event, Message) and event.text:
            clean_text = event.text.strip()
            if clean_text in ("/switch", "/logout", "🔄 Профиль"):
                clear_user_session(user.id)
                await event.answer(
                    "🔄 <b>Смена режима работы</b>\n\n"
                    "Введи кодовое слово для подключения к нужному счету:",
                    reply_markup=ReplyKeyboardRemove(),
                )
                return None

        session = get_user_session(user.id)
        if not session:
            if isinstance(event, Message):
                entered = (event.text or "").strip().lower()
                display_name = user.first_name or user.username or f"ID {user.id}"

                if entered == FAMILY_CODE:
                    # Постоянная привязка максимум 2 аккаунтов к семейному счету
                    allowed = try_register_family_member(user.id, display_name)
                    if not allowed:
                        await event.answer(
                            "⛔️ <b>Доступ запрещен</b>\n\n"
                            "Семейный счет уже занят! Зарегистрировано максимальное количество участников (2 из 2).\n"
                            "Больше никто не может подключиться по этому кодовому слову."
                        )
                        return None

                    set_user_session(user.id, FAMILY_WORKSPACE_ID, "family", display_name)
                    await event.answer(
                        f"👨‍👩‍👦 <b>Семейный бюджет активирован!</b>\n\n"
                        f"Привет, <b>{esc(display_name)}</b>! Теперь все твои записи идут в общую семейную копилку.\n"
                        f"Просто пиши траты текстом (например: <code>продукты 35</code>), диктуй голосом или отправляй фото чека.",
                        reply_markup=main_menu_kb(),
                    )
                    return None

                if entered == PERSONAL_CODE:
                    set_user_session(user.id, user.id, "personal", display_name)
                    await event.answer(
                        f"👤 <b>Личный финансовый кабинет активирован!</b>\n\n"
                        f"Привет, <b>{esc(display_name)}</b>! Твоя база полностью изолирована и доступна только тебе.\n"
                        f"Для смены режима в любой момент нажми «🔄 Профиль» или введи /switch.",
                        reply_markup=main_menu_kb(),
                    )
                    return None

                await event.answer(
                    "🔐 <b>Доступ защищен</b>\n\n"
                    "Для начала работы введи секретное кодовое слово в ответном сообщении:"
                )
            elif isinstance(event, CallbackQuery):
                await event.answer("Сначала введи кодовое слово в чат!", show_alert=True)
            return None

        # Любая команда или кнопка меню выводит из ручного ввода периода.
        # Иначе пользователь «застрянет» в FSM и следующая трата уйдет в парсер дат.
        if isinstance(event, Message) and event.text:
            fsm_ctx = data.get("state")
            txt = event.text.strip()
            if fsm_ctx is not None and (txt.startswith("/") or txt in MAIN_MENU_LABELS):
                try:
                    await fsm_ctx.clear()
                except Exception:
                    logger.warning("Не удалось сбросить FSM-состояние")

        data["workspace_id"] = int(session["workspace_id"])
        data["workspace_mode"] = session["mode"]
        data["author_name"] = session["display_name"]
        return await handler(event, data)


def undo_batch_kb(batch_id: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🗑 Отменить запись", callback_data=f"undo_batch:{batch_id}"
                )
            ]
        ]
    )


def confirm_import_kb(import_id: str, count: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text=f"✅ Записать уникальные ({count} шт.)",
                    callback_data=f"confirm_imp:{import_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    text="❌ Отменить импорт",
                    callback_data=f"cancel_imp:{import_id}",
                )
            ],
        ]
    )


def make_retry_notifier(status_msg: Message, label: str):
    async def _on_retry(attempt: int, max_attempts: int) -> None:
        try:
            await status_msg.edit_text(
                f"⏳ <i>{esc(label)}\n"
                f"Сервер Google сейчас загружен (503), дожимаю запрос (попытка {attempt}/{max_attempts})...</i>"
            )
        except Exception:
            pass

    return _on_retry


async def notify_family_members(
    bot: Bot, sender_id: int, text: str
) -> None:
    member_ids = get_family_member_ids(exclude_tg_id=sender_id)
    for m_id in member_ids:
        try:
            await bot.send_message(chat_id=m_id, text=text)
        except Exception:
            pass


async def process_parsed_batch(
    message: Message,
    bot: Bot,
    parsed: ParsedBatch,
    workspace_id: int,
    workspace_mode: str,
    author_name: str,
    force_reconcile: bool = False,
) -> None:
    if not parsed.transactions:
        await message.answer(
            "🤔 Не удалось найти финансовых операций. "
            "Пример: <code>бензин 65 #авто</code>, <code>продукты 25</code> или скинь фото чека."
        )
        return

    items_dicts = [tx.model_dump() for tx in parsed.transactions]

    if force_reconcile or parsed.is_bank_statement:
        import_id, new_items, dup_items = await asyncio.to_thread(
            analyze_and_stage_import,
            workspace_id,
            items_dicts,
            DEFAULT_CURRENCY,
        )
        total_found = len(items_dicts)
        dup_count = len(dup_items)
        new_count = len(new_items)

        lines = [
            "🔍 <b>Умная сверка банковской выписки:</b>\n",
            f"• Всего найдено операций: <b>{total_found}</b>",
            f"• Уже есть в базе (пропущено дублей): <b>{dup_count}</b>",
            f"• Новых уникальных операций: <b>{new_count}</b>\n",
        ]

        if dup_items:
            lines.append("<i>Пропущенные совпадения (уже вносились ранее):</i>")
            for d in dup_items[:5]:
                lines.append(
                    f"〰️ <code>{esc(d['tx_date'])}</code> | {d['amount']:.2f} {esc(d['currency'])} — {esc(d['description'])}"
                )
            if len(dup_items) > 5:
                lines.append(f"<i>...и еще {len(dup_items) - 5} дублей</i>\n")
            else:
                lines.append("")

        if not new_items:
            lines.append("✅ <b>Все операции из этой выписки уже записаны в базу!</b>")
            await message.answer("\n".join(lines))
            return

        lines.append("<b>Новые операции к добавлению:</b>")
        for n in new_items[:12]:
            sign = "🟢 +" if n["tx_type"] == "income" else "🔴 -"
            lines.append(
                f"<code>{esc(n['tx_date'])}</code> {sign}<b>{n['amount']:.2f} {esc(n['currency'])}</b> — {esc(n['category'])} <i>({esc(n['description'])})</i>"
            )
        if len(new_items) > 12:
            lines.append(f"<i>...и еще {len(new_items) - 12} новых операций</i>")

        await message.answer(
            "\n".join(lines),
            reply_markup=confirm_import_kb(import_id, new_count),
        )
        return

    batch_id = await asyncio.to_thread(
        add_transactions,
        workspace_id,
        items_dicts,
        DEFAULT_CURRENCY,
        author_name,
    )

    mode_badge = "👨‍👩‍👦 Семейный бюджет" if workspace_mode == "family" else "👤 Личный счет"
    lines = [f"✅ <b>Записано в базу ({mode_badge}):</b>\n"]
    tx_notification_lines = []
    total_expense_byn = 0.0
    total_income_byn = 0.0
    affected_categories = []

    for tx in parsed.transactions:
        sign = "🟢 +" if tx.tx_type == "income" else "🔴 -"
        byn_val = convert_to_byn(tx.amount, tx.currency)
        cur_note = (
            f" (≈ {byn_val:.2f} {DEFAULT_CURRENCY})"
            if tx.currency.upper() != DEFAULT_CURRENCY
            else ""
        )
        tag_note = f"  <code>#{esc(tx.project_tag)}</code>" if tx.project_tag else ""
        item_line = (
            f"{sign}<b>{tx.amount:.2f} {esc(tx.currency)}</b>{cur_note} — {esc(tx.category)} <i>({esc(tx.description)})</i>{tag_note}"
        )
        lines.append(item_line)
        tx_notification_lines.append(item_line)

        if tx.tx_type == "expense":
            total_expense_byn += byn_val
            affected_categories.append(tx.category)
        else:
            total_income_byn += byn_val

    if len(parsed.transactions) > 1:
        lines.append("")
        if total_expense_byn > 0:
            lines.append(f"📉 Итого расход: <b>{total_expense_byn:.2f} {DEFAULT_CURRENCY}</b>")
        if total_income_byn > 0:
            lines.append(f"📈 Итого доход: <b>{total_income_byn:.2f} {DEFAULT_CURRENCY}</b>")

    warnings = check_budget_warnings(workspace_id, affected_categories)
    if warnings:
        lines.append("\n" + "\n".join(esc(w) for w in warnings))

    await message.answer(
        "\n".join(lines),
        reply_markup=undo_batch_kb(batch_id),
    )

    if workspace_mode == "family":
        notify_text = (
            f"👨‍👩‍👦 <b>Обновление семейного бюджета</b> (внёс: <b>{esc(author_name)}</b>):\n\n"
            + "\n".join(tx_notification_lines)
        )
        await notify_family_members(bot=bot, sender_id=message.from_user.id, text=notify_text)


@router.message(CommandStart())
@router.message(Command("help"))
async def cmd_start(
    message: Message,
    state: FSMContext,
    workspace_mode: str,
    author_name: str,
) -> None:
    # /start и /help всегда сбрасывают незавершенный ручной ввод периода
    await state.clear()

    mode_title = (
        "👨‍👩‍👦 <b>Режим: Семейный бюджет (общая база)</b>"
        if workspace_mode == "family"
        else f"👤 <b>Режим: Личный кабинет ({esc(author_name)})</b>"
    )
    text = (
        f"🔥 <b>Финансовый Комбайн v3.3 (Gemini 3.8 Flash)</b>\n"
        f"{mode_title}\n\n"
        "<b>Как пользоваться:</b>\n"
        "• <b>Текст / Голос / Чеки:</b> <code>продукты 45</code>, <code>бензин 65 #авто</code>, <code>+150 аванс</code>\n"
        "• <b>Выписки без дублей:</b> скинь файл выписки (<code>PDF</code>, <code>XLSX</code>, <code>CSV</code>) или скриншот из банка — бот отсеет то, что уже вносили руками.\n"
        "• <b>Смена режима:</b> нажми «🔄 Профиль» или <code>/switch</code>, чтобы переключиться между личным и семейным счетом."
    )
    await message.answer(text, reply_markup=main_menu_kb())


@router.message(Command("stats"))
@router.message(F.text == "📊 Статистика")
async def cmd_stats(message: Message, workspace_id: int, workspace_mode: str) -> None:
    curr_ym = now_minsk().strftime("%Y-%m")
    chart_bytes, caption, kb = await asyncio.to_thread(
        _sync_build_stats_view, workspace_id, workspace_mode, curr_ym
    )
    photo_file = BufferedInputFile(chart_bytes, filename=f"stats_{curr_ym}.png")
    await message.answer_photo(photo=photo_file, caption=caption, reply_markup=kb)


@router.callback_query(F.data.startswith("stats_ym:"))
async def cb_stats_month(
    call: CallbackQuery, workspace_id: int, workspace_mode: str
) -> None:
    target_ym = call.data.split(":", 1)[1]
    chart_bytes, caption, kb = await asyncio.to_thread(
        _sync_build_stats_view, workspace_id, workspace_mode, target_ym
    )
    photo_file = BufferedInputFile(chart_bytes, filename=f"stats_{target_ym}.png")
    try:
        await call.message.edit_media(
            media=InputMediaPhoto(media=photo_file, caption=caption, parse_mode=ParseMode.HTML),
            reply_markup=kb,
        )
    except Exception:
        pass
    await call.answer()

@router.callback_query(F.data == "stats_custom")
async def cb_stats_custom(call: CallbackQuery, state: FSMContext) -> None:
    await state.set_state(StatsPeriodForm.waiting_period)
    hint = (
        "🗓 <b>Ручной период статистики</b>\n\n"
        "Отправь одним сообщением две даты:\n"
        "<code>01.05.2025 - 31.05.2025</code>\n\n"
        "Можно одну дату — посчитаю за этот день: <code>15.05.2025</code>\n"
        "Отмена: /cancel"
    )
    try:
        await call.message.answer(hint, reply_markup=main_menu_kb())
    except Exception:
        logger.exception("Не удалось отправить подсказку ручного периода")
    await call.answer()


@router.message(StatsPeriodForm.waiting_period, F.text)
async def handle_stats_period_input(
    message: Message,
    state: FSMContext,
    workspace_id: int,
    workspace_mode: str,
) -> None:
    raw = (message.text or "").strip()

    if raw.lower() in CANCEL_WORDS or raw in MAIN_MENU_LABELS or raw.startswith("/"):
        await state.clear()
        await message.answer(
            "👌 Ручной ввод периода отменен. Нажми нужную кнопку меню заново.",
            reply_markup=main_menu_kb(),
        )
        return

    try:
        start_date, end_date = parse_period_input(raw)
    except PeriodInputError as exc:
        # Мягко сообщаем об ошибке и ОСТАЕМСЯ в состоянии ожидания
        await message.answer(
            f"⚠️ <b>{esc(exc)}</b>\n\n"
            f"Пример: <code>01.05.2025 - 31.05.2025</code>\n"
            "Отмена: /cancel"
        )
        return
    except Exception:
        logger.exception("Неожиданная ошибка разбора периода: %r", raw)
        await message.answer(
            f"⚠️ Не удалось разобрать даты. Введите в формате {PERIOD_FORMAT_HINT}"
        )
        return

    await state.clear()

    status_msg = await message.answer("📊 <i>Считаю статистику за период...</i>")
    try:
        chart_bytes, caption, kb = await asyncio.to_thread(
            _sync_build_range_stats_view,
            workspace_id,
            workspace_mode,
            start_date.isoformat(),
            end_date.isoformat(),
        )
        await status_msg.delete()
        photo_file = BufferedInputFile(
            chart_bytes,
            filename=f"stats_{start_date:%Y%m%d}_{end_date:%Y%m%d}.png",
        )
        await message.answer_photo(photo=photo_file, caption=caption, reply_markup=kb)
    except Exception:
        logger.exception("Ошибка построения статистики за ручной период")
        try:
            await status_msg.edit_text(
                "❌ Не удалось построить отчет за этот период. Попробуй другой диапазон."
            )
        except Exception:
            await message.answer(
                "❌ Не удалось построить отчет за этот период. Попробуй другой диапазон."
            )


@router.callback_query(F.data.startswith("stats_xlsx:"))
async def cb_stats_range_xlsx(
    call: CallbackQuery, workspace_id: int, workspace_mode: str
) -> None:
    try:
        _, start_iso, end_iso = call.data.split(":", 2)
        start_date = date.fromisoformat(start_iso)
        end_date = date.fromisoformat(end_iso)
    except Exception:
        await call.answer("Период устарел, запросите его заново.", show_alert=True)
        return

    try:
        xlsx_bytes = await asyncio.to_thread(
            export_transactions_xlsx,
            workspace_id,
            start_date.isoformat(),
            end_date.isoformat(),
        )
        prefix = "family" if workspace_mode == "family" else "personal"
        file = BufferedInputFile(
            xlsx_bytes,
            filename=f"{prefix}_{start_date:%Y%m%d}_{end_date:%Y%m%d}.xlsx",
        )
        await call.message.answer_document(
            document=file,
            caption=(
                f"📥 Выгрузка за период <b>{start_date:%d.%m.%Y} — {end_date:%d.%m.%Y}</b>."
            ),
        )
    except Exception:
        logger.exception("Ошибка выгрузки Excel за период")
        await call.answer("Не удалось сформировать Excel за период.", show_alert=True)
    await call.answer()


@router.message(Command("projects"))
@router.message(F.text == "🏷 Проекты")
async def cmd_projects(message: Message, workspace_id: int) -> None:
    rows = get_projects_stats(workspace_id)
    if not rows:
        await message.answer(
            "🏷 <b>Учет по проектам (#хештеги)</b>\n\n"
            "Пока нет записей с тегами. Просто добавляй тег к любой трате или доходу:\n"
            "• <code>химия для салона 85 #детейлинг</code>\n"
            "• <code>+250 химчистка салона #детейлинг</code>\n"
            "• <code>оплата сервера 12 #marketbot</code>"
        )
        return

    lines = ["🏷 <b>Аналитика по проектам (за всё время):</b>\n"]
    for r in rows:
        tag = esc(r["project_tag"])
        inc = float(r["income_byn"] or 0.0)
        exp = float(r["expense_byn"] or 0.0)
        net = inc - exp
        sign = "+" if net >= 0 else ""
        lines.append(
            f"🔹 <b>#{tag}</b> (операций: {r['tx_count']})\n"
            f"   Доход: <code>+{inc:.2f}</code> | Расход: <code>-{exp:.2f}</code> | "
            f"Итог: <b>{sign}{net:.2f} {DEFAULT_CURRENCY}</b>\n"
        )
    await message.answer("\n".join(lines))


@router.message(Command("today"))
@router.message(F.text == "📅 Сегодня")
async def cmd_today(message: Message, workspace_id: int, workspace_mode: str) -> None:
    summary = get_today_summary(workspace_id)
    lines = [
        f"📅 <b>Сводка за сегодня ({esc(summary['date'])}):</b>",
        f"🟢 Доход: <b>{summary['income']:.2f} {DEFAULT_CURRENCY}</b> | 🔴 Расход: <b>{summary['expense']:.2f} {DEFAULT_CURRENCY}</b>\n",
    ]
    if not summary["items"]:
        lines.append("<i>За сегодня записей еще нет.</i>")
    else:
        for item in summary["items"]:
            icon = "🟢 +" if item["tx_type"] == "income" else "🔴 -"
            tag_str = f" <code>#{esc(item['project_tag'])}</code>" if item.get("project_tag") else ""
            author_badge = (
                f" • 👤 {esc(item['author_name'])}"
                if workspace_mode == "family" and item.get("author_name")
                else ""
            )
            lines.append(
                f"<code>{esc(item['time_str'])}</code> {icon}<b>{item['amount']:.2f} {esc(item['currency'])}</b> — "
                f"{esc(item['category'])} <i>({esc(item['description'])}{author_badge})</i>{tag_str}"
            )
    await message.answer("\n".join(lines))


@router.message(Command("last"))
@router.message(F.text == "🧾 Последние записи")
async def cmd_last(message: Message, workspace_id: int, workspace_mode: str) -> None:
    rows = get_recent_transactions(workspace_id, limit=10)
    if not rows:
        await message.answer("📭 В этом профиле пока нет ни одной записи.")
        return

    lines = ["🧾 <b>Последние 10 операций (нажми ❌ для удаления):</b>\n"]
    buttons = []
    for r in rows:
        icon = "🟢 +" if r["tx_type"] == "income" else "🔴 -"
        tag_str = f" <code>#{esc(r['project_tag'])}</code>" if r.get("project_tag") else ""
        author_badge = (
            f" • {esc(r['author_name'])}"
            if workspace_mode == "family" and r.get("author_name")
            else ""
        )
        lines.append(
            f"<b>#{r['id']}</b> [<code>{esc(r['tx_date'])}</code>] {icon}<b>{r['amount']:.2f} {esc(r['currency'])}</b> — "
            f"{esc(r['category'])} <i>({esc(r['description'])}{author_badge})</i>{tag_str}"
        )
        buttons.append(
            InlineKeyboardButton(
                text=f"❌ #{r['id']} ({r['amount']:.0f})",
                callback_data=f"del_tx:{r['id']}",
            )
        )

    kb_rows = [buttons[i : i + 2] for i in range(0, len(buttons), 2)]
    await message.answer(
        "\n".join(lines),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_rows),
    )


@router.message(Command("budget"))
@router.message(F.text == "🎯 Лимиты")
async def cmd_budget(message: Message, workspace_id: int) -> None:
    text = (message.text or "").strip()
    if text == "🎯 Лимиты" or text == "/budget":
        stats = get_monthly_stats(workspace_id)
        budgets = stats["budgets"]
        spent_map = {
            r["category"].lower(): float(r["total"]) for r in stats["expense_categories"]
        }

        lines = [
            "🎯 <b>Месячные лимиты по категориям:</b>\n",
            "Чтобы задать лимит, напиши:\n<code>/budget Продукты 450</code> (или <code>0</code> для удаления)\n",
            f"<b>Категории:</b> {', '.join(EXPENSE_CATEGORIES)}\n",
        ]
        if budgets:
            lines.append("<b>Установленные лимиты:</b>")
            for cat, limit_val in budgets.items():
                spent = spent_map.get(cat.lower(), 0.0)
                pct = (spent / limit_val * 100.0) if limit_val > 0 else 0.0
                status_icon = "⚠️" if spent > limit_val else "✅"
                lines.append(
                    f"{status_icon} {esc(cat)}: <b>{spent:.2f} / {limit_val:.2f} {DEFAULT_CURRENCY}</b> ({pct:.0f}%)"
                )
        else:
            lines.append("<i>Лимиты пока не заданы.</i>")

        await message.answer("\n".join(lines))
        return

    raw_args = text[len("/budget") :].strip()
    tokens = raw_args.rsplit(maxsplit=1)
    if len(tokens) != 2:
        await message.answer("Формат: <code>/budget Категория Сумма</code>")
        return

    raw_cat, amount_str = tokens[0].strip(), tokens[1].replace(",", ".").strip()
    canonical_map = {c.lower(): c for c in EXPENSE_CATEGORIES}
    category_name = canonical_map.get(raw_cat.lower(), raw_cat)

    try:
        limit_val = float(amount_str)
        if limit_val < 0:
            raise ValueError
    except ValueError:
        await message.answer("Сумма лимита должна быть числом (или 0 для удаления лимита).")
        return

    set_budget(workspace_id, category_name, limit_val)
    if limit_val == 0:
        await message.answer(f"🗑 Лимит для «<b>{esc(category_name)}</b>» удален.")
    else:
        await message.answer(
            f"✅ Лимит для «<b>{esc(category_name)}</b>» установлен на <b>{limit_val:.2f} {DEFAULT_CURRENCY}</b>."
        )


@router.message(Command("export"))
@router.message(F.text == "📥 Экспорт Excel")
async def cmd_export(message: Message, workspace_id: int, workspace_mode: str) -> None:
    xlsx_bytes = await asyncio.to_thread(export_transactions_xlsx, workspace_id)
    fname = "family_finance.xlsx" if workspace_mode == "family" else "personal_finance.xlsx"
    file = BufferedInputFile(xlsx_bytes, filename=fname)
    await message.answer_document(
        document=file,
        caption="📥 Выгрузка в формате Excel (.xlsx).",
    )


@router.message(Command("backup"))
@router.message(F.text == "💾 Бэкап БД")
async def cmd_backup(message: Message, workspace_id: int, workspace_mode: str) -> None:
    stamp = now_minsk().strftime("%Y-%m-%d_%H-%M")
    prefix = "family" if workspace_mode == "family" else "personal"
    backup_bytes = await asyncio.to_thread(create_isolated_workspace_backup, workspace_id)
    db_file = BufferedInputFile(backup_bytes, filename=f"{prefix}_backup_{stamp}.db")
    await message.answer_document(
        document=db_file,
        caption=f"💾 Изолированная резервная копия вашего профиля (SQLite) от <code>{stamp}</code>.",
    )


@router.callback_query(F.data.startswith("confirm_imp:"))
async def cb_confirm_import(
    call: CallbackQuery, bot: Bot, workspace_id: int, workspace_mode: str, author_name: str
) -> None:
    import_id = call.data.split(":", 1)[1]
    res = await asyncio.to_thread(commit_staged_import, workspace_id, import_id, author_name)
    if not res:
        await call.answer("Импорт уже был обработан или устарел.", show_alert=True)
        return
    batch_id, count, total_byn = res
    await call.message.edit_text(
        f"✅ <b>Импорт выписки завершен!</b>\n"
        f"Записано уникальных операций: <b>{count}</b> (на общую сумму <b>{total_byn:.2f} {DEFAULT_CURRENCY}</b>).",
        reply_markup=undo_batch_kb(batch_id),
    )
    if workspace_mode == "family":
        await notify_family_members(
            bot=bot,
            sender_id=call.from_user.id,
            text=f"👨‍👩‍👦 <b>{esc(author_name)}</b> импортировал(а) выписку: <b>{count}</b> операций на сумму <b>{total_byn:.2f} {DEFAULT_CURRENCY}</b>.",
        )
    await call.answer()


@router.callback_query(F.data.startswith("cancel_imp:"))
async def cb_cancel_import(call: CallbackQuery, workspace_id: int) -> None:
    import_id = call.data.split(":", 1)[1]
    cancel_staged_import(workspace_id, import_id)
    await call.message.edit_text("🚫 <i>Импорт выписки отменен. Ничего не записано.</i>")
    await call.answer()


@router.callback_query(F.data.startswith("undo_batch:"))
async def cb_undo_batch(
    call: CallbackQuery, bot: Bot, workspace_id: int, workspace_mode: str, author_name: str
) -> None:
    batch_id = call.data.split(":", 1)[1]
    deleted_count = delete_batch(workspace_id, batch_id)
    if deleted_count > 0:
        await call.message.edit_text(
            f"🗑 <i>Запись отменена (удалено позиций: {deleted_count}).</i>"
        )
        if workspace_mode == "family":
            await notify_family_members(
                bot=bot,
                sender_id=call.from_user.id,
                text=f"🗑 <b>{esc(author_name)}</b> отменил(а) свою последнюю запись ({deleted_count} поз.).",
            )
    else:
        await call.answer("Запись уже была удалена.", show_alert=True)
    await call.answer()


@router.callback_query(F.data.startswith("del_tx:"))
async def cb_del_tx(
    call: CallbackQuery, workspace_id: int, workspace_mode: str
) -> None:
    tx_id = int(call.data.split(":", 1)[1])
    ok = delete_transaction_by_id(workspace_id, tx_id)
    if ok:
        await call.answer(f"Транзакция #{tx_id} удалена!")
        rows = get_recent_transactions(workspace_id, limit=10)
        if not rows:
            await call.message.edit_text("📭 Все записи удалены.")
            return
        lines = ["🧾 <b>Последние 10 операций (нажми ❌ для удаления):</b>\n"]
        buttons = []
        for r in rows:
            icon = "🟢 +" if r["tx_type"] == "income" else "🔴 -"
            author_badge = (
                f" • {esc(r['author_name'])}"
                if workspace_mode == "family" and r.get("author_name")
                else ""
            )
            lines.append(
                f"<b>#{r['id']}</b> [<code>{esc(r['tx_date'])}</code>] {icon}<b>{r['amount']:.2f} {esc(r['currency'])}</b> — "
                f"{esc(r['category'])} <i>({esc(r['description'])}{author_badge})</i>"
            )
            buttons.append(
                InlineKeyboardButton(
                    text=f"❌ #{r['id']} ({r['amount']:.0f})",
                    callback_data=f"del_tx:{r['id']}",
                )
            )
        kb_rows = [buttons[i : i + 2] for i in range(0, len(buttons), 2)]
        await call.message.edit_text(
            "\n".join(lines),
            reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_rows),
        )
    else:
        await call.answer("Транзакция уже удалена.", show_alert=True)


def _extract_xlsx_text(raw_bytes: bytes) -> str:
    wb = load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
    sheet = wb.active
    extracted_rows = []
    for row in sheet.iter_rows(values_only=True):
        cleaned = [str(cell).strip() for cell in row if cell is not None and str(cell).strip()]
        if cleaned:
            extracted_rows.append(" | ".join(cleaned))
        if len(extracted_rows) >= 2500:
            break
    wb.close()
    return "\n".join(extracted_rows)


@router.message(F.document)
async def handle_document_statement(
    message: Message,
    bot: Bot,
    workspace_id: int,
    workspace_mode: str,
    author_name: str,
) -> None:
    doc = message.document
    if doc.file_size and doc.file_size > 15 * 1024 * 1024:
        await message.answer("⚠️ Файл слишком большой (максимум 15 МБ).")
        return

    status_msg = await message.answer("📂 <i>Анализирую банковскую выписку и сверяю дубли...</i>")
    try:
        filename = (doc.file_name or "").lower()
        buf = io.BytesIO()
        await bot.download(doc, destination=buf)
        raw_bytes = buf.getvalue()

        text_fallback = ""
        mime_type = doc.mime_type or "application/pdf"

        if filename.endswith(".xlsx"):
            text_fallback = await asyncio.to_thread(_extract_xlsx_text, raw_bytes)
        elif filename.endswith((".csv", ".txt")):
            for enc in ("utf-8", "cp1251", "latin-1"):
                try:
                    text_fallback = raw_bytes.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue

        parsed = await parse_document_statement(
            file_bytes=raw_bytes,
            mime_type=mime_type,
            text_fallback=text_fallback,
            caption=message.caption or "",
            on_retry=make_retry_notifier(status_msg, "Анализирую банковскую выписку..."),
            actor_id=message.from_user.id,
            actor_name=author_name,
        )
        await status_msg.delete()
        await process_parsed_batch(
            message,
            bot,
            parsed,
            workspace_id=workspace_id,
            workspace_mode=workspace_mode,
            author_name=author_name,
            force_reconcile=True,
        )
    except Exception as exc:
        logger.exception("Ошибка при разборе файла выписки")
        await status_msg.edit_text(f"❌ Ошибка чтения выписки: <code>{esc(exc)}</code>")


@router.message(F.photo)
async def handle_receipt_photo(
    message: Message,
    bot: Bot,
    workspace_id: int,
    workspace_mode: str,
    author_name: str,
) -> None:
    status_msg = await message.answer("⏳ <i>Распознаю чек через Gemini 3.8 Flash...</i>")
    try:
        photo = message.photo[-1]
        buf = io.BytesIO()
        await bot.download(photo, destination=buf)
        parsed = await parse_receipt_image(
            buf.getvalue(),
            caption=message.caption or "",
            on_retry=make_retry_notifier(status_msg, "Распознаю фото чека..."),
            actor_id=message.from_user.id,
            actor_name=author_name,
        )
        await status_msg.delete()
        await process_parsed_batch(
            message,
            bot,
            parsed,
            workspace_id=workspace_id,
            workspace_mode=workspace_mode,
            author_name=author_name,
        )
    except Exception as exc:
        logger.exception("Ошибка при разборе фото")
        await status_msg.edit_text(f"❌ Ошибка распознавания: <code>{esc(exc)}</code>")


@router.message(F.voice)
async def handle_voice(
    message: Message,
    bot: Bot,
    workspace_id: int,
    workspace_mode: str,
    author_name: str,
) -> None:
    status_msg = await message.answer("🎧 <i>Распознаю голосовое через Gemini 3.8 Flash...</i>")
    try:
        buf = io.BytesIO()
        await bot.download(message.voice, destination=buf)
        parsed = await parse_voice_message(
            buf.getvalue(),
            on_retry=make_retry_notifier(status_msg, "Распознаю голосовое сообщение..."),
            actor_id=message.from_user.id,
            actor_name=author_name,
        )
        await status_msg.delete()
        await process_parsed_batch(
            message,
            bot,
            parsed,
            workspace_id=workspace_id,
            workspace_mode=workspace_mode,
            author_name=author_name,
        )
    except Exception as exc:
        logger.exception("Ошибка при разборе голосового сообщения")
        await status_msg.edit_text(f"❌ Ошибка разбора аудио: <code>{esc(exc)}</code>")


@router.message(F.text)
async def handle_text_transaction(
    message: Message,
    bot: Bot,
    workspace_id: int,
    workspace_mode: str,
    author_name: str,
) -> None:
    status_msg = await message.answer("⚡ <i>Считаю...</i>")
    try:
        parsed = await parse_text_message(
            message.text,
            on_retry=make_retry_notifier(status_msg, "Разбираю текст..."),
            actor_id=message.from_user.id,
            actor_name=author_name,
        )
        await status_msg.delete()
        await process_parsed_batch(
            message,
            bot,
            parsed,
            workspace_id=workspace_id,
            workspace_mode=workspace_mode,
            author_name=author_name,
        )
    except Exception as exc:
        logger.exception("Ошибка при разборе текста")
        await status_msg.edit_text(f"❌ Ошибка обращения к ИИ: <code>{esc(exc)}</code>")


async def main() -> None:
    init_db()
    bot = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    dp = Dispatcher()
    dp.message.middleware(WorkspaceAuthMiddleware())
    dp.callback_query.middleware(WorkspaceAuthMiddleware())
    dp.include_router(router)

    api_runner = None
    try:
        api_runner = await start_api()
    except Exception:
        logger.exception("Не удалось запустить Finance API, бот продолжит работу без него")
    logger.info("Финансовый комбайн v3.3 (Gemini 3.8 Flash) запущен!")
    try:
        await bot.delete_webhook(drop_pending_updates=True)
        await dp.start_polling(bot)
    finally:
        await stop_api(api_runner)
        await bot.session.close()


if __name__ == "__main__":
    asyncio.run(main())