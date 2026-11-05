import asyncio
import base64
import io
import logging
import os
import re
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Literal, Optional, Union

import httpx
from PIL import Image
from pydantic import BaseModel, Field

from config import (
    CHILD_CATEGORIES,
    CURRENCY_ALIASES,
    DEFAULT_CURRENCY,
    EXPENSE_CATEGORIES,
    EXPENSE_KEYWORDS,
    GEMINI_API_KEY,
    GEMINI_MODEL,
    INCOME_CATEGORIES,
    INCOME_KEYWORDS,
    INCOME_TRIGGER_WORDS,
    now_minsk,
)

logger = logging.getLogger("finance_bot.gemini")


def _load_fresh_env() -> Dict[str, str]:
    values: Dict[str, str] = {}
    for env_path in (Path("/app/.env"), Path(".env")):
        if env_path.is_file():
            try:
                for line in env_path.read_text(encoding="utf-8").splitlines():
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    values[k.strip()] = v.strip().strip("'").strip('"')
                break
            except Exception:
                pass
    return values


_ENV_DISK = _load_fresh_env()

API_KEY = (
    _ENV_DISK.get("OPENAI_API_KEY")
    or _ENV_DISK.get("GEMINI_API_KEY")
    or os.getenv("OPENAI_API_KEY")
    or GEMINI_API_KEY
    or "sk-fc-brUkPKHb4iXRDA7N4mrkVtD1I-DpCQt6-lAJkGGjflA"
).strip()

BASE_URL = (
    _ENV_DISK.get("OPENAI_BASE_URL")
    or _ENV_DISK.get("GEMINI_BASE_URL")
    or os.getenv("OPENAI_BASE_URL")
    or "https://freecodex.fun/v1"
).rstrip("/")

ACTIVE_MODEL = (
    _ENV_DISK.get("GEMINI_MODEL")
    or GEMINI_MODEL
    or "gemini-3.8-flash"
).strip()

http_client = httpx.AsyncClient(
    timeout=httpx.Timeout(90.0, connect=15.0),
    headers={
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
        "User-Agent": "OpenAI/Python 1.54.0",
    },
)

client = http_client

RetryCallback = Optional[Callable[[int, int], Awaitable[None]]]


class ParsedTransaction(BaseModel):
    tx_type: Literal["income", "expense"] = Field(
        description="Тип операции: 'income' для доходов/поступлений, 'expense' для трат/покупок."
    )
    amount: float = Field(
        description="Итоговая сумма операции (положительное число)."
    )
    currency: Literal["BYN", "RUB", "USD", "EUR"] = Field(
        default="BYN",
        description=f"Код валюты: BYN, RUB, USD или EUR. По умолчанию {DEFAULT_CURRENCY}.",
    )
    category: str = Field(
        description="Наиболее подходящая категория строго из разрешенного списка."
    )
    description: str = Field(
        description="Краткое и понятное название товара, магазина, маршрута или источника дохода."
    )
    project_tag: str = Field(
        default="",
        description="Тег проекта без символа # в нижнем регистре, если указан #хештег. Иначе пустая строка.",
    )
    tx_date: str = Field(
        default="",
        description="Дата операции YYYY-MM-DD из выписки или чека, либо пустая строка.",
    )


class ParsedBatch(BaseModel):
    is_bank_statement: bool = Field(
        default=False,
        description="True только для банковской выписки со списком операций за период. False для обычного чека или сообщения.",
    )
    transactions: List[ParsedTransaction] = Field(
        description="Список всех распознанных транзакций."
    )


def _word_matches(keyword: str, text: str) -> bool:
    kw = keyword.lower().strip()
    if not kw:
        return False
    if len(kw) <= 3 and " " not in kw:
        pattern = rf"(?<![а-яёa-z0-9]){re.escape(kw)}(?![а-яёa-z0-9])"
    else:
        pattern = rf"(?<![а-яёa-z0-9]){re.escape(kw)}"
    return re.search(pattern, text) is not None


def _classify_category_raw(desc_lower: str, tx_type: str, project_tag: str) -> str:
    check_str = f"{desc_lower} {project_tag}".strip()
    if tx_type == "income":
        if project_tag in ("детейлинг", "химчистка", "гараж"):
            return "Бизнес (Детейлинг)"
        if project_tag in ("marketbot", "бот", "it", "ит"):
            return "Фриланс и IT-проекты"
        for cat, words in INCOME_KEYWORDS.items():
            for w in words:
                if _word_matches(w, check_str):
                    return cat
        return "Прочие доходы"
    else:
        for w in EXPENSE_KEYWORDS.get("Паша", ["паша", "паше", "пашу", "паши", "пашей"]):
            if _word_matches(w, check_str):
                return "Паша"
        for w in EXPENSE_KEYWORDS.get("Вера", ["вера", "вере", "веру", "верой", "веры", "верочка"]):
            if _word_matches(w, check_str):
                return "Вера"
        if project_tag in ("детейлинг", "химчистка", "marketbot", "бизнес"):
            return "Бизнес и расходники"
        if project_tag in ("авто", "машина", "audi", "ауди"):
            return "Авто"
        for cat, words in EXPENSE_KEYWORDS.items():
            for w in words:
                if " " in w and _word_matches(w, check_str):
                    return cat
        for cat, words in EXPENSE_KEYWORDS.items():
            for w in words:
                if _word_matches(w, check_str):
                    return cat
        return "Прочие расходы"


def _classify_category(
    desc_lower: str, tx_type: str, project_tag: str, actor_name: str = ""
) -> str:
    return _classify_category_raw(desc_lower, tx_type, project_tag)


def try_fast_local_parse(
    text: str, allow_fallback_mode: bool = False, actor_name: str = ""
) -> Optional[ParsedBatch]:
    raw = (text or "").strip()
    if not raw or len(raw) > 600:
        return None

    if re.search(r"\b202\d[-./]\d{2}[-./]\d{2}\b|\b\d{2}\.\d{2}\.202\d\b", raw):
        if not allow_fallback_mode:
            return None

    normalized = re.sub(r"(?<=\d),(?=\d)", ".", raw)
    chunks = [c.strip() for c in re.split(r"[,;\n]+", normalized) if c.strip()]
    if not chunks:
        return None

    parsed_items: List[ParsedTransaction] = []

    for chunk in chunks:
        tag_match = re.search(r"#([a-zA-Zа-яА-ЯёЁ0-9_]+)", chunk)
        project_tag = tag_match.group(1).lower() if tag_match else ""
        clean_chunk = re.sub(r"#[a-zA-Zа-яА-ЯёЁ0-9_]+", " ", chunk).strip()

        num_matches = list(
            re.finditer(
                r"(?<![a-zA-Zа-яА-ЯёЁ0-9])([+-]?\d+(?:\.\d{1,2})?)(?![a-zA-Zа-яА-ЯёЁ0-9])",
                clean_chunk,
            )
        )
        if len(num_matches) != 1:
            if not allow_fallback_mode:
                return None
            if not num_matches:
                continue

        match = num_matches[0]
        num_str = match.group(1)
        has_plus = num_str.startswith("+")
        amount = abs(float(num_str))
        if amount <= 0:
            continue

        before_num = clean_chunk[: match.start()].strip()
        after_num = clean_chunk[match.end() :].strip()
        currency = DEFAULT_CURRENCY
        combined_words = f"{before_num} {after_num}".split()
        remaining_words = []
        currency_found = False

        for w in combined_words:
            w_clean = w.lower().strip(".,-—:()")
            if not currency_found and w_clean in CURRENCY_ALIASES:
                currency = CURRENCY_ALIASES[w_clean]
                currency_found = True
            elif w_clean in ("-", "—", "на", "за", "в", "по"):
                continue
            else:
                remaining_words.append(w)

        desc_raw = " ".join(remaining_words).strip(" .,-—:+")
        if not desc_raw:
            if not allow_fallback_mode:
                return None
            desc_raw = "Операция"

        desc_lower = desc_raw.lower()
        is_income = (
            has_plus
            or any(w in desc_lower.split() for w in INCOME_TRIGGER_WORDS)
            or any(_word_matches(k, desc_lower) for k in ("пришло", "заработал", "скинули", "вернули"))
        )
        tx_type: Literal["income", "expense"] = "income" if is_income else "expense"
        category = _classify_category(desc_lower, tx_type, project_tag, actor_name)

        if category in ("Прочие расходы", "Прочие доходы") and not allow_fallback_mode:
            return None

        description_clean = (
            desc_raw[0].upper() + desc_raw[1:] if len(desc_raw) > 1 else desc_raw.upper()
        )

        parsed_items.append(
            ParsedTransaction(
                tx_type=tx_type,
                amount=round(amount, 2),
                currency=currency,
                category=category,
                description=description_clean,
                project_tag=project_tag,
                tx_date="",
            )
        )

    if not parsed_items:
        return None

    return ParsedBatch(is_bank_statement=False, transactions=parsed_items)


def _optimize_image_bytes(image_bytes: bytes) -> bytes:
    try:
        with Image.open(io.BytesIO(image_bytes)) as img:
            img = img.convert("RGB")
            img.thumbnail((1024, 1024))
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=78, optimize=True)
            return out.getvalue()
    except Exception:
        return image_bytes


def _build_system_prompt(
    actor_id: Optional[int] = None, actor_name: str = ""
) -> str:
    today_iso = now_minsk().strftime("%Y-%m-%d")

    rules = [
        "1. «Топливо» — ИСКЛЮЧИТЕЛЬНО покупка горючего: бензин, дизель, газ, АЗС, заправка, литры, колонка. "
        "Любые другие авто-расходы (ремонт, запчасти, шины, мойка, страховка, техосмотр, парковка, штраф, "
        "аксессуары) — строго «Авто». Никогда не смешивай «Топливо» и «Авто».",
        "2. «Буфет» — перекусы и мелочи на работе: буфет, столовая на работе, кофе/снек/булочка на смене. "
        "Кафе, рестораны, доставка еды, фастфуд на улице — «Кафе и фастфуд». Не путай эти две категории.",
        "3. «Паша» — отдельная категория расходов. Любое упоминание «Паша/Паше/Пашу/Паши/Пашей» ВСЕГДА и безусловно относится строго к категории «Паша».",
        "4. «Вера» — отдельная категория расходов. Любое упоминание «Вера/Вере/Веру/Верой/Веры/Верочка» ВСЕГДА и безусловно относится строго к категории «Вера».",
        "5. Категория всегда строго из списка выше. Не придумывай новые категории и не добавляй пояснений.",
    ]

    schema = (
        '{"is_bank_statement": false, "transactions": [{'
        '"tx_type": "expense", "amount": 10.0, '
        '"currency": "' + DEFAULT_CURRENCY + '", '
        '"category": "Продукты", '
        '"description": "Название", "project_tag": "", "tx_date": ""}]}'
    )

    parts = [
        "Ты финансовый парсер Telegram-бота семейного бюджета. Сегодня: %s. Валюта по умолчанию: %s." % (today_iso, DEFAULT_CURRENCY),
        "",
        "КАТЕГОРИИ РАСХОДОВ (строго из списка, дословно):",
        ", ".join(EXPENSE_CATEGORIES),
        "",
        "КАТЕГОРИИ ДОХОДОВ (строго из списка, дословно):",
        ", ".join(INCOME_CATEGORIES),
        "",
        "ЖЕСТКИЕ ПРАВИЛА КАТЕГОРИЗАЦИИ:",
        "",
        "ПРОЧИЕ ПРАВИЛА:",
        "- В чеке магазина/билете извлекай позиции или итоговую сумму (is_bank_statement=false).",
        "- В банковской выписке за период ставь is_bank_statement=true и дату YYYY-MM-DD.",
        "- Если дата в тексте не указана — tx_date = пустая строка.",
        "- project_tag заполняй только при наличии #хештега, без символа #, в нижнем регистре.",
        "- Отвечай СТРОГО валидным JSON без markdown-обертки по схеме:",
        schema,
    ]

    head = parts[:parts.index("ПРОЧИЕ ПРАВИЛА:")]
    tail = parts[parts.index("ПРОЧИЕ ПРАВИЛА:"):]
    return chr(10).join(head + rules + [""] + tail)


def _parse_json_response(raw_text: str) -> ParsedBatch:
    cleaned = (raw_text or "").strip()
    if not cleaned:
        return ParsedBatch(is_bank_statement=False, transactions=[])

    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)

    start_idx = cleaned.find("{")
    end_idx = cleaned.rfind("}")
    if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
        cleaned = cleaned[start_idx : end_idx + 1]

    return ParsedBatch.model_validate_json(cleaned)


async def _call_openai_gateway_with_patience(
    user_content: Union[str, List[Dict[str, Any]]],
    on_retry: RetryCallback = None,
    fallback_user_content: Optional[List[Dict[str, Any]]] = None,
    actor_id: Optional[int] = None,
    actor_name: str = "",
) -> ParsedBatch:
    max_attempts = 6
    backoff_seconds = [2, 4, 8, 12, 16]
    last_err: Optional[Exception] = None
    endpoint = f"{BASE_URL}/chat/completions"
    active_content = user_content
    model_candidates = [ACTIVE_MODEL, "gemini-3.8-flash", "gemini-3-flash", "gemini-3.1-pro"]
    model_idx = 0

    for attempt in range(1, max_attempts + 1):
        current_model = model_candidates[min(model_idx, len(model_candidates) - 1)]
        try:
            payload: Dict[str, Any] = {
                "model": current_model,
                "messages": [
                    {"role": "system", "content": _build_system_prompt(actor_id, actor_name)},
                    {"role": "user", "content": active_content},
                ],
                "temperature": 0.1,
            }

            resp = await http_client.post(endpoint, json=payload)

            if (
                resp.status_code in (400, 422)
                and fallback_user_content is not None
                and active_content is not fallback_user_content
            ):
                active_content = fallback_user_content
                payload["messages"][1]["content"] = active_content
                resp = await http_client.post(endpoint, json=payload)

            if resp.status_code in (400, 404, 503) and "model" in resp.text.lower():
                model_idx += 1

            if resp.status_code >= 400:
                raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:400]}")

            data = resp.json()
            choices = data.get("choices") or []
            if not choices:
                return ParsedBatch(is_bank_statement=False, transactions=[])

            msg_obj = choices[0].get("message", {})
            content_raw = msg_obj.get("content", "")
            if isinstance(content_raw, list):
                content_raw = "\n".join(
                    part.get("text", "") for part in content_raw if isinstance(part, dict)
                )
            return _parse_json_response(str(content_raw))

        except Exception as exc:
            last_err = exc
            err_str = str(exc) or type(exc).__name__
            logger.warning(
                "Шлюз %s (модель %s, попытка %d/%d): %s",
                BASE_URL,
                current_model,
                attempt,
                max_attempts,
                err_str,
            )

            if attempt < max_attempts:
                if on_retry is not None:
                    try:
                        await on_retry(attempt + 1, max_attempts)
                    except Exception:
                        pass
                wait_sec = backoff_seconds[min(attempt - 1, len(backoff_seconds) - 1)]
                await asyncio.sleep(wait_sec)
                continue
            break

    err_msg = str(last_err) or type(last_err).__name__
    raise RuntimeError(f"API ({ACTIVE_MODEL}) не ответил после {max_attempts} попыток: {err_msg}")


_call_gemini_38_with_patience = _call_openai_gateway_with_patience


async def parse_text_message(
    text: str,
    on_retry: RetryCallback = None,
    actor_id: Optional[int] = None,
    actor_name: str = "",
) -> ParsedBatch:
    fast_result = try_fast_local_parse(
        text, allow_fallback_mode=False, actor_name=actor_name
    )
    if fast_result is not None:
        return fast_result

    prompt = f"Разбери на финансовые транзакции:\n{text}"
    try:
        return await _call_openai_gateway_with_patience(
            prompt, on_retry=on_retry, actor_id=actor_id, actor_name=actor_name
        )
    except Exception as exc:
        logger.warning("Фоллбэк на локальный парсер текста после ошибки: %s", exc)
        fallback = try_fast_local_parse(
            text, allow_fallback_mode=True, actor_name=actor_name
        )
        if fallback is not None:
            return fallback
        raise


async def parse_receipt_image(
    image_bytes: bytes, caption: str = "", on_retry: RetryCallback = None
) -> ParsedBatch:
    optimized_bytes = await asyncio.to_thread(_optimize_image_bytes, image_bytes)
    b64_img = base64.b64encode(optimized_bytes).decode("ascii")

    prompt = "Извлеки транзакции или итоговую сумму из этого чека/билета/выписки."
    if caption:
        prompt += f" Комментарий: {caption}"

    user_content: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt},
        {
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{b64_img}"},
        },
    ]
    return await _call_openai_gateway_with_patience(user_content, on_retry=on_retry)


async def parse_voice_message(
    audio_bytes: bytes, on_retry: RetryCallback = None
) -> ParsedBatch:
    b64_audio = base64.b64encode(audio_bytes).decode("ascii")
    prompt = "Извлеки все упомянутые доходы и расходы из аудиосообщения."

    primary_content: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt},
        {
            "type": "image_url",
            "image_url": {"url": f"data:audio/ogg;base64,{b64_audio}"},
        },
    ]
    fallback_content: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt},
        {
            "type": "input_audio",
            "input_audio": {"data": b64_audio, "format": "ogg"},
        },
    ]
    return await _call_openai_gateway_with_patience(
        primary_content,
        on_retry=on_retry,
        fallback_user_content=fallback_content,
    )


async def parse_document_statement(
    file_bytes: bytes,
    mime_type: str,
    text_fallback: str = "",
    caption: str = "",
    on_retry: RetryCallback = None,
    actor_id: Optional[int] = None,
    actor_name: str = "",
) -> ParsedBatch:
    prompt = (
        "Это банковская выписка. Извлеки все операции, даты YYYY-MM-DD, "
        "суммы и категории (is_bank_statement=true)."
    )
    if caption:
        prompt += f" Комментарий: {caption}"

    if text_fallback:
        trimmed_text = text_fallback[:120_000]
        return await _call_openai_gateway_with_patience(
            f"{prompt}\n\nТекст:\n{trimmed_text}",
            on_retry=on_retry,
            actor_id=actor_id,
            actor_name=actor_name,
        )

    b64_doc = base64.b64encode(file_bytes).decode("ascii")
    user_content: List[Dict[str, Any]] = [
        {"type": "text", "text": prompt},
        {
            "type": "image_url",
            "image_url": {"url": f"data:{mime_type};base64,{b64_doc}"},
        },
    ]
    return await _call_openai_gateway_with_patience(user_content, on_retry=on_retry)