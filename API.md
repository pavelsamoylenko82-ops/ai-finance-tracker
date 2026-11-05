# Finance API для бота-напоминатора

Сервис поднимается вместе с финансовым ботом на внутреннем порту `8081`
(наружу публиковать не обязательно).

## Переменные

```env
FINANCE_API_HOST=0.0.0.0
FINANCE_API_PORT=8081
FINANCE_API_TOKEN=секрет
```

Если `FINANCE_API_TOKEN` пуст, API работает без проверки токена. Это допустимо
только внутри закрытой Docker-сети.

## POST /api/expenses

Заголовок при заданном токене:

```
X-API-Token: секрет
```

Тело:

```json
{
  "source": "family_reminder_bot",
  "author": "Паша (Сын)",
  "external_id": "shop-84563f38db621aef4bb1",
  "items": [
    {
      "description": "Молоко",
      "amount": 2.5,
      "currency": "BYN",
      "category": "Продукты",
      "tx_type": "expense",
      "tx_date": "2026-10-28"
    }
  ]
}
```

Ответ:

```json
{
  "ok": true,
  "duplicate": false,
  "batch_id": "shop-84563f38db621aef4bb1",
  "count": 2,
  "total": 3.7
}
```

`external_id` обеспечивает идемпотентность: повторный запрос с тем же
идентификатором не создаёт дубликаты, а возвращает `duplicate: true`.

## GET /health

```json
{"ok": true, "service": "finance_bot", "api_version": 1, "auth_required": true}
```
