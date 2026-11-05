# AI Finance Tracker Bot

Multi-user Telegram bot designed for automated personal and family expense management using LLM-driven transaction parsing.

## Key Features
- **AI Transaction Parsing:** Converts free-form natural language text and voice logs into structured financial records.
- **Multi-Tenant Routing:** Group and user workspace isolation for shared family budgets.
- **Analytics & Reporting:** Dynamic period filtering, Pandas-powered financial aggregation, and automated database exports.

## Tech Stack
- **Core:** Python 3.11, aiogram 3
- **Data Engine:** SQLite, Pandas
- **AI Processing:** DeepSeek / Gemini LLM API integration
- **Deployment:** Docker, Docker Compose