import json
import threading
import time
import urllib.request
from typing import Dict

FALLBACK_RATES_TO_BYN: Dict[str, float] = {
    "BYN": 1.0,
    "USD": 3.32,
    "EUR": 3.58,
    "RUB": 0.0345,
}

_cache_rates: Dict[str, float] = FALLBACK_RATES_TO_BYN.copy()
_cache_timestamp: float = 0.0
_rate_lock = threading.Lock()
CACHE_TTL_SECONDS = 21600  # 6 часов


def get_rates_to_byn() -> Dict[str, float]:
    global _cache_rates, _cache_timestamp
    now = time.time()
    if now - _cache_timestamp < CACHE_TTL_SECONDS:
        return _cache_rates

    with _rate_lock:
        if time.time() - _cache_timestamp < CACHE_TTL_SECONDS:
            return _cache_rates
        try:
            req = urllib.request.Request(
                "https://api.nbrb.by/exrates/rates?periodicity=0",
                headers={"User-Agent": "FinanceBot/3.3"},
            )
            with urllib.request.urlopen(req, timeout=2.5) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            rates = {"BYN": 1.0}
            for item in data:
                abbr = item.get("Cur_Abbreviation")
                scale = float(item.get("Cur_Scale") or 1.0)
                official = float(item.get("Cur_OfficialRate") or 0.0)
                if abbr and official > 0:
                    rates[abbr.upper()] = official / scale
            for k, v in FALLBACK_RATES_TO_BYN.items():
                rates.setdefault(k, v)
            _cache_rates = rates
            _cache_timestamp = time.time()
        except Exception:
            _cache_timestamp = time.time() - (CACHE_TTL_SECONDS - 300)

    return _cache_rates


def convert_to_byn(amount: float, currency: str) -> float:
    cur = (currency or "BYN").upper().strip()
    if cur == "BYN":
        return round(float(amount), 2)
    rates = get_rates_to_byn()
    rate = rates.get(cur, FALLBACK_RATES_TO_BYN.get(cur, 1.0))
    return round(float(amount) * rate, 2)