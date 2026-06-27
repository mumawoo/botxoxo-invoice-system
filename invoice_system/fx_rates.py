from __future__ import annotations

import html
import re
import ssl
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Iterable

SAFE_RMB_QUERY_URL = "https://www.safe.gov.cn/AppStructured/hlw/RMBQuery.do"

DIRECT_CNY_PER_100_CODES = {
    "USD",
    "EUR",
    "JPY",
    "HKD",
    "GBP",
    "AUD",
    "NZD",
    "SGD",
    "CHF",
    "CAD",
    "MOP",
}

INDIRECT_FOREIGN_PER_100_CNY_CODES = {
    "MYR",
    "RUB",
    "ZAR",
    "KRW",
    "AED",
    "SAR",
    "HUF",
    "PLN",
    "DKK",
    "SEK",
    "NOK",
    "TRY",
    "MXN",
    "THB",
}

SAFE_CURRENCY_CODES = [
    "USD",
    "EUR",
    "JPY",
    "HKD",
    "GBP",
    "AUD",
    "NZD",
    "SGD",
    "CHF",
    "CAD",
    "MOP",
    "MYR",
    "RUB",
    "ZAR",
    "KRW",
    "AED",
    "SAR",
    "HUF",
    "PLN",
    "DKK",
    "SEK",
    "NOK",
    "TRY",
    "MXN",
    "THB",
]


@dataclass(frozen=True)
class ExchangeRate:
    rate_date: date
    usd_cny_per_100: float
    mxn_per_100_cny: float
    source_url: str = SAFE_RMB_QUERY_URL
    fetched_at: str = ""
    rates: dict[str, float] = field(default_factory=dict)

    @property
    def usd_to_mxn(self) -> float:
        return (self.usd_cny_per_100 / 100.0) * (self.mxn_per_100_cny / 100.0)

    @property
    def cny_to_mxn(self) -> float:
        return self.mxn_per_100_cny / 100.0

    def multiplier_to_mxn(self, currency: str) -> float | None:
        code = normalize_safe_currency(currency)
        if code == "MXN":
            return 1.0
        if code in {"CNY", "RMB"}:
            return self.cny_to_mxn
        value = self.rate_value(code)
        if value <= 0:
            return None
        if code in DIRECT_CNY_PER_100_CODES:
            return (value / 100.0) * self.cny_to_mxn
        if code in INDIRECT_FOREIGN_PER_100_CNY_CODES:
            return self.mxn_per_100_cny / value
        return None

    def rate_value(self, currency: str) -> float:
        code = normalize_safe_currency(currency)
        if code == "USD":
            return self.usd_cny_per_100
        if code == "MXN":
            return self.mxn_per_100_cny
        return float(self.rates.get(code) or 0)


def normalize_safe_currency(value: str) -> str:
    normalized = re.sub(r"[^A-Za-z]", "", value or "").upper()
    aliases = {
        "US": "USD",
        "DOLLAR": "USD",
        "DOLLARS": "USD",
        "EURO": "EUR",
        "EUROS": "EUR",
        "RMB": "CNY",
        "YUAN": "CNY",
        "RENMINBI": "CNY",
        "MX": "MXN",
        "MN": "MXN",
        "PESO": "MXN",
        "PESOS": "MXN",
        "RMBYUAN": "CNY",
    }
    return aliases.get(normalized, normalized or "MXN")


def fetch_safe_exchange_rates(start_date: date | None = None, end_date: date | None = None, timeout: int = 15) -> list[ExchangeRate]:
    url = _safe_url(start_date, end_date)
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 invoice-system",
            "Accept": "text/html,application/xhtml+xml",
        },
    )
    raw = _urlopen_safe(request, timeout)
    text = html.unescape(raw.decode("utf-8", errors="ignore"))
    fetched_at = datetime.now().isoformat(timespec="seconds")
    rates: list[ExchangeRate] = []
    for rate_date, numbers in _parse_rate_rows(text):
        values = {
            code: numbers[index]
            for index, code in enumerate(SAFE_CURRENCY_CODES)
            if index < len(numbers)
        }
        rates.append(
            ExchangeRate(
                rate_date=rate_date,
                usd_cny_per_100=values.get("USD", 0.0),
                mxn_per_100_cny=values.get("MXN", 0.0),
                source_url=url,
                fetched_at=fetched_at,
                rates=values,
            )
        )
    return rates


def nearest_rate_on_or_before(rates: Iterable[ExchangeRate], target: date) -> ExchangeRate | None:
    eligible = [rate for rate in rates if rate.rate_date <= target]
    if not eligible:
        return None
    return max(eligible, key=lambda rate: rate.rate_date)


def _safe_url(start_date: date | None, end_date: date | None) -> str:
    if not start_date and not end_date:
        return SAFE_RMB_QUERY_URL
    params = {
        "startDate": (start_date or end_date or date.today()).isoformat(),
        "endDate": (end_date or start_date or date.today()).isoformat(),
        "queryYN": "true",
    }
    return SAFE_RMB_QUERY_URL + "?" + urllib.parse.urlencode(params)


def _urlopen_safe(request: urllib.request.Request, timeout: int) -> bytes:
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.read()
    except urllib.error.URLError as exc:
        if "CERTIFICATE_VERIFY_FAILED" not in str(exc):
            raise
        context = ssl._create_unverified_context()
        with urllib.request.urlopen(request, timeout=timeout, context=context) as response:
            return response.read()


def _parse_rate_rows(text: str) -> list[tuple[date, list[float]]]:
    cells = [_clean_cell(value) for value in re.findall(r"<td[^>]*>(.*?)</td>", text, flags=re.IGNORECASE | re.DOTALL)]
    cells = [value for value in cells if value]
    rows: list[tuple[date, list[float]]] = []
    index = 0
    while index < len(cells):
        value = cells[index]
        if re.fullmatch(r"20\d{2}-\d{2}-\d{2}", value):
            numbers: list[float] = []
            for offset in range(1, 26):
                if index + offset >= len(cells):
                    break
                if not re.fullmatch(r"\d+(?:\.\d+)?", cells[index + offset]):
                    break
                numbers.append(float(cells[index + offset]))
            if len(numbers) >= 24:
                rows.append((date.fromisoformat(value), numbers))
                index += len(numbers)
        index += 1
    if rows:
        return rows
    stripped = re.sub(r"<[^>]+>", " ", text)
    for match in re.finditer(r"(20\d{2}-\d{2}-\d{2})\s+((?:\d+(?:\.\d+)?\s+){24}\d+(?:\.\d+)?)", stripped):
        numbers = [float(value) for value in match.group(2).split()]
        rows.append((date.fromisoformat(match.group(1)), numbers))
    return rows


def _clean_cell(value: str) -> str:
    value = re.sub(r"<[^>]+>", " ", value)
    return re.sub(r"\s+", " ", value).strip()
