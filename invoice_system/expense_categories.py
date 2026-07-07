from __future__ import annotations

import unicodedata
from pathlib import Path

from .company_profile import company_category_override

EXPENSE_CATEGORIES = (
    "Food",
    "Gas",
    "Car repair",
    "Toll/Parking",
    "Utilities",
    "Internet",
    "Phone",
    "Office supplies",
    "Hotel",
    "Flight",
    "Other",
)

DEFAULT_EXPENSE_CATEGORY = "Other"

_CATEGORY_KEYWORDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    (
        "Food",
        (
            "餐饮",
            "food",
            "restaurant",
            "restaurante",
            "beverage",
            "cafe",
            "coffee",
            "bar",
            "comida",
            "alimentos",
            "bebida",
            "mercado pago",
        ),
    ),
    (
        "Gas",
        (
            "汽油",
            "gasolina",
            "combustible",
            "fuel",
            "gas station",
            "petrol",
            "pemex",
            "shell",
            "mobil",
            "bp",
            "oxxo gas",
        ),
    ),
    (
        "Car repair",
        (
            "车辆维修",
            "auto repair",
            "car repair",
            "mechanic",
            "mecanico",
            "mecanica",
            "taller",
            "refaccion",
            "refacciones",
            "llanta",
            "llantas",
            "aceite",
            "servicio automotriz",
        ),
    ),
    (
        "Toll/Parking",
        (
            "过路费",
            "停车费",
            "toll",
            "caseta",
            "peaje",
            "parking",
            "estacionamiento",
            "parquimetro",
            "autopista",
        ),
    ),
    (
        "Utilities",
        (
            "水电煤",
            "electricidad",
            "luz",
            "water",
            "agua",
            "gas natural",
            "gas lp",
            "cfe",
            "sadm",
        ),
    ),
    (
        "Internet",
        (
            "网络费",
            "internet",
            "wifi",
            "fibra",
            "telmex",
            "izzi",
            "totalplay",
            "megacable",
            "modem",
        ),
    ),
    (
        "Phone",
        (
            "电话费",
            "telefono",
            "teléfono",
            "phone",
            "mobile",
            "celular",
            "telcel",
            "movistar",
            "at&t",
            "att",
        ),
    ),
    (
        "Office supplies",
        (
            "办公用品",
            "office supplies",
            "papeleria",
            "papelería",
            "stationery",
            "printer",
            "toner",
            "paper",
            "papel",
            "pluma",
            "lapiz",
            "lápiz",
            "libreta",
            "office depot",
            "officemax",
        ),
    ),
    (
        "Hotel",
        (
            "住宿",
            "hotel",
            "motel",
            "hospedaje",
            "alojamiento",
            "lodging",
            "inn",
            "airbnb",
            "hilton",
            "marriott",
            "holiday inn",
            "fiesta inn",
            "city express",
        ),
    ),
    (
        "Flight",
        (
            "机票",
            "airfare",
            "flight",
            "airline",
            "boarding pass",
            "boleto de avion",
            "boleto de avión",
            "vuelo",
            "aeromexico",
            "aeroméxico",
            "viva aerobus",
            "volaris",
        ),
    ),
    (
        "Other",
        (
            "其他",
            "other",
            "otros",
            "misc",
            "miscellaneous",
        ),
    ),
)


def normalize_expense_category(value: str, evidence: str = "", *, company_profile: str | None = None, root: Path | None = None) -> str:
    override = company_category_override(value, evidence, name=company_profile, root=root)
    if override:
        known = _known_category(override)
        if known:
            return known
    text = " ".join(part for part in (value, evidence) if part)
    normalized = _normalize_text(text).casefold().replace("_", " ")
    for category in EXPENSE_CATEGORIES:
        if _normalize_text(category).casefold() in normalized:
            return category
    for category, keywords in _CATEGORY_KEYWORDS:
        if any(_normalize_text(keyword).casefold() in normalized for keyword in keywords):
            return category
    return DEFAULT_EXPENSE_CATEGORY


def _known_category(value: str) -> str:
    normalized = _normalize_text(value).casefold()
    for category in EXPENSE_CATEGORIES:
        if _normalize_text(category).casefold() == normalized:
            return category
    return ""


def _normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).strip()
