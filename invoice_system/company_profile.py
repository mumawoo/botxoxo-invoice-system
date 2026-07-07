from __future__ import annotations

import json
import os
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path


DEFAULT_COMPANY_PROFILE = "default"
COMPANY_PROFILE_ENV = "COMPANY_PROFILE"
COMPANY_PROFILES_DIR = Path("rules") / "company_profiles"


@dataclass(frozen=True)
class CategoryOverride:
    match: str
    category: str


@dataclass(frozen=True)
class CompanyProfile:
    name: str
    path: Path
    category_overrides: tuple[CategoryOverride, ...] = ()
    loaded: bool = False
    warning: str = ""


def configured_company_profile(value: str | None = None) -> str:
    text = str(value if value is not None else os.getenv(COMPANY_PROFILE_ENV, DEFAULT_COMPANY_PROFILE)).strip()
    return text or DEFAULT_COMPANY_PROFILE


def load_company_profile(name: str | None = None, root: Path | None = None) -> CompanyProfile:
    profile_name = configured_company_profile(name)
    base = Path("." if root is None else root)
    path = base / COMPANY_PROFILES_DIR / f"{_safe_profile_name(profile_name)}.json"
    if not path.exists():
        return CompanyProfile(profile_name, path, loaded=False, warning=f"Company profile not found: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return CompanyProfile(profile_name, path, loaded=False, warning=f"Company profile could not be loaded: {exc}")
    overrides: list[CategoryOverride] = []
    for item in data.get("category_overrides", []) if isinstance(data, dict) else []:
        if not isinstance(item, dict):
            continue
        match = str(item.get("match") or "").strip()
        category = str(item.get("category") or "").strip()
        if match and category:
            overrides.append(CategoryOverride(match, category))
    return CompanyProfile(str(data.get("name") or profile_name), path, tuple(overrides), loaded=True)


def company_category_override(value: str, evidence: str = "", *, name: str | None = None, root: Path | None = None) -> str:
    profile = load_company_profile(name=name, root=root)
    text = " ".join(part for part in (value, evidence) if part)
    normalized = _normalize_match_text(text)
    if not normalized:
        return ""
    for override in profile.category_overrides:
        if _normalize_match_text(override.match) in normalized:
            return override.category
    return ""


def company_profile_status(name: str | None = None, root: Path | None = None) -> str:
    profile = load_company_profile(name=name, root=root)
    if profile.loaded:
        return f"{profile.name} ({len(profile.category_overrides)} rule(s))"
    return f"{profile.name} (missing)"


def company_profile_warning(name: str | None = None, root: Path | None = None) -> str:
    return load_company_profile(name=name, root=root).warning


def _safe_profile_name(value: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip())
    return safe or DEFAULT_COMPANY_PROFILE


def _normalize_match_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    without_marks = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", without_marks.casefold())
