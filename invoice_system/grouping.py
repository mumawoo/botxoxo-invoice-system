from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


GROUP_STATE_FILE = "telegram_group_state.json"


def group_state_path(output_dir: Path) -> Path:
    return output_dir / GROUP_STATE_FILE


def arm_next_group(output_dir: Path) -> None:
    data = _load_group_state(output_dir)
    data["armed"] = True
    data["updated_at"] = _timestamp()
    _save_group_state(output_dir, data)


def mark_source_for_group_if_armed(output_dir: Path, source_path: Path) -> bool:
    data = _load_group_state(output_dir)
    if not data.get("armed"):
        return False
    data["armed"] = False
    forced = data.setdefault("forced_sources", {})
    forced[_source_key(source_path)] = {"path": str(source_path), "created_at": _timestamp()}
    data["updated_at"] = _timestamp()
    _save_group_state(output_dir, data)
    return True


def is_forced_group_source(output_dir: Path, source_path: Path) -> bool:
    data = _load_group_state(output_dir)
    return _source_key(source_path) in dict(data.get("forced_sources") or {})


def clear_forced_group_source(output_dir: Path, source_path: Path) -> None:
    data = _load_group_state(output_dir)
    forced = dict(data.get("forced_sources") or {})
    forced.pop(_source_key(source_path), None)
    data["forced_sources"] = forced
    data["updated_at"] = _timestamp()
    _save_group_state(output_dir, data)


def save_pending_group(output_dir: Path, crop_ids: list[str] | tuple[str, ...]) -> None:
    data = _load_group_state(output_dir)
    data["pending"] = {"crop_ids": list(crop_ids), "created_at": _timestamp()}
    data["updated_at"] = _timestamp()
    _save_group_state(output_dir, data)


def load_pending_group(output_dir: Path) -> list[str]:
    data = _load_group_state(output_dir)
    pending = data.get("pending") if isinstance(data.get("pending"), dict) else {}
    return [str(item) for item in pending.get("crop_ids", [])]


def clear_pending_group(output_dir: Path) -> None:
    data = _load_group_state(output_dir)
    data.pop("pending", None)
    data["armed"] = False
    data["updated_at"] = _timestamp()
    _save_group_state(output_dir, data)


def _load_group_state(output_dir: Path) -> dict:
    path = group_state_path(output_dir)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def _save_group_state(output_dir: Path, data: dict) -> None:
    path = group_state_path(output_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def _source_key(path: Path) -> str:
    try:
        return str(path.resolve())
    except OSError:
        return str(path)


def _timestamp() -> str:
    return datetime.now().replace(microsecond=0).isoformat()
