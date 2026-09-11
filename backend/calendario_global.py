"""Consolida calendarios y persiste movimientos/eventos manuales."""

from __future__ import annotations

import hashlib
import json
import uuid
from collections import defaultdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from .calendario_rentas import get_saved_result
from .contingente import get_saved_contingente


ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
CUSTOM_PATH = DATA_DIR / "calendario_global_personalizado.json"


def _load_custom() -> dict[str, Any]:
    default = {"moves": {}, "manual": []}
    if not CUSTOM_PATH.is_file():
        return default
    try:
        saved = json.loads(CUSTOM_PATH.read_text(encoding="utf-8"))
        return {
            "moves": saved.get("moves") or {},
            "manual": saved.get("manual") or [],
        }
    except (OSError, json.JSONDecodeError):
        return default


def _save_custom(custom: dict[str, Any]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = CUSTOM_PATH.with_suffix(".tmp")
    temp_path.write_text(
        json.dumps(custom, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temp_path.replace(CUSTOM_PATH)


def _event_id(source: str, day: str, index: int, event: dict[str, Any]) -> str:
    payload = json.dumps(
        {"source": source, "day": day, "index": index, "event": event},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]


def _base_events() -> tuple[list[dict[str, Any]], dict[str, Any], dict[str, Any]]:
    rentas = get_saved_result()
    contingente = get_saved_contingente()
    result: list[dict[str, Any]] = []

    for day, events in (rentas.get("calendar") or {}).get("by_day", {}).items():
        for index, event in enumerate(events):
            result.append(
                {
                    "id": _event_id("Rentas", day, index, event),
                    "date": day,
                    "original_date": day,
                    "source": "Rentas",
                    "title": event.get("ticker") or "",
                    "subtitle": event.get("type") or "Pago",
                    "ticker": event.get("ticker"),
                    "quantity": event.get("quantity"),
                    "summary": event.get("summary"),
                }
            )

    for day, events in (contingente.get("calendar") or {}).get("by_day", {}).items():
        for index, event in enumerate(events):
            estado = event.get("estado") or ""
            result.append(
                {
                    "id": _event_id("Contingente", day, index, event),
                    "date": day,
                    "original_date": day,
                    "source": "Contingente",
                    "title": event.get("empresa") or "",
                    "subtitle": estado,
                    "amount": event.get("monto"),
                    "aforo": event.get("aforo"),
                    "due_date": event.get("vencimiento"),
                    "claimed": "reclamada por el banco" in str(estado).lower(),
                }
            )
    return result, rentas, contingente


def _validate_date(value: Any) -> str:
    raw = str(value or "").strip()
    try:
        return date.fromisoformat(raw).isoformat()
    except ValueError as exc:
        raise ValueError("La fecha debe tener formato AAAA-MM-DD.") from exc


def get_global_calendar() -> dict[str, Any]:
    base_events, rentas, contingente = _base_events()
    custom = _load_custom()
    events = base_events + [dict(event) for event in custom["manual"]]
    by_day: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for event in events:
        original_date = event.get("original_date") or event.get("date")
        moved_date = custom["moves"].get(event["id"])
        event["original_date"] = original_date
        event["date"] = moved_date or event.get("date")
        event["moved"] = bool(moved_date and moved_date != original_date)
        by_day[event["date"]].append(event)

    for day_events in by_day.values():
        day_events.sort(key=lambda event: (event["source"], event["title"]))

    years = sorted({day[:4] for day in by_day})
    rentas_count = sum(1 for event in base_events if event["source"] == "Rentas")
    contingente_count = sum(
        1 for event in base_events if event["source"] == "Contingente"
    )
    manual_count = len(custom["manual"])

    return {
        "calendar": {"by_day": dict(by_day), "years": years},
        "summary": {
            "total": rentas_count + contingente_count + manual_count,
            "rentas": rentas_count,
            "contingente": contingente_count,
            "manual": manual_count,
            "moved": sum(
                1
                for day_events in by_day.values()
                for event in day_events
                if event.get("moved")
            ),
        },
        "sources": {
            "rentas": {
                "available": bool(rentas.get("meta")),
                "filename": (rentas.get("meta") or {}).get("position_file"),
                "updated_at": (rentas.get("meta") or {}).get("updated_at"),
            },
            "contingente": {
                "available": bool(contingente.get("meta")),
                "filename": contingente.get("filename"),
                "updated_at": (contingente.get("meta") or {}).get("updated_at"),
            },
        },
    }


def move_global_event(event_id: str, new_date: Any) -> dict[str, Any]:
    event_id = str(event_id or "").strip()
    target_date = _validate_date(new_date)
    current = get_global_calendar()
    all_events = [
        event
        for events in current["calendar"]["by_day"].values()
        for event in events
    ]
    event = next((item for item in all_events if item.get("id") == event_id), None)
    if event is None:
        raise ValueError("El evento ya no existe en el calendario.")

    custom = _load_custom()
    if target_date == event.get("original_date"):
        custom["moves"].pop(event_id, None)
    else:
        custom["moves"][event_id] = target_date
    _save_custom(custom)
    return get_global_calendar()


def add_manual_event(
    event_type: Any, denomination: Any, event_date: Any, comment: Any = ""
) -> dict[str, Any]:
    event_type = str(event_type or "").strip()
    denomination = str(denomination or "").strip()
    comment = str(comment or "").strip()
    normalized_date = _validate_date(event_date)
    if not event_type:
        raise ValueError("Indicá el tipo de evento.")
    if not denomination:
        raise ValueError("Indicá la denominación.")

    custom = _load_custom()
    custom["manual"].append(
        {
            "id": "manual-" + uuid.uuid4().hex,
            "date": normalized_date,
            "original_date": normalized_date,
            "source": "Manual",
            "title": denomination,
            "subtitle": event_type,
            "comment": comment,
            "summary": comment,
            "created_at": datetime.now(timezone.utc).isoformat(),
        }
    )
    _save_custom(custom)
    return get_global_calendar()
