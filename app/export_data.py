"""Export dashboard JSON from consolidated workbook."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from app.config import JSON_FILE, STATUS_FILE, WORKBOOK_FILE
from app.data_loader import build_payload, load_workbook_records, slim


def export_json(
    workbook_path: Path,
    *,
    version: str | None = None,
    warnings: List[str] | None = None,
) -> Dict[str, Any]:
    version = version or datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    raw = load_workbook_records(workbook_path)
    records = [slim(r) for r in raw]
    payload = build_payload(records, workbook_path, version)

    JSON_FILE.parent.mkdir(parents=True, exist_ok=True)
    JSON_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    status = {
        "lastRefresh": datetime.now(timezone.utc).isoformat(),
        "modified": payload["modified"],
        "rowCount": payload["rowCount"],
        "version": version,
        "workbook": workbook_path.name,
        "warnings": warnings or [],
    }
    STATUS_FILE.write_text(json.dumps(status, indent=2), encoding="utf-8")
    return {"payload": payload, "status": status}


def read_status() -> Dict[str, Any]:
    if not STATUS_FILE.exists():
        return {
            "lastRefresh": None,
            "modified": None,
            "rowCount": 0,
            "version": None,
            "workbook": None,
            "warnings": [],
        }
    return json.loads(STATUS_FILE.read_text(encoding="utf-8"))
