"""Atomic publish of workbook + JSON into data/current."""
from __future__ import annotations

import json
import os
import shutil
from datetime import datetime, timezone
from pathlib import Path

from app.config import ARCHIVE, CURRENT, STAGING, WORKBOOK_FILE
from app.export_data import export_json


def _copy_atomic(src: Path, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".tmp")
    shutil.copy2(src, tmp)
    os.replace(tmp, dest)


def publish_workbook(
    staged_workbook: Path,
    *,
    upload_id: str,
    warnings: list[str] | None = None,
) -> dict:
    version = upload_id or datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    archive_dir = ARCHIVE / version
    archive_dir.mkdir(parents=True, exist_ok=True)

    archived_wb = archive_dir / staged_workbook.name
    shutil.copy2(staged_workbook, archived_wb)

    export_result = export_json(archived_wb, version=version, warnings=warnings)
    _copy_atomic(archived_wb, WORKBOOK_FILE)

    # Snapshot published JSON + status into archive.
    for name in ("dashboard-data.json", "status.json"):
        src = CURRENT / name
        if src.exists():
            shutil.copy2(src, archive_dir / name)

    manifest = {
        "version": version,
        "publishedAt": datetime.now(timezone.utc).isoformat(),
        "workbook": staged_workbook.name,
        "rowCount": export_result["payload"]["rowCount"],
        "warnings": warnings or [],
    }
    (archive_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    (ARCHIVE / "history.jsonl").open("a", encoding="utf-8").write(json.dumps(manifest) + "\n")
    return manifest


def rollback(version: str) -> dict:
    archive_dir = ARCHIVE / version
    if not archive_dir.exists():
        raise FileNotFoundError(f"Archived version not found: {version}")

    wb = archive_dir / "Consolidated_Ocean_DSR.xlsx"
    data_json = archive_dir / "dashboard-data.json"
    status_json = archive_dir / "status.json"
    if not wb.exists() or not data_json.exists():
        raise FileNotFoundError("Archive is missing workbook or dashboard JSON.")

    _copy_atomic(wb, WORKBOOK_FILE)
    _copy_atomic(data_json, CURRENT / "dashboard-data.json")
    if status_json.exists():
        _copy_atomic(status_json, CURRENT / "status.json")

    manifest_path = archive_dir / "manifest.json"
    if manifest_path.exists():
        return json.loads(manifest_path.read_text(encoding="utf-8"))
    return {"version": version, "rolledBack": True}


def list_history(limit: int = 30) -> list[dict]:
    history_file = ARCHIVE / "history.jsonl"
    if not history_file.exists():
        return []
    lines = history_file.read_text(encoding="utf-8").strip().splitlines()
    items = [json.loads(line) for line in lines if line.strip()]
    return list(reversed(items[-limit:]))
