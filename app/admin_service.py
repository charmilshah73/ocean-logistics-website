"""Admin upload pipeline: validate → ETL → publish."""
from __future__ import annotations

import json
import shutil
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from werkzeug.datastructures import FileStorage

from app.config import INPUT, MAX_UPLOAD_MB, STAGING
from app.etl import etl_summary, run_etl
from app.publish import publish_workbook

ALLOWED = {".xlsx"}
REQUIRED_TYPES = {"liyana", "gocomet"}
SOURCE_HINTS = {
    "liyana": ("liyana",),
    "gocomet": ("gocomet", "detailed", "tracking"),
    "scanglobal": ("scanglobal", "scan global",),
    "cargomar": ("cargomar",),
    "sinpex": ("sinpex", "china shipment",),
    "open_order": ("open order", "open_order"),
    "hotlist": ("hotlist", "hot list", "hot_list", "corporate hot", "corporate_hot"),
}

CANONICAL_NAMES = {
    "hotlist": "Corporate Hot list.xlsx",
    "open_order": "Open Order Report.xlsx",
}

def _safe_filename(original: str, source_type: str) -> str:
    if source_type in CANONICAL_NAMES:
        return CANONICAL_NAMES[source_type]
    name = Path(original).name.replace("\x00", "").replace("/", " ").replace("\\", " ").strip()
    name = " ".join(name.split())
    if not name.lower().endswith(".xlsx"):
        name = f"{name or 'upload'}.xlsx"
    return name

_lock = threading.Lock()


def _norm_name(filename: str) -> str:
    return Path(filename).name.lower().replace("_", " ").replace("-", " ")


def _detect_type(filename: str) -> str:
    name = _norm_name(filename)
    if "hot" in name and "list" in name:
        return "hotlist"
    if "open" in name and "order" in name:
        return "open_order"
    for source_type, hints in SOURCE_HINTS.items():
        if any(h in name for h in hints):
            return source_type
    return "unknown"


def _validate_upload(files: List[FileStorage]) -> Dict[str, Any]:
    if not files:
        raise ValueError("No files were uploaded.")

    max_bytes = MAX_UPLOAD_MB * 1024 * 1024
    detected: Dict[str, str] = {}
    errors: List[str] = []

    for fs in files:
        if not fs or not fs.filename:
            continue
        ext = Path(fs.filename).suffix.lower()
        if ext not in ALLOWED:
            errors.append(f"{fs.filename}: only .xlsx files are allowed.")
            continue
        fs.stream.seek(0, 2)
        size = fs.stream.tell()
        fs.stream.seek(0)
        if size > max_bytes:
            errors.append(f"{fs.filename}: exceeds {MAX_UPLOAD_MB} MB limit.")
        source_type = _detect_type(fs.filename)
        detected[fs.filename] = source_type

    if errors:
        raise ValueError("\n".join(errors))

    found_types = set(detected.values())
    missing = REQUIRED_TYPES - found_types
    if missing:
        raise ValueError(
            "Missing required file types: "
            + ", ".join(sorted(missing))
            + ". Upload must include Liyana DSR and GoComet/Detailed Tracking."
        )
    return {"files": detected}


def process_upload(files: List[FileStorage]) -> Dict[str, Any]:
    with _lock:
        validation = _validate_upload(files)
        upload_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        staging_input = STAGING / upload_id / "input"
        staging_output = STAGING / upload_id / "output"
        if staging_input.exists():
            shutil.rmtree(staging_input)
        staging_input.mkdir(parents=True, exist_ok=True)
        staging_output.mkdir(parents=True, exist_ok=True)

        saved: List[str] = []
        used_names: set[str] = set()
        for fs in files:
            if not fs or not fs.filename:
                continue
            source_type = _detect_type(fs.filename)
            safe = _safe_filename(fs.filename, source_type)
            if not safe:
                safe = f"upload_{len(saved)+1}.xlsx"
            if safe.lower() in used_names:
                stem = Path(safe).stem
                safe = f"{stem}_{len(saved)+1}.xlsx"
            used_names.add(safe.lower())
            dest = staging_input / safe
            fs.save(dest)
            saved.append(safe)

        for path in list(staging_input.glob("*.xlsx")):
            kind = _detect_type(path.name)
            canonical = CANONICAL_NAMES.get(kind)
            if not canonical or path.name == canonical:
                continue
            dest = staging_input / canonical
            if dest.exists() and dest.resolve() != path.resolve():
                dest.unlink()
            path.rename(dest)
            saved = [canonical if name == path.name else name for name in saved]

        code = run_etl(staging_input, staging_output)
        summary = etl_summary(staging_output)
        if code != 0:
            raise RuntimeError(
                "Consolidation failed. Ensure Liyana and GoComet files are valid."
            )
        if summary["rowCount"] <= 0:
            raise RuntimeError("Consolidation produced zero rows.")

        workbook = Path(summary["workbook"])
        manifest = publish_workbook(
            workbook,
            upload_id=upload_id,
            warnings=summary.get("warnings", []),
        )
        manifest["savedFiles"] = saved
        manifest["detectedTypes"] = validation["files"]
        return manifest
