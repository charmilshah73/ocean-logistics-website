"""Paths and settings — all data stays inside logistics_website."""
from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
V10_PKG = ROOT.parent / "v10_pkg"

DATA = ROOT / "data"
CURRENT = DATA / "current"
STAGING = DATA / "staging"
ARCHIVE = DATA / "archive"
INPUT = DATA / "input"
LOGS = DATA / "logs"

for folder in (CURRENT, STAGING, ARCHIVE, INPUT, LOGS):
    folder.mkdir(parents=True, exist_ok=True)

JSON_FILE = CURRENT / "dashboard-data.json"
STATUS_FILE = CURRENT / "status.json"
WORKBOOK_FILE = CURRENT / "Consolidated_Ocean_DSR.xlsx"
OUTPUT_SHEET_NAME = "Consolidated DSR"

HOST = os.environ.get("OCEAN_HOST", "127.0.0.1")
PORT = int(os.environ.get("OCEAN_PORT", "8050"))
MAX_UPLOAD_MB = int(os.environ.get("OCEAN_MAX_UPLOAD_MB", "80"))

# Fallback workbook from the existing automation pipeline (read-only).
LEGACY_WORKBOOK = Path(os.environ.get(
    "OCEAN_LEGACY_WORKBOOK",
    r"C:\Users\cshah\OneDrive - Virginia Transformer Corp\Project\ocean\automation\Output\Consolidated_Ocean_DSR.xlsx",
))

PUBLIC_FIELDS = (
    "Container Number",
    "PO#",
    "Delivery Location",
    "Vessel Name",
    "Forwarder",
    "Port of Loading",
    "Port Of Discharge",
    "Hotlist",
    "Class",
    "Shipper Name",
    "Steamship Line",
    "CTR SIZE / LCL",
    "Booked ETA Port",
    "Dlv Date",
    "Estimated Time of Departure",
    "Vessel Departed",
    "ETA To Port of Discharge",
    "ETA Door",
    "Vessel Arrived",
    "Container Delivered",
    "Current Status",
    "MOT",
)

DATE_FIELDS = {
    "Booked ETA Port",
    "Dlv Date",
    "Estimated Time of Departure",
    "Vessel Departed",
    "ETA To Port of Discharge",
    "ETA Door",
    "Vessel Arrived",
    "Container Delivered",
}
