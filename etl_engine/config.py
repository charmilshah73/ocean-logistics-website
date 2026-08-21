from pathlib import Path
import os

# Paths stay inside this website folder.
_ROOT = Path(__file__).resolve().parents[1]

INPUT_DIR = Path(os.environ.get("OCEAN_DSR_INPUT_DIR", _ROOT / "data" / "input"))
OUTPUT_DIR = Path(os.environ.get("OCEAN_DSR_OUTPUT_DIR", _ROOT / "data" / "output"))

OUTPUT_FILE = "Consolidated_Ocean_DSR.xlsx"
OUTPUT_SHEET_NAME = "Consolidated DSR"
OPEN_ORDER_SHEET_NAME = "Raw Data"
DATE_FORMAT = "mm/dd/yyyy"
CREATE_AUDIT_SHEETS = True
