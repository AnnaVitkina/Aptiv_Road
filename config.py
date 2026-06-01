"""Shared path configuration — auto-detects Google Colab vs local machine."""

from __future__ import annotations

from pathlib import Path

# --- Google Colab (Drive) paths — data folders on Shared Drive ---
COLAB_DRIVE_BASE = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Road"
)
COLAB_INPUT_DIR = f"{COLAB_DRIVE_BASE}/input"
COLAB_OUTPUT_DIR = f"{COLAB_DRIVE_BASE}/output"
COLAB_PROCESSING_DIR = f"{COLAB_DRIVE_BASE}/processing"

LAYOUTS: tuple[str, ...] = ("layout1", "layout2", "layout3", "layout4")


def _is_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


IS_COLAB = _is_colab()

_SCRIPT_ROOT = Path(__file__).resolve().parent
_DRIVE_INPUT = Path(COLAB_INPUT_DIR)

if IS_COLAB and _DRIVE_INPUT.is_dir():
    # Colab: code from GitHub clone; input/output/processing on Google Drive.
    PROJECT_ROOT = _SCRIPT_ROOT
    INPUT_DIR = _DRIVE_INPUT
    OUTPUT_DIR = Path(COLAB_OUTPUT_DIR)
    PROCESSING_DIR = Path(COLAB_PROCESSING_DIR)
else:
    PROJECT_ROOT = _SCRIPT_ROOT
    INPUT_DIR = PROJECT_ROOT / "input"
    OUTPUT_DIR = PROJECT_ROOT / "output"
    PROCESSING_DIR = PROJECT_ROOT / "processing"


def ensure_dirs() -> None:
    """Create processing/ and output/ trees (including per-layout subfolders)."""
    for base in (PROCESSING_DIR, OUTPUT_DIR):
        base.mkdir(parents=True, exist_ok=True)
        for layout in LAYOUTS:
            (base / layout).mkdir(parents=True, exist_ok=True)


def matrix_output_path(converted_path: Path) -> Path:
    """Mirror processing/<layout>/… under output/<layout>/… for matrix files."""
    stem = converted_path.stem.replace("_converted", "") + "_matrix.xlsx"
    try:
        rel = converted_path.resolve().relative_to(PROCESSING_DIR.resolve())
        if len(rel.parts) > 1:
            return OUTPUT_DIR / rel.parent / stem
    except ValueError:
        pass
    return OUTPUT_DIR / stem
