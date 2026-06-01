"""Shared path configuration — auto-detects Google Colab vs local machine."""

from __future__ import annotations

from pathlib import Path

# --- Google Colab (Drive) paths ---
COLAB_SCRIPTS_DIR = (
    "/content/Aptiv_Road"
)
COLAB_INPUT_DIR = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Road/input"
)
COLAB_OUTPUT_DIR = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Road/output"
)
COLAB_PROCESSING_DIR = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Road/processing"
)

LAYOUTS: tuple[str, ...] = ("layout1", "layout2", "layout3", "layout4")

def _is_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


IS_COLAB = _is_colab()

_SCRIPT_ROOT = Path(__file__).resolve().parent
_DRIVE_ROOT = Path(COLAB_SCRIPTS_DIR)

if IS_COLAB and _DRIVE_ROOT.is_dir() and (_DRIVE_ROOT / "config.py").is_file():
    # Scripts live on Google Drive (production Colab setup).
    PROJECT_ROOT = _DRIVE_ROOT
    INPUT_DIR = Path(COLAB_INPUT_DIR)
    OUTPUT_DIR = Path(COLAB_OUTPUT_DIR)
    PROCESSING_DIR = Path(COLAB_PROCESSING_DIR)
else:
    # Local machine, or Colab clone e.g. /content/Aptiv_Road from GitHub.
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

