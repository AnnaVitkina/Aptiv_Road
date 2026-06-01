#!/usr/bin/env python3
"""
End-to-end Aptiv Road pipeline: convert ratebooks → matrix Excel.

Runs non-interactively (batch). Suitable for local CLI and Google Colab.

Local:
    python pipeline.py
    python pipeline.py --layout layout1
    python pipeline.py --export-only

Colab (after copying this repo to COLAB_SCRIPTS_DIR on Drive):
    from pipeline import setup_environment, run_pipeline
    setup_environment()
    run_pipeline()
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

import config

# Ensure project modules resolve when launched from another cwd (e.g. Colab).
if str(config.PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(config.PROJECT_ROOT))

from common import output_path, save_dataframe
from converters import CONVERTERS
from export_matrix import export_converted_file_to_matrix


@dataclass
class PipelineResult:
    converted: list[Path] = field(default_factory=list)
    matrices: list[Path] = field(default_factory=list)
    convert_skipped: list[str] = field(default_factory=list)
    convert_errors: list[str] = field(default_factory=list)
    export_errors: list[str] = field(default_factory=list)


def setup_environment(*, mount_drive: bool | None = None) -> None:
    """
    Prepare Colab or local run: mount Drive (Colab), ensure folders exist.

    Call once at the top of a Colab notebook before run_pipeline().
    """
    if mount_drive is None:
        mount_drive = config.IS_COLAB
    if mount_drive:
        from google.colab import drive

        drive.mount("/content/drive", force_remount=False)
    if not config.PROJECT_ROOT.is_dir():
        raise FileNotFoundError(
            f"Scripts directory not found: {config.PROJECT_ROOT}\n"
            "On Colab, sync this repo to COLAB_SCRIPTS_DIR in config.py."
        )
    config.ensure_dirs()
    print(f"Environment: {'Colab' if config.IS_COLAB else 'local'}")
    print(f"  Scripts:    {config.PROJECT_ROOT}")
    print(f"  Input:      {config.INPUT_DIR}")
    print(f"  Processing: {config.PROCESSING_DIR}")
    print(f"  Output:     {config.OUTPUT_DIR}")


def install_dependencies() -> None:
    """Install pinned requirements (useful in a fresh Colab runtime)."""
    req = config.PROJECT_ROOT / "requirements.txt"
    if not req.is_file():
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q", "pandas>=2.0", "openpyxl>=3.1", "xlrd>=2.0"]
        )
        return
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "-r", str(req)])


def _list_excel_files(layout: str) -> list[Path]:
    folder = config.INPUT_DIR / layout
    if not folder.is_dir():
        return []
    return sorted(
        p
        for p in folder.iterdir()
        if p.suffix.lower() in (".xlsx", ".xls") and p.is_file() and not p.name.startswith("~$")
    )


def convert_one(
    layout: str,
    source: Path,
    *,
    sheets: list[str] | None = None,
) -> Path | None:
    """Convert one workbook to long-format *_converted.xlsx under processing/."""
    converter = CONVERTERS[layout]
    try:
        df = converter(source, sheets=sheets)
    except ValueError as exc:
        print(f"  ERROR converting {source.name}: {exc}")
        return None

    if df.empty:
        print(f"  WARNING: No price rows — {source.name}")
        return None

    dest = output_path(config.PROCESSING_DIR, layout, source)
    save_dataframe(df, dest)
    print(f"  Converted {source.name} -> {dest} ({len(df):,} rows)")
    return dest


def export_one(converted: Path) -> Path | None:
    """Export one *_converted.xlsx to matrix layout under output/."""
    out = config.matrix_output_path(converted)
    try:
        export_converted_file_to_matrix(converted, output_path=out)
    except Exception as exc:
        print(f"  ERROR exporting {converted.name}: {exc}")
        return None
    print(f"  Matrix {converted.name} -> {out}")
    return out


def find_converted_files(layouts: list[str] | None = None) -> list[Path]:
    layouts = layouts or list(config.LAYOUTS)
    files: list[Path] = []
    for layout in layouts:
        folder = config.PROCESSING_DIR / layout
        if not folder.is_dir():
            continue
        files.extend(
            sorted(
                p
                for p in folder.glob("*_converted.xlsx")
                if p.is_file() and not p.name.startswith("~$")
            )
        )
    return files


def run_pipeline(
    *,
    layouts: list[str] | None = None,
    files: list[Path] | None = None,
    sheets: list[str] | None = None,
    convert: bool = True,
    export: bool = True,
) -> PipelineResult:
    """
    Run convert and/or matrix export for all matching workbooks.

    Parameters
    ----------
    layouts:
        Subset of layout1–layout4. Default: all layouts in config.LAYOUTS.
    files:
        If set, only these input paths are converted (layout inferred from parent).
    sheets:
        Sheet names to convert; None = auto-detect price tabs per file.
    convert:
        When True, read input/*.xlsx and write processing/*_converted.xlsx.
    export:
        When True, read processing/*_converted.xlsx and write output/*_matrix.xlsx.
    """
    result = PipelineResult()
    layouts = layouts or list(config.LAYOUTS)

    if convert:
        print("\n--- Step 1: Convert ratebooks ---")
        for layout in layouts:
            if files is not None:
                sources = [p for p in files if p.parent.name == layout]
            else:
                sources = _list_excel_files(layout)
            if not sources:
                print(f"  [{layout}] no input files")
                continue
            print(f"  [{layout}] {len(sources)} file(s)")
            for source in sources:
                dest = convert_one(layout, source, sheets=sheets)
                if dest is None:
                    result.convert_skipped.append(str(source))
                else:
                    result.converted.append(dest)

    if export:
        print("\n--- Step 2: Export matrix Excel ---")
        converted_files = result.converted if convert else find_converted_files(layouts)
        if not converted_files:
            print("  No *_converted.xlsx files to export.")
        for converted in converted_files:
            out = export_one(converted)
            if out is None:
                result.export_errors.append(str(converted))
            else:
                result.matrices.append(out)

    print("\n--- Summary ---")
    print(f"  Converted: {len(result.converted)}")
    print(f"  Matrices:  {len(result.matrices)}")
    if result.convert_skipped:
        print(f"  Skipped:   {len(result.convert_skipped)}")
    if result.export_errors:
        print(f"  Export errors: {len(result.export_errors)}")
    return result


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Aptiv Road: convert ratebooks and export matrix Excel."
    )
    parser.add_argument(
        "--layout",
        action="append",
        dest="layouts",
        choices=config.LAYOUTS,
        help="Process only this layout (repeatable). Default: all layouts.",
    )
    parser.add_argument(
        "--file",
        action="append",
        dest="files",
        type=Path,
        help="Process only this input file (repeatable; path under input/<layout>/).",
    )
    parser.add_argument(
        "--sheets",
        help="Comma-separated sheet names (default: auto-detect price tabs).",
    )
    parser.add_argument(
        "--convert-only",
        action="store_true",
        help="Only run conversion step.",
    )
    parser.add_argument(
        "--export-only",
        action="store_true",
        help="Only export matrices from existing *_converted.xlsx files.",
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Mount Colab Drive and create folders, then exit.",
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="pip install requirements.txt, then exit.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    if args.install:
        install_dependencies()
        return 0

    if args.setup or config.IS_COLAB:
        setup_environment()

    sheets: list[str] | None = None
    if args.sheets:
        sheets = [s.strip() for s in args.sheets.split(",") if s.strip()]

    convert = not args.export_only
    export = not args.convert_only
    if args.setup and not convert and not export:
        return 0

    result = run_pipeline(
        layouts=args.layouts,
        files=args.files,
        sheets=sheets,
        convert=convert,
        export=export,
    )
    failed = len(result.convert_skipped) + len(result.export_errors)
    return 1 if failed and not result.converted and not result.matrices else 0


if __name__ == "__main__":
    raise SystemExit(main())
