#!/usr/bin/env python3
"""
End-to-end Aptiv Road pipeline: convert ratebooks → matrix Excel.

Two modes:
  - **Interactive** (like convert.py): pick layout, file, tabs — then export matrix.
  - **Batch**: process all files in input/ automatically.

Local:
    python pipeline.py -i              # interactive (prompts)
    python pipeline.py                 # batch all files
    python pipeline.py --layout layout1

Colab (recommended):
    import sys
    sys.path.insert(0, "/content/Aptiv_Road")
    from pipeline import setup_environment, run_interactive
    setup_environment()
    run_interactive()
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path


def _bootstrap_project_path() -> Path | None:
    """Find project root (folder with config.py) and add it to sys.path."""
    candidates: list[Path] = []
    env_root = os.environ.get("APTIV_ROAD_ROOT")
    if env_root:
        candidates.append(Path(env_root))
    try:
        candidates.append(Path(__file__).resolve().parent)
    except NameError:
        pass
    candidates.extend(
        [
            Path("/content/Aptiv_Road"),
            Path("/content/drive/MyDrive/Aptiv_Road"),
            Path.cwd(),
        ]
    )
    seen: set[str] = set()
    for root in candidates:
        key = str(root.resolve()) if root.exists() else str(root)
        if key in seen:
            continue
        seen.add(key)
        if (root / "config.py").is_file():
            root_s = str(root.resolve())
            if root_s not in sys.path:
                sys.path.insert(0, root_s)
            return root.resolve()
    return None


_bootstrap_project_path()

import config

if str(config.PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(config.PROJECT_ROOT))

from common import get_sheet_names, output_path, save_dataframe
from converters import CONVERTERS
from export_matrix import export_converted_file_to_matrix


def _prompt(message: str) -> str:
    try:
        return input(message).strip()
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        raise SystemExit(0) from None


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
    sheet_configs: dict | None = None,
) -> Path | None:
    """Convert one workbook to long-format *_converted.xlsx under processing/."""
    converter = CONVERTERS[layout]
    kwargs: dict = {}
    if layout == "layout1" and sheet_configs:
        kwargs["sheet_configs"] = sheet_configs
    try:
        df = converter(source, sheets=sheets, **kwargs)
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


def _parse_tab_selection(raw: str, sheet_names: list[str]) -> list[str] | None:
    text = raw.strip().lower()
    if not text or text in ("0", "a", "auto", "all"):
        return None
    if re.fullmatch(r"[\d,\s]+", text):
        indices: list[int] = []
        for part in re.split(r"[,]+", text):
            part = part.strip()
            if not part:
                continue
            if not part.isdigit():
                raise ValueError(f"Invalid tab number: {part}")
            indices.append(int(part))
        resolved: list[str] = []
        for idx in indices:
            if idx < 1 or idx > len(sheet_names):
                raise ValueError(f"Tab number {idx} out of range (1–{len(sheet_names)})")
            name = sheet_names[idx - 1]
            if name not in resolved:
                resolved.append(name)
        return resolved
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    by_lower = {n.strip().lower(): n for n in sheet_names}
    resolved = []
    for part in parts:
        key = part.strip().lower()
        if key in by_lower:
            resolved.append(by_lower[key])
            continue
        matches = [n for n in sheet_names if key in n.strip().lower()]
        if len(matches) == 1:
            resolved.append(matches[0])
        elif len(matches) > 1:
            raise ValueError(f"Ambiguous tab '{part}'. Matches: {', '.join(matches)}")
        else:
            raise ValueError(f"Unknown tab: {part}")
    return resolved


def _choose_layout() -> str | None:
    print("\nAvailable layout folders:")
    for i, name in enumerate(config.LAYOUTS, 1):
        folder = config.INPUT_DIR / name
        count = len(_list_excel_files(name)) if folder.is_dir() else 0
        print(f"  {i}. {name} ({count} files)")
    print("  0. Exit")
    while True:
        choice = _prompt("\nSelect layout number: ")
        if choice == "0":
            return None
        if choice.isdigit() and 1 <= int(choice) <= len(config.LAYOUTS):
            return config.LAYOUTS[int(choice) - 1]
        print("Invalid choice. Enter a number from the list.")


def _choose_files(files: list[Path]) -> list[Path]:
    print(f"\nFiles in input/{files[0].parent.name if files else ''}:")
    for i, path in enumerate(files, 1):
        print(f"  {i}. {path.name}")
    print(f"  {len(files) + 1}. ALL files in this layout (auto price tabs only)")
    print("  0. Back / exit")
    while True:
        choice = _prompt("\nSelect file number: ")
        if choice == "0":
            return []
        if choice.isdigit():
            n = int(choice)
            if n == len(files) + 1:
                return files
            if 1 <= n <= len(files):
                return [files[n - 1]]
        print("Invalid choice.")


def _choose_sheets(path: Path) -> list[str] | None:
    try:
        sheet_names = get_sheet_names(path)
    except Exception as exc:
        print(f"  ERROR reading workbook: {exc}")
        return None
    if not sheet_names:
        print("  No sheets found in workbook.")
        return []
    print(f"\nTabs in {path.name}:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    print("  0 / auto — auto-detect price tabs only")
    print("\nEnter tab number(s) or name(s), comma-separated (e.g. 3,4 or FTL,LTL):")
    while True:
        raw = _prompt("Tabs to convert: ")
        try:
            return _parse_tab_selection(raw, sheet_names)
        except ValueError as exc:
            print(f"  {exc}")
            retry = _prompt("Try again? [Y/n]: ").lower()
            if retry in ("n", "no"):
                return None


def run_interactive() -> PipelineResult:
    """
    Interactive pipeline: choose layout → file → tabs → convert → export matrix.
    Same prompts as convert.py, plus optional matrix export step.
    """
    result = PipelineResult()
    print("=" * 60)
    print("  Aptiv Road — interactive pipeline")
    print("=" * 60)
    print(f"  Input:      {config.INPUT_DIR}")
    print(f"  Processing: {config.PROCESSING_DIR}")
    print(f"  Output:     {config.OUTPUT_DIR}")

    while True:
        layout = _choose_layout()
        if layout is None:
            break

        files = _list_excel_files(layout)
        if not files:
            print(f"No Excel files found in input/{layout}/")
            continue

        selected = _choose_files(files)
        if not selected:
            continue

        for path in selected:
            sheets: list[str] | None = None
            sheet_configs = None
            if len(selected) == 1:
                sheets = _choose_sheets(path)
                if sheets == []:
                    continue
                if layout == "layout1":
                    from converters.layout1 import gather_column_overrides

                    sheet_configs = gather_column_overrides(
                        path, sheets, input_fn=_prompt
                    )

            dest = convert_one(
                layout, path, sheets=sheets, sheet_configs=sheet_configs
            )
            if dest is None:
                result.convert_skipped.append(str(path))
                continue
            result.converted.append(dest)

            export_now = _prompt("\nExport matrix Excel for this file? [Y/n]: ").lower()
            if export_now not in ("n", "no"):
                out = export_one(dest)
                if out is None:
                    result.export_errors.append(str(dest))
                else:
                    result.matrices.append(out)

        print(f"\nDone this round: {len(result.converted)} converted, {len(result.matrices)} matrices.")
        again = _prompt("\nProcess another file? [y/N]: ").lower()
        if again not in ("y", "yes"):
            break

    print("\nGoodbye.")
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
    parser.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        help="Prompt for layout, file, and tabs (like convert.py).",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Process all input files without prompts (Colab default without -i).",
    )
    if argv is None:
        # Colab/Jupyter passes `-f kernel-*.json`; ignore unknown args.
        args, _ = parser.parse_known_args()
    else:
        args = parser.parse_args(argv)
    return args


def _in_notebook() -> bool:
    try:
        get_ipython()  # type: ignore[name-defined]
        return True
    except NameError:
        return False


def _use_interactive_mode(args: argparse.Namespace, argv: list[str] | None) -> bool:
    if args.interactive:
        return True
    if args.batch or args.layouts or args.files or args.export_only or args.convert_only:
        return False
    # Colab notebook / script with no filters → ask which file to convert.
    return config.IS_COLAB and argv is None


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)

    if args.install:
        install_dependencies()
        return 0

    if args.setup or config.IS_COLAB:
        drive_mounted = Path("/content/drive/MyDrive").exists() or Path(
            "/content/drive/Shareddrives"
        ).exists()
        setup_environment(mount_drive=not drive_mounted)

    if _use_interactive_mode(args, argv):
        run_interactive()
        return 0

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
    exit_code = main()
    if not _in_notebook():
        raise SystemExit(exit_code)
