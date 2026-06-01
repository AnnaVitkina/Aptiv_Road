#!/usr/bin/env python3
"""
Interactive converter: select layout, file, and tab(s); export normalized rates to Excel.

Usage (from project root):
    python convert.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from common import get_sheet_names, output_path, save_dataframe
from converters import CONVERTERS

INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "processing"
LAYOUTS = ("layout1", "layout2", "layout3", "layout4")


def _prompt(message: str) -> str:
    try:
        return input(message).strip()
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        sys.exit(0)


def _choose_layout() -> str:
    print("\nAvailable layout folders:")
    for i, name in enumerate(LAYOUTS, 1):
        folder = INPUT_DIR / name
        count = len(list(folder.glob("*.xls*"))) if folder.is_dir() else 0
        print(f"  {i}. {name} ({count} files)")
    print("  0. Exit")

    while True:
        choice = _prompt("\nSelect layout number: ")
        if choice == "0":
            sys.exit(0)
        if choice.isdigit() and 1 <= int(choice) <= len(LAYOUTS):
            return LAYOUTS[int(choice) - 1]
        print("Invalid choice. Enter a number from the list.")


def _list_files(layout: str) -> list[Path]:
    folder = INPUT_DIR / layout
    if not folder.is_dir():
        return []
    return sorted(
        p for p in folder.iterdir() if p.suffix.lower() in (".xlsx", ".xls") and p.is_file()
    )


def _choose_file(files: list[Path]) -> list[Path]:
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


def _parse_tab_selection(raw: str, sheet_names: list[str]) -> list[str] | None:
    """
    Parse user tab selection.
    Returns None for auto-detect, or a list of resolved sheet names.
    """
    text = raw.strip().lower()
    if not text or text in ("0", "a", "auto", "all"):
        return None

    # Comma-separated 1-based indices
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
                raise ValueError(
                    f"Tab number {idx} out of range (1–{len(sheet_names)})"
                )
            name = sheet_names[idx - 1]
            if name not in resolved:
                resolved.append(name)
        return resolved

    # Comma-separated sheet names (partial match allowed)
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
            raise ValueError(
                f"Ambiguous tab '{part}'. Matches: {', '.join(matches)}"
            )
        else:
            raise ValueError(f"Unknown tab: {part}")
    return resolved


def _choose_sheets(path: Path) -> list[str] | None:
    """Prompt for one or more tabs. None = auto-detect price tabs."""
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


def _convert_one(
    layout: str,
    source: Path,
    sheets: list[str] | None,
    sheet_configs: dict | None = None,
) -> Path | None:
    converter = CONVERTERS[layout]
    tab_desc = "auto price tabs" if sheets is None else ", ".join(sheets)
    print(f"\nConverting: {source.name}")
    print(f"  Tabs: {tab_desc}")

    kwargs: dict = {}
    if layout == "layout1" and sheet_configs:
        kwargs["sheet_configs"] = sheet_configs

    try:
        df = converter(source, sheets=sheets, **kwargs)
    except ValueError as exc:
        print(f"  ERROR: {exc}")
        return None

    if df.empty:
        print("  WARNING: No price rows extracted.")
        return None

    dest = output_path(OUTPUT_DIR, layout, source)
    save_dataframe(df, dest)
    print(f"  Saved {len(df):,} rows -> {dest.relative_to(PROJECT_ROOT)}")
    return dest


def main() -> None:
    print("=" * 60)
    print("  Aptiv Road — rate layout converter")
    print("=" * 60)
    print(f"  Input:  {INPUT_DIR}")
    print(f"  Output: {OUTPUT_DIR}/<layout>/")

    while True:
        layout = _choose_layout()
        files = _list_files(layout)
        if not files:
            print(f"No Excel files found in input/{layout}/")
            continue

        selected = _choose_file(files)
        if not selected:
            continue

        ok = 0
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
            df_sheets = sheets
            if _convert_one(layout, path, df_sheets, sheet_configs):
                ok += 1

        print(f"\nDone: {ok}/{len(selected)} file(s) converted.")
        again = _prompt("\nConvert another? [y/N]: ").lower()
        if again not in ("y", "yes"):
            break

    print("Goodbye.")


if __name__ == "__main__":
    main()
