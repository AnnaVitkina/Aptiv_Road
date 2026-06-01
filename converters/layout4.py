"""Convert layout4 simple pricelists (lane-style, grid-style, or corridor matrix)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from common import (
    add_metadata,
    is_blank,
    normalize_header,
    read_sheet_rows,
    sheets_to_convert,
    should_skip_tab,
)

from converters.layout1 import _detect_lane_structure, _parse_lane_sheet
from converters.layout3 import _parse_grid_sheet

SKIP_LAYOUT4 = re.compile(
    r"tt\s+details|accessorial|accesorial|fsc|revision",
    re.I,
)


def _is_price_tab(name: str) -> bool:
    if SKIP_LAYOUT4.search(name):
        return False
    if should_skip_tab(name, "layout4"):
        return False
    lower = name.strip().lower()
    if any(
        x in lower
        for x in (
            "ftl",
            "ltl",
            "pallet",
            "ldm",
            "milkrun",
            "xd_",
            "x-dock",
            "corbas",
            "groupage",
        )
    ):
        return True
    return False


def _parse_corridor_matrix(rows: list[list[Any]]) -> pd.DataFrame:
    """GEFCO TN style: Origin, Destination, trailer price columns."""
    header_idx = None
    for idx in range(min(8, len(rows))):
        text = " ".join(normalize_header(c) for c in rows[idx])
        if "origin" in text and "destination" in text:
            header_idx = idx
            break
    if header_idx is None:
        return pd.DataFrame()

    header = rows[header_idx]
    sub = rows[header_idx + 1] if header_idx + 1 < len(rows) else []
    price_cols: list[tuple[int, str]] = []
    for i in range(len(header)):
        label = str(sub[i]).strip() if i < len(sub) and not is_blank(sub[i]) else ""
        if not label:
            label = str(header[i]).strip() if not is_blank(header[i]) else ""
        top = normalize_header(header[i])
        if i >= 5 and (label or top not in ("", "origin", "destination")):
            if top not in ("origin", "origin country", "destination", "dest. country", "nb of trans"):
                price_cols.append((i, label or top))

    records: list[dict[str, Any]] = []
    for row_idx, row in enumerate(rows[header_idx + 2 :], start=header_idx + 3):
        if is_blank(row[0] if row else None):
            continue
        if normalize_header(row[0]).startswith("please"):
            break
        base = {
            "origin": row[0] if len(row) > 0 else None,
            "origin_country": row[1] if len(row) > 1 else None,
            "destination": row[2] if len(row) > 2 else None,
            "dest_country": row[3] if len(row) > 3 else None,
            "nb_trans_per_year": row[4] if len(row) > 4 else None,
            "row_number": row_idx,
        }
        for col_i, col_label in price_cols:
            price = row[col_i] if col_i < len(row) else None
            if is_blank(price):
                continue
            records.append({**base, "price_column": col_label, "price": price})

    return pd.DataFrame(records)


def _parse_sheet(rows: list[list[Any]], sheet_name: str) -> pd.DataFrame:
    # Rate grid style (unlikely in layout4 but safe)
    for idx in range(min(5, len(rows))):
        text = " ".join(normalize_header(c) for c in rows[idx])
        if "carrier name" in text:
            return _parse_grid_sheet(rows)

    # Corridor matrix (GEFCO TN)
    r0 = " ".join(normalize_header(c) for c in rows[0]) if rows else ""
    if "trailer" in r0 or (
        rows
        and normalize_header(rows[0][0] if rows[0] else "") == ""
        and any("origin" in normalize_header(c) for c in rows[1] if rows[1:])
    ):
        df = _parse_corridor_matrix(rows)
        if not df.empty:
            return df

    # Lane ratebook style (Aptiv, Fulop, EMONS, WMT Delphi, etc.)
    if _detect_lane_structure(rows):
        return _parse_lane_sheet(rows, sheet_name)

    return pd.DataFrame()


def convert_file(path: Path, sheets: list[str] | None = None) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in sheets_to_convert(path, sheets=sheets, auto_include=_is_price_tab):
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        df = _parse_sheet(rows, sheet)
        if df.empty:
            continue
        frames.append(
            add_metadata(
                df,
                source_file=path.name,
                layout="layout4",
                sheet_name=sheet,
            )
        )
    if not frames:
        return pd.DataFrame()
    from common import reorder_converted_df
    from converters.layout2 import LAYOUT1_DATA_COLUMNS

    return reorder_converted_df(pd.concat(frames, ignore_index=True), LAYOUT1_DATA_COLUMNS)
