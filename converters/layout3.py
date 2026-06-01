"""Convert layout3 rate grid tabs to layout1-compatible column schema."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from common import (
    NON_PRICE_COLUMN_RE,
    add_metadata,
    enrich_location_fields,
    format_cell_as_displayed,
    is_blank,
    is_roundtrip_rate_group,
    is_roundtrip_trip,
    is_truthy,
    normalize_header,
    read_sheet_rows,
    sheets_to_convert,
)

from converters.layout2 import LAYOUT1_DATA_COLUMNS

# Headers that are lane/meta fields, not price bands
_META_HEADER_MAP: tuple[tuple[str, str], ...] = (
    ("carrier name", "origin_name"),
    ("origin country", "origin_country"),
    ("origin zip", "origin_zip"),
    ("destination country", "dest_country"),
    ("destination zip", "dest_zip"),
    ("milkrun id", "lane_id"),
    ("origin", "origin_city"),
    ("destination", "dest_city"),
    ("currency", "currency"),
    ("price per", "price_per"),
    ("service", "_service"),
    ("tt in days", "_tt"),
    ("multistop equipment type", "equipment_type"),
    ("equipment type", "equipment_type"),
    ("roundtrip", "roundtrip"),
)

_SKIP_HEADERS = frozenset(
    {
        "mode",
        "trip type",
        "bid offered?",
        "edi 210?",
        "edi 214?",
        "track/trace?",
        "frequency",
        "stops",
        "#km",
        "volume max m3",
        "volume max ldm",
        "weight max.kg",
        "weight max kg",
    }
)

_STOP_COL_RE = re.compile(r"^stop\d+$", re.I)


def _is_price_tab(name: str) -> bool:
    return "rate grid" in name.lower()


def _base_record(**kwargs: Any) -> dict[str, Any]:
    rec = {col: None for col in LAYOUT1_DATA_COLUMNS}
    rec.update(kwargs)
    return rec


def _cell(row: list[Any], idx: int | None) -> Any:
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _field_for_header(header: str) -> str | None:
    t = normalize_header(header)
    for pattern, field in _META_HEADER_MAP:
        if pattern in t or t == pattern:
            return field
    return None


def _is_price_band_header(header: str) -> bool:
    t = normalize_header(header)
    if not t or t in _SKIP_HEADERS:
        return False
    if _field_for_header(header):
        return False
    if _STOP_COL_RE.match(t):
        return False
    if NON_PRICE_COLUMN_RE.search(t):
        return False
    if t == "price":
        return False
    return True


def _find_header_row(rows: list[list[Any]]) -> int | None:
    for idx in range(min(8, len(rows))):
        text = " ".join(normalize_header(c) for c in rows[idx])
        if "carrier" in text and ("origin" in text or "zip" in text or "milkrun" in text):
            return idx
    return None


def _parse_grid_sheet(rows: list[list[Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()

    header_idx = _find_header_row(rows)
    if header_idx is None:
        return pd.DataFrame()

    header_row = rows[header_idx]
    headers = [
        str(c).strip() if not is_blank(c) else ""
        for c in header_row
    ]

    meta_cols: dict[str, int] = {}
    extra: dict[str, int] = {}
    mode_col: int | None = None
    trip_col: int | None = None
    price_cols: list[tuple[int, str]] = []
    single_price_col: int | None = None

    for i, header in enumerate(headers):
        if not header:
            continue
        t = normalize_header(header)
        if t == "mode":
            mode_col = i
            continue
        if t == "trip type":
            trip_col = i
            continue
        field = _field_for_header(header)
        if field:
            if field.startswith("_"):
                extra[field] = i
            else:
                meta_cols[field] = i
            continue
        if t == "price":
            single_price_col = i
            continue
        if _is_price_band_header(header):
            price_cols.append((i, header))

    # Multi-stop: one Price column; equipment type is a separate field per row
    if single_price_col is not None and not price_cols:
        price_cols = [(single_price_col, "Price")]

    records: list[dict[str, Any]] = []
    for row_idx in range(header_idx + 1, len(rows)):
        row = rows[row_idx]
        if not any(not is_blank(c) for c in row):
            continue
        joined = " ".join(normalize_header(c) for c in row[:3])
        if joined.startswith("total") or "validation" in joined:
            break

        carrier = _cell(row, meta_cols.get("origin_name", 0))
        if is_blank(carrier) and is_blank(_cell(row, meta_cols.get("lane_id"))):
            continue
        if isinstance(carrier, str) and normalize_header(carrier) == "carrier name":
            continue

        mode_val = _cell(row, mode_col)
        trip_val = _cell(row, trip_col)
        rate_group_parts = [str(x).strip() for x in (mode_val, trip_val) if not is_blank(x)]
        rate_group = " | ".join(rate_group_parts)
        is_roundtrip = (
            is_truthy(_cell(row, meta_cols.get("roundtrip")))
            or is_roundtrip_trip(trip_val)
            or is_roundtrip_rate_group(rate_group)
        )

        lane_desc_parts: list[str] = []
        if extra.get("_service") is not None:
            v = _cell(row, extra["_service"])
            if not is_blank(v):
                lane_desc_parts.append(f"Service: {v}")
        if extra.get("_tt") is not None:
            v = _cell(row, extra["_tt"])
            if not is_blank(v):
                lane_desc_parts.append(f"TT: {v} days")

        # Multi-stop: route text in Origin / Stop* / Destination columns
        route_parts: list[str] = []
        for i, header in enumerate(headers):
            t = normalize_header(header)
            if t in ("origin", "destination") or _STOP_COL_RE.match(t):
                v = _cell(row, i)
                if not is_blank(v):
                    route_parts.append(f"{header}: {v}")
        if route_parts:
            lane_desc_parts.append(" | ".join(route_parts[:6]))

        lane_id_raw = _cell(row, meta_cols.get("lane_id"))
        if not is_blank(lane_id_raw):
            lane_id_raw = format_cell_as_displayed(lane_id_raw)

        origin_city, origin_zip, origin_country = enrich_location_fields(
            _cell(row, meta_cols.get("origin_city")),
            _cell(row, meta_cols.get("origin_zip")),
            _cell(row, meta_cols.get("origin_country")),
        )
        dest_city, dest_zip, dest_country = enrich_location_fields(
            _cell(row, meta_cols.get("dest_city")),
            _cell(row, meta_cols.get("dest_zip")),
            _cell(row, meta_cols.get("dest_country")),
        )

        lane_ctx = {
            "lane_id": lane_id_raw,
            "origin_name": _cell(row, meta_cols.get("origin_name")),
            "origin_zip": origin_zip,
            "origin_city": origin_city,
            "origin_country": origin_country,
            "dest_zip": dest_zip,
            "dest_city": dest_city,
            "dest_country": dest_country,
            "lane_description": " | ".join(lane_desc_parts) if lane_desc_parts else None,
            "currency": _cell(row, meta_cols.get("currency")),
            "price_per": _cell(row, meta_cols.get("price_per")),
            "equipment_type": _cell(row, meta_cols.get("equipment_type")),
            "roundtrip": is_roundtrip,
        }

        if not price_cols:
            continue
        for col_i, default_rate_label in price_cols:
            price = _cell(row, col_i)
            if is_blank(price) or str(price).strip() in ("-", ""):
                continue
            if isinstance(price, str) and normalize_header(price) in ("true", "false"):
                continue
            if is_roundtrip and len(price_cols) == 1:
                rate_label = "Roundtrip"
            else:
                rate_label = default_rate_label
            records.append(
                _base_record(
                    **lane_ctx,
                    rate_column=rate_label,
                    rate_group=rate_group or None,
                    price=price,
                    row_number=row_idx + 1,
                )
            )

    if not records:
        return pd.DataFrame(columns=list(LAYOUT1_DATA_COLUMNS))
    return pd.DataFrame(records)


def convert_file(path: Path, sheets: list[str] | None = None) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in sheets_to_convert(path, sheets=sheets, auto_include=_is_price_tab):
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        df = _parse_grid_sheet(rows)
        if df.empty:
            continue
        frames.append(
            add_metadata(
                df,
                source_file=path.name,
                layout="layout3",
                sheet_name=sheet,
            )
        )
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)
