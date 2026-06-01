"""Convert layout2 (WMT Romania Premium) to layout1-compatible column schema."""

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
)

from converters.layout1 import _detect_lane_structure, _parse_lane_sheet

PRICE_TABS = {
    "a) fixed rates",
    "b) per km rates",
    "c) ltl pallets per km",
    "e) tunisia -romania",
    "f) morocco - ldm",
}

LAYOUT1_DATA_COLUMNS = (
    "lane_id",
    "lane_number",
    "paid_by",
    "origin_name",
    "origin_zip",
    "origin_city",
    "origin_country",
    "dest_name",
    "dest_zip",
    "dest_city",
    "dest_country",
    "lane_description",
    "currency",
    "price_per",
    "description",
    "cost_component",
    "equipment_type",
    "roundtrip",
    "rate_column",
    "rate_group",
    "price",
    "row_number",
)


def _is_price_tab(name: str) -> bool:
    return name.strip().lower() in PRICE_TABS


def _base_record(**kwargs: Any) -> dict[str, Any]:
    rec = {col: None for col in LAYOUT1_DATA_COLUMNS}
    rec.update(kwargs)
    return rec


def _records_to_df(records: list[dict[str, Any]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=list(LAYOUT1_DATA_COLUMNS))
    return pd.DataFrame(records)


def _cell(row: list[Any], idx: int | None) -> Any:
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _clean_country(code: Any) -> str | None:
    if is_blank(code):
        return None
    text = re.sub(r"[^A-Za-z]", "", str(code).strip())
    return text.upper()[:2] if text else None


def _trip_label_for_col(col_i: int, trip_markers: list[tuple[int, str]]) -> str:
    label = ""
    for start, trip in sorted(trip_markers, key=lambda x: x[0]):
        if col_i >= start:
            label = trip
    return label


def _is_fixed_rates_sheet(rows: list[list[Any]]) -> bool:
    for row in rows[:20]:
        text = " ".join(normalize_header(c) for c in row)
        if "post code" in text and text.count("country") >= 2:
            return True
    return False


def _parse_fixed_rates_sheet(rows: list[list[Any]]) -> pd.DataFrame:
    """Lane-style tab: Origin/Destination country, post code, city + vehicle fixed prices."""
    field_idx: int | None = None
    for idx, row in enumerate(rows[:20]):
        text = " ".join(normalize_header(c) for c in row)
        if "post code" in text and "country" in text and "city" in text:
            field_idx = idx
            break
    if field_idx is None:
        return pd.DataFrame()

    field_row = rows[field_idx]
    vehicle_row = rows[field_idx - 1] if field_idx > 0 else field_row

    cols: dict[str, int] = {}
    country_n = postcode_n = city_n = 0
    for i, cell in enumerate(field_row):
        t = normalize_header(cell)
        if t == "country":
            country_n += 1
            cols["origin_country" if country_n == 1 else "dest_country"] = i
        elif "post code" in t:
            postcode_n += 1
            cols["origin_zip" if postcode_n == 1 else "dest_zip"] = i
        elif t == "city":
            city_n += 1
            cols["origin_city" if city_n == 1 else "dest_city"] = i
        elif t == "currency":
            cols["currency"] = i
        elif t in ("(km)", "km") or "distance" in t:
            cols["distance"] = i

    rate_group = ""
    for cell in vehicle_row:
        t = normalize_header(cell)
        if "fixed price" in t or "one way" in t:
            rate_group = str(cell).strip()
            break

    rate_labels: list[tuple[int, str]] = []
    for i, cell in enumerate(vehicle_row):
        if is_blank(cell):
            continue
        t = normalize_header(cell)
        if "vehicle" in t or ("pallet" in t and "fast" in t):
            rate_labels.append((i, str(cell).strip()))
    if not rate_labels:
        for i, cell in enumerate(field_row):
            if is_blank(cell):
                continue
            t = normalize_header(cell)
            if "vehicle" in t:
                rate_labels.append((i, str(cell).strip()))

    records: list[dict[str, Any]] = []
    for row_idx in range(field_idx + 1, len(rows)):
        row = rows[row_idx]
        if not any(not is_blank(c) for c in row):
            continue
        origin_country = _cell(row, cols.get("origin_country"))
        if is_blank(origin_country):
            continue
        if normalize_header(origin_country) in ("country", "origin"):
            continue

        currency = _cell(row, cols.get("currency"))
        distance = _cell(row, cols.get("distance"))
        lane_desc = f"{distance} km" if not is_blank(distance) else None

        lane_ctx = {
            "origin_zip": _cell(row, cols.get("origin_zip")),
            "origin_city": _cell(row, cols.get("origin_city")),
            "origin_country": origin_country,
            "dest_zip": _cell(row, cols.get("dest_zip")),
            "dest_city": _cell(row, cols.get("dest_city")),
            "dest_country": _cell(row, cols.get("dest_country")),
            "lane_description": lane_desc,
            "currency": currency,
        }

        for col_i, rate_label in rate_labels:
            price = _cell(row, col_i)
            if is_blank(price) or str(price).strip() in ("-", ""):
                continue
            records.append(
                _base_record(
                    **lane_ctx,
                    rate_column=rate_label,
                    rate_group=rate_group or "Fixed one way",
                    price=price,
                    row_number=row_idx + 1,
                )
            )

    return _records_to_df(records)


def _parse_matrix_sheet(rows: list[list[Any]]) -> pd.DataFrame:
    """Country × vehicle/pallet matrix (per-km and similar tabs)."""
    records: list[dict[str, Any]] = []
    section = ""
    active_blocks: list[dict[str, Any]] = []
    trip_markers: list[tuple[int, str]] = []

    for row_idx, row in enumerate(rows):
        if not any(not is_blank(c) for c in row):
            continue

        row_join = " ".join(normalize_header(c) for c in row if not is_blank(c))

        if "domestic transport" in row_join or "international transport" in row_join:
            section = row_join[:160]
            trip_markers = []
            active_blocks = []
            continue

        for i, cell in enumerate(row):
            t = normalize_header(cell)
            if "one way" in t or "round trip" in t:
                trip_markers.append((i, str(cell).strip()))

        is_band_row = any(
            kw in row_join
            for kw in ("vehicle", "pallet", "boxes", "loaded km", "ldm")
        ) and not re.match(r"^[a-z]{2}", normalize_header(row[0] if row else ""))

        if is_band_row and ("vehicle" in row_join or "pallet" in row_join or "boxes" in row_join):
            active_blocks = []
            current_bands: list[tuple[int, str]] = []
            block_currency: int | None = None
            block_trip = ""

            for i, cell in enumerate(row):
                if is_blank(cell):
                    continue
                t = normalize_header(cell)
                if t in ("origin/destination",):
                    continue
                if "currency" in t:
                    if current_bands:
                        trip = _trip_label_for_col(current_bands[0][0], trip_markers)
                        active_blocks.append(
                            {
                                "bands": current_bands,
                                "currency_col": block_currency if block_currency is not None else i,
                                "rate_group": " | ".join(p for p in (section, trip) if p),
                            }
                        )
                        current_bands = []
                    block_currency = i
                    block_trip = _trip_label_for_col(i + 1, trip_markers)
                    continue
                if i == 0:
                    continue
                label = str(cell).strip()
                if label:
                    current_bands.append((i, label))

            if current_bands:
                trip = _trip_label_for_col(current_bands[0][0], trip_markers)
                active_blocks.append(
                    {
                        "bands": current_bands,
                        "currency_col": block_currency,
                        "rate_group": " | ".join(p for p in (section, trip) if p),
                    }
                )
            continue

        country = _clean_country(row[0] if row else None)
        if not country or not active_blocks:
            continue
        if country in ("TO", "NOTE") or normalize_header(row[0]) in ("note:", "origin/destination"):
            continue

        is_domestic = country == "RO" or str(row[0]).strip().upper().startswith("RO")
        origin_country = "RO"
        dest_country = "RO" if is_domestic else country

        origin_hint = ""
        if "sannicolau" in section.lower():
            origin_hint = "Sannicolau Mare area"

        for block in active_blocks:
            currency = _cell(row, block.get("currency_col"))
            for col_i, band in block["bands"]:
                price = _cell(row, col_i)
                if is_blank(price) or str(price).strip() in ("-", ""):
                    continue
                records.append(
                    _base_record(
                        origin_country=origin_country,
                        origin_city=origin_hint or None,
                        dest_country=dest_country,
                        lane_description=section[:120] if section else None,
                        currency=currency,
                        rate_column=band,
                        rate_group=block.get("rate_group") or section,
                        price=price,
                        row_number=row_idx + 1,
                    )
                )

    return _records_to_df(records)


def _parse_tunisia_sheet(rows: list[list[Any]]) -> pd.DataFrame:
    """Tunisia–Romania corridor: cost rows × CBM bands."""
    records: list[dict[str, Any]] = []
    band_headers: list[tuple[int, str]] = []
    band_row_idx: int | None = None

    for row_idx, row in enumerate(rows):
        if not any(not is_blank(c) for c in row):
            continue
        row_join = " ".join(normalize_header(c) for c in row if not is_blank(c))

        if band_row_idx is None and "cbm" in row_join and ("0 to" in row_join or "min" in row_join):
            band_headers = []
            for i, cell in enumerate(row):
                if is_blank(cell) or i <= 1:
                    continue
                t = normalize_header(cell)
                if "transit" in t:
                    continue
                band_headers.append((i, str(cell).strip()))
            band_row_idx = row_idx
            continue

        if not band_headers:
            continue

        cost_label = _cell(row, 1)
        if is_blank(cost_label) or not isinstance(cost_label, str):
            continue
        cost_norm = normalize_header(cost_label)
        if cost_norm not in ("transport costs", "marpol/cbm"):
            continue

        for col_i, band in band_headers:
            price = _cell(row, col_i)
            if is_blank(price) or str(price).strip() in ("-", ""):
                continue
            if isinstance(price, str) and normalize_header(price) in ("eur", "days"):
                continue
            records.append(
                _base_record(
                    origin_country="TN",
                    dest_country="RO",
                    lane_description="Tunisia - Romania",
                    currency="EUR",
                    cost_component=str(cost_label).strip(),
                    rate_column=band,
                    rate_group="Tunisia - Romania LTL",
                    price=price,
                    row_number=row_idx + 1,
                )
            )

    return _records_to_df(records)


def _parse_sheet(rows: list[list[Any]], sheet_name: str) -> pd.DataFrame:
    name = sheet_name.strip().lower()
    if _is_fixed_rates_sheet(rows):
        return _parse_fixed_rates_sheet(rows)
    if "tunisia" in name:
        return _parse_tunisia_sheet(rows)
    if _detect_lane_structure(rows) is not None:
        return _parse_lane_sheet(rows, sheet_name)
    return _parse_matrix_sheet(rows)


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
                layout="layout2",
                sheet_name=sheet,
            )
        )
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)
