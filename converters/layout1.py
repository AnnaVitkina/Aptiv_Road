"""Convert layout1 lane ratebooks to a normalized DataFrame."""

from __future__ import annotations

import re
from pathlib import Path
from collections.abc import Callable
from typing import Any, TypedDict

import pandas as pd

from common import (
    NON_PRICE_COLUMN_RE,
    add_metadata,
    forward_fill_row_labels,
    is_blank,
    normalize_header,
    read_sheet_rows,
    sheets_to_convert,
    should_skip_tab,
)

LANE_HEADER_KEYS = ["origin", "zip"]
LANE_ID_KEYS = ["lane id", "lane #", "lane number"]


def _is_price_tab(name: str) -> bool:
    if should_skip_tab(name, "layout1"):
        return False
    lower = name.strip().lower()
    if any(
        x in lower
        for x in (
            "ftl",
            "ltl",
            "pallet",
            "ldm",
            "domestic",
            "shuttle",
            "courier",
            "octabin",
        )
    ):
        return True
    # Route-specific tabs (e.g. DE-51_ES-08)
    if re.search(r"[A-Z]{2}[-_][A-Z0-9]", name, re.I):
        return True
    return False


def _detect_lane_structure(
    rows: list[list[Any]], column_overrides: dict[str, int] | None = None
) -> dict[str, Any] | None:
    """Detect lane matrix header block (Origin/Destination + ZIP row)."""
    header_idx: int | None = None
    subheader_idx: int | None = None

    for idx in range(min(40, len(rows))):
        row_text = " ".join(normalize_header(c) for c in rows[idx])
        if not (
            "lane id" in row_text
            or "lane #" in row_text
            or "lane number" in row_text
        ):
            continue
        subheader_idx = idx
        prev_text = (
            " ".join(normalize_header(c) for c in rows[idx - 1]) if idx > 0 else ""
        )
        if "origin" in prev_text or "destination" in prev_text:
            header_idx = idx - 1
        else:
            header_idx = idx
        break

    if header_idx is None:
        # Fallback: combined Origin + ZIP in adjacent rows (avoid ISO CODE false positive)
        for idx in range(min(40, len(rows))):
            block = " ".join(
                normalize_header(cell)
                for row in rows[idx : idx + 2]
                for cell in row
            )
            if "origin" in block and re.search(r"zip", block):
                header_idx = idx
                subheader_idx = idx + 1
                break

    if header_idx is None:
        return None

    header = rows[header_idx]
    subheader = rows[subheader_idx] if subheader_idx is not None and subheader_idx < len(rows) else []

    lane_col = col_index_by_keywords(subheader, "lane id", "lane #", "lane number")
    if lane_col is None:
        lane_col = col_index_by_keywords(header, "lane id", "lane #", "lane number")

    # Field row with ZIP/City may sit below group header (e.g. DSV Poland)
    base_idx = subheader_idx if subheader_idx is not None else header_idx
    field_row_idx = base_idx
    for r in range(base_idx, min(base_idx + 6, len(rows))):
        row_text = " ".join(normalize_header(c) for c in rows[r])
        if re.search(r"\bzip\b", row_text) or "zip code" in row_text:
            field_row_idx = r
            break
    field_row = rows[field_row_idx] if field_row_idx < len(rows) else subheader

    # Row with MINIMUM / FLAT RATE / Van (below lane-id row; e.g. DSV row 5)
    vehicle_header_idx = header_idx
    vehicle_keywords_row = (
        "minimum",
        "flat rate",
        "min price",
        "van",
        "truck",
        "jumbo",
        "mega",
        "per 100",
        "one way price",
        "round trip price",
        "0,5 t",
        "standard 3",
    )
    for r in range(base_idx + 1, min(base_idx + 5, len(rows))):
        row_text = " ".join(normalize_header(c) for c in rows[r])
        if any(k in row_text for k in vehicle_keywords_row):
            vehicle_header_idx = r
            break
    vehicle_header = (
        rows[vehicle_header_idx] if vehicle_header_idx < len(rows) else header
    )

    def col_in_rows(*keywords: str, rows_to_scan: tuple[list[Any], ...]) -> int | None:
        for row in rows_to_scan:
            for kw in keywords:
                for i, cell in enumerate(row):
                    if kw in normalize_header(cell):
                        return i
        return None

    def col_lane_description() -> int | None:
        for row in (field_row, subheader, vehicle_header, header):
            for i, cell in enumerate(row):
                t = normalize_header(cell)
                if t in ("lane description", "lane desc"):
                    return i
                if "description" in t and "lane" in t and "lane id" not in t:
                    return i
        return None

    def col_row_description(exclude_col: int | None) -> int | None:
        """Row-type label column (e.g. Rate / Maut / Total Cost), not Lane Description."""
        for row in (field_row, subheader, vehicle_header, header):
            for i, cell in enumerate(row):
                if exclude_col is not None and i == exclude_col:
                    continue
                t = normalize_header(cell)
                if t == "description" or t == "cost":
                    return i
        return None

    def col_lane_route_code() -> int | None:
        """Route code column (e.g. DE86 - MK10), labeled LANE ID beside numeric Lane ID."""
        sub_idx_local = subheader_idx if subheader_idx is not None else header_idx
        row_indices = [sub_idx_local] + list(
            range(sub_idx_local + 1, min(sub_idx_local + 3, len(rows)))
        )
        for r in row_indices:
            if r < 0 or r >= len(rows):
                continue
            for i, cell in enumerate(rows[r]):
                t = normalize_header(cell)
                if t in ("lane id", "lane #") and i != lane_col:
                    return i
        return None

    lane_route_col = col_lane_route_code()
    lane_desc_col = col_lane_description() or lane_route_col
    row_desc_col = col_row_description(lane_desc_col)

    mapping = {
        "header_idx": header_idx,
        "lane_id": lane_col,
        "lane_route_col": lane_route_col,
        "origin_zip": None,
        "origin_city": None,
        "origin_country": None,
        "origin_name": None,
        "dest_zip": None,
        "dest_city": None,
        "dest_country": None,
        "dest_name": None,
        "lane_description": lane_desc_col,
        "description": row_desc_col,
        "currency": col_in_rows(
            "rate currency", "currency (iso", rows_to_scan=(field_row, subheader, header)
        ),
        "paid_by": col_in_rows("paid by", rows_to_scan=(field_row, subheader, header)),
        "data_start": field_row_idx + 1,
    }

    combined = field_row

    # Second ZIP/City/Country block = destination
    zip_cols = [
        i
        for i, c in enumerate(combined)
        if re.search(r"\bzip\b", normalize_header(c))
        or "zip code" in normalize_header(c)
    ]
    city_cols = [i for i, c in enumerate(combined) if "city" in normalize_header(c)]
    country_cols = [
        i for i, c in enumerate(combined) if "country" in normalize_header(c)
    ]
    supplier_cols = [
        i
        for i, c in enumerate(combined)
        if "supplier" in normalize_header(c) and "name" in normalize_header(c)
    ]
    if len(zip_cols) >= 2:
        mapping["origin_zip"] = zip_cols[0]
        mapping["dest_zip"] = zip_cols[1]
    if len(city_cols) >= 2:
        mapping["origin_city"] = city_cols[0]
        if len(city_cols) >= 3 and len(country_cols) == 1:
            # e.g. VPT: ZIP, City, Country (mislabeled City), ZIP, City, Country
            mapping["origin_country"] = city_cols[1]
            mapping["dest_city"] = city_cols[2]
            mapping["dest_country"] = country_cols[0]
        else:
            mapping["dest_city"] = city_cols[1]
    if len(country_cols) >= 2:
        mapping["origin_country"] = country_cols[0]
        mapping["dest_country"] = country_cols[1]
    else:
        from_col: int | None = None
        to_col: int | None = None
        for i, c in enumerate(combined):
            t = normalize_header(c)
            if t == "from":
                from_col = i
            elif t == "to":
                to_col = i
        if from_col is not None and to_col is not None:
            mapping["origin_country"] = from_col
            mapping["dest_country"] = to_col
    if len(supplier_cols) >= 2:
        mapping["origin_name"] = supplier_cols[0]
        mapping["dest_name"] = supplier_cols[1]
    else:
        for i, c in enumerate(combined):
            t = normalize_header(c)
            if "origin" in t and "name" in t and mapping.get("origin_name") is None:
                mapping["origin_name"] = i
            if "destination" in t and "name" in t and mapping.get("dest_name") is None:
                mapping["dest_name"] = i

    # Currency column — header may sit on lane row or above (e.g. AL7, DSV)
    currency_col = mapping.get("currency")
    scan_through = min(max(field_row_idx, vehicle_header_idx, header_idx) + 3, len(rows))
    for row_idx in range(0, scan_through):
        for i, c in enumerate(rows[row_idx]):
            t = normalize_header(c)
            if "rate currency" in t or t == "currency" or "currency (iso" in t:
                currency_col = i
                break
        if currency_col is not None:
            break
    mapping["currency"] = currency_col

    template_rows = tuple(
        rows[r] for r in range(base_idx, min(base_idx + 4, len(rows)))
    )
    cost_col = col_in_rows(
        "cost", rows_to_scan=(field_row, subheader, header, *template_rows)
    )
    if cost_col is not None and mapping.get("description") is None:
        mapping["description"] = cost_col

    # Price block starts after currency / lane fields
    price_start = None
    if mapping.get("currency") is not None:
        price_start = mapping["currency"] + 1
    if cost_col is not None and price_start is not None and price_start <= cost_col:
        price_start = cost_col + 1

    vehicle_row_labels = forward_fill_row_labels(vehicle_header)
    vehicle_keywords = (
        "van",
        "truck",
        "jumbo",
        "mega",
        "flat rate",
        "min price",
        "minimum",
        "one way price",
        "round trip price",
        "ldm",
        "pallet",
        "per 100",
        "0,5 t",
        "standard",
    )
    for i, label in enumerate(vehicle_row_labels):
        t = normalize_header(label)
        if any(k in t for k in vehicle_keywords):
            price_start = i if price_start is None else min(price_start, i)

    for scan in range(header_idx + 1, min(header_idx + 4, len(rows))):
        row = rows[scan]
        for i, cell in enumerate(row):
            t = normalize_header(cell)
            if any(
                x in t
                for x in ("minimum", "flat rate", "min price", "ltl standard")
            ):
                price_start = i if price_start is None else min(price_start, i)

    if cost_col is not None and price_start is not None and price_start <= cost_col:
        price_start = cost_col + 1

    if price_start is None:
        after_meta = mapping.get("lane_description") or 8
        if mapping.get("description") is not None:
            after_meta = max(after_meta, mapping["description"])
        price_start = after_meta + 1

    rate_labels = _build_rate_labels(
        rows,
        header_idx,
        subheader_idx,
        vehicle_header_idx,
        vehicle_header,
        subheader,
        price_start,
        mapping["data_start"],
    )

    mapping["price_start"] = price_start
    mapping["rate_labels"] = rate_labels
    mapping["subheader_idx"] = subheader_idx if subheader_idx is not None else header_idx
    mapping["vehicle_header_idx"] = vehicle_header_idx

    if column_overrides:
        for field, col in column_overrides.items():
            mapping[field] = col

    return mapping


WEIGHT_BRACKET_HEADER_RE = re.compile(
    r"\bkg\b|till\s*[\d.,]+\s*kg|\d[\d.,]*\s*kg\s*-\s*[\d.,]+\s*kg|^\d+[\d.,]*\s*-\s*\d",
    re.I,
)
TRANSIT_META_COLUMN_RE = re.compile(
    r"transit\s*time|departure\s+day|arrival|time\s+d\s*>?\s*d",
    re.I,
)
PALLET_INDEX_HEADER_RE = re.compile(r"^\d{1,4}$")
PALLET_PLT_HEADER_RE = re.compile(r"^\d+\s*plt\s*$", re.I)
GENERIC_RATE_COLUMN_NAMES = frozenset(
    {
        "flat rate",
        "minimum (flat rate)",
        "maximum (flat rate)",
        "ltl standard",
        "per 100 kg",
        "min price",
    }
)


def _looks_like_weight_bracket_header(text: str) -> bool:
    t = normalize_header(text)
    return bool(t and WEIGHT_BRACKET_HEADER_RE.search(t))


def _looks_like_pallet_column_header(text: str) -> bool:
    """Pallet-count or band index on field row (e.g. Lagermax cols 1…66, FTL)."""
    t = normalize_header(text)
    if not t:
        return False
    if t == "ftl":
        return True
    if PALLET_PLT_HEADER_RE.match(str(text).strip()):
        return True
    if PALLET_INDEX_HEADER_RE.match(t):
        return True
    if t in ("min. charge", "min charge", "minimum charge", "max price"):
        return True
    return False


def _is_transit_meta_column(text: str) -> bool:
    return bool(TRANSIT_META_COLUMN_RE.search(normalize_header(text)))


def _normalize_charge_type(charge_type: str, rate_group: str, rate_column: str) -> str:
    """Map vehicle/group headers to FLAT RATE / PER 100 kg / MIN / MAX charge labels."""
    ct = normalize_header(charge_type or "")
    rg = normalize_header(rate_group or "")
    rc = normalize_header(rate_column or "")
    if ct in ("minimum (flat rate)", "maximum (flat rate)", "min price", "max price"):
        return charge_type.strip() if charge_type else ""
    if ct == "flat rate":
        return charge_type.strip() if charge_type else "FLAT RATE"
    if "per 100" in ct and "kg" in ct:
        return "PER 100 kg"
    if "per 100" in rg or ("groupage" in rg and "100" in rg and "kg" in rg):
        return "PER 100 kg"
    if _looks_like_weight_bracket_header(rate_column) and (
        "groupage" in rg or "per 100" in rg
    ):
        return "PER 100 kg"
    if ct == "ltl standard":
        if rc in ("minimum (flat rate)", "maximum (flat rate)"):
            return rate_column.strip() if rate_column else ""
        return "FLAT RATE" if rc.endswith(" kg") else ""
    if re.search(r"\bt\b|\bcbm\b|jumbo|mega|standard\s+\d", ct):
        return charge_type.strip() if charge_type else ""
    return charge_type.strip() if charge_type else ""


def _field_row_rate_label(sub_txt: str, vehicle_label: str, rate_group: str) -> str | None:
    """Prefer field-row label over repeated FLAT RATE / vehicle row."""
    if _looks_like_weight_bracket_header(sub_txt) or _looks_like_pallet_column_header(sub_txt):
        return sub_txt
    vehicle_norm = normalize_header(vehicle_label)
    if vehicle_norm in GENERIC_RATE_COLUMN_NAMES and _label_is_price_column(sub_txt, rate_group):
        return sub_txt
    if not vehicle_label or not _label_is_price_column(vehicle_label, rate_group):
        if _label_is_price_column(sub_txt, rate_group):
            return sub_txt
    return None


META_RATE_GROUP_RE = re.compile(
    r"rate\s+currency|^\s*origin\s*$|^\s*destination\s*$|conversion:|cbm|ldm\s*=",
    re.I,
)


def _effective_rate_group(rate_group: str) -> str:
    """Drop forward-filled meta headers (currency row) mistaken for rate_group."""
    if not rate_group:
        return ""
    if META_RATE_GROUP_RE.search(rate_group) or NON_PRICE_COLUMN_RE.search(rate_group):
        return ""
    return rate_group


def _label_is_price_column(rate_column: str, rate_group: str) -> bool:
    if not rate_column:
        return False
    if NON_PRICE_COLUMN_RE.search(rate_column):
        return False
    rate_group = _effective_rate_group(rate_group)
    if rate_group and NON_PRICE_COLUMN_RE.search(rate_group):
        return False
    group_only = (
        "ftl one way",
        "ftl round trip",
        "ltl standard",
        "transit time",
        "rates without",
        "maut fee",
    )
    if normalize_header(rate_column) in group_only:
        return False
    if "maut" in normalize_header(rate_column):
        return False
    return True


def _weight_break_label(
    rows: list[list[Any]], col_i: int, sub_idx: int
) -> str | None:
    """Build e.g. '100 kg' from subheader lower bounds + template upper-bound row."""
    lower: str | None = None
    upper: str | None = None
    if sub_idx < len(rows) and col_i < len(rows[sub_idx]):
        val = rows[sub_idx][col_i]
        if not is_blank(val):
            text = str(val).strip()
            if text.replace(",", "").isdigit():
                lower = text.replace(",", "")
    for r in range(sub_idx + 1, min(sub_idx + 5, len(rows))):
        if col_i >= len(rows[r]):
            continue
        val = rows[r][col_i]
        if is_blank(val):
            continue
        text = str(val).strip()
        if text in ("-", "min price", "max price") or not re.match(
            r"^\d+$", text.replace(",", "")
        ):
            continue
        if lower is None:
            lower = text.replace(",", "")
        else:
            upper = text.replace(",", "")
    if lower and upper:
        return f"{upper} kg"
    if upper:
        return f"{upper} kg"
    if lower:
        return f"{lower} kg"
    return None


def _build_rate_labels(
    rows: list[list[Any]],
    header_idx: int,
    subheader_idx: int | None,
    vehicle_header_idx: int,
    vehicle_header: list[Any],
    subheader: list[Any],
    price_start: int,
    data_start: int,
) -> list[tuple[int, str, str]]:
    """Build (column_index, rate_column, rate_group) using multi-row headers."""
    sub_idx = subheader_idx if subheader_idx is not None else header_idx
    group_rows = [rows[i] for i in range(0, min(vehicle_header_idx, len(rows)))]
    super_filled = [forward_fill_row_labels(r) for r in group_rows]
    vehicle_filled = forward_fill_row_labels(vehicle_header)

    max_col = max((len(r) for r in rows), default=len(vehicle_header))

    rate_labels: list[tuple[int, str, str, str]] = []
    seen: set[int] = set()

    for col_i in range(price_start, max_col):
        rate_group = ""
        for sf in super_filled:
            if col_i < len(sf) and sf[col_i]:
                candidate = sf[col_i]
                if _effective_rate_group(candidate):
                    rate_group = candidate

        charge_type = ""
        if col_i < len(vehicle_filled) and not is_blank(vehicle_filled[col_i]):
            charge_type = str(vehicle_filled[col_i]).strip()

        sub_txt = ""
        if col_i < len(subheader) and not is_blank(subheader[col_i]):
            sub_txt = str(subheader[col_i]).strip()
        if sub_txt and _is_transit_meta_column(sub_txt):
            continue

        rate_column = ""
        sub_norm = normalize_header(sub_txt) if sub_txt else ""
        if sub_norm == "min price":
            rate_column = "MINIMUM (Flat rate)"
        elif sub_norm == "max price":
            rate_column = "MAXIMUM (Flat rate)"

        wb = _weight_break_label(rows, col_i, sub_idx)
        field_label = (
            _field_row_rate_label(sub_txt, charge_type, rate_group)
            if sub_txt and not rate_column
            else None
        )

        if not rate_column and wb:
            rate_column = wb
        elif not rate_column and field_label:
            rate_column = field_label
        elif not rate_column and charge_type:
            charge_norm = normalize_header(charge_type)
            if charge_norm not in ("flat rate", "per 100 kg", "ltl standard"):
                rate_column = charge_type

        if normalize_header(rate_column) == "cost":
            continue
        if not _label_is_price_column(rate_column, rate_group):
            continue
        if col_i in seen:
            continue
        seen.add(col_i)
        norm_charge = _normalize_charge_type(charge_type, rate_group, rate_column)
        rate_labels.append(
            (col_i, rate_column or rate_group, rate_group, norm_charge)
        )

    if not rate_labels:
        for col_i in range(price_start, len(vehicle_header)):
            label = normalize_header(vehicle_header[col_i])
            if label and _label_is_price_column(label, ""):
                charge = (
                    str(vehicle_header[col_i]).strip()
                    if col_i < len(vehicle_header) and not is_blank(vehicle_header[col_i])
                    else ""
                )
                rate_labels.append((col_i, label, "", charge))

    return rate_labels


def col_index_by_keywords(row: list[Any], *keywords: str) -> int | None:
    for idx, cell in enumerate(row):
        text = normalize_header(cell)
        for kw in keywords:
            if kw in text:
                return idx
    return None


def _cell(row: list[Any], idx: int | None) -> Any:
    if idx is None or idx >= len(row):
        return None
    return row[idx]


class SheetColumnConfig(TypedDict):
    overrides: dict[str, int]
    skip_columns: set[int]


MAPPABLE_LANE_FIELDS: tuple[tuple[str, str], ...] = (
    ("origin_zip", "Origin ZIP"),
    ("origin_city", "Origin city"),
    ("origin_country", "Origin country"),
    ("origin_name", "Origin name / supplier"),
    ("dest_zip", "Destination ZIP"),
    ("dest_city", "Destination city"),
    ("dest_country", "Destination country"),
    ("dest_name", "Destination name / supplier"),
    ("lane_description", "Lane Description"),
    ("description", "Row description (Rate / Maut)"),
    ("currency", "Currency"),
    ("paid_by", "Paid by"),
)


def _mapped_column_indices(spec: dict[str, Any]) -> set[int]:
    skip_keys = frozenset(
        {
            "header_idx",
            "data_start",
            "price_start",
            "subheader_idx",
            "vehicle_header_idx",
        }
    )
    mapped: set[int] = set()
    for key, val in spec.items():
        if key in skip_keys:
            continue
        if isinstance(val, int):
            mapped.add(val)
    for col_i, _, _, _ in spec.get("rate_labels", []):
        mapped.add(col_i)
    return mapped


def find_unmapped_lane_columns(
    rows: list[list[Any]],
    spec: dict[str, Any],
    skip_columns: set[int] | None = None,
) -> list[tuple[int, str, Any]]:
    """Columns with a header and sample data that are not mapped to a known field."""
    sub_idx = spec.get("subheader_idx") or spec["header_idx"]
    if sub_idx >= len(rows):
        return []
    field_row = rows[sub_idx]
    data_start = spec.get("data_start") or sub_idx + 1
    mapped = _mapped_column_indices(spec)
    skip_headers = {
        "lane id",
        "lane #",
        "lane number",
        "min price",
        "rate currency (iso code)",
        "rate currency",
        "ltl standard",
    }
    skipped = skip_columns or set()
    unmapped: list[tuple[int, str, Any]] = []
    for i, cell in enumerate(field_row):
        if i in mapped or i in skipped:
            continue
        header = str(cell).replace("\n", " ").strip() if cell is not None else ""
        if not header:
            continue
        t = normalize_header(header)
        if t in skip_headers or re.fullmatch(r"\d+", t):
            continue
        if _looks_like_weight_bracket_header(header) or _looks_like_pallet_column_header(
            header
        ):
            continue
        sample: Any = None
        for r in range(data_start, min(data_start + 20, len(rows))):
            val = _cell(rows[r], i)
            if not is_blank(val) and normalize_header(val) not in skip_headers:
                sample = val
                break
        if sample is None:
            continue
        unmapped.append((i, header, sample))
    return unmapped


def prompt_sheet_column_overrides(
    source_name: str,
    sheet_name: str,
    rows: list[list[Any]],
    *,
    input_fn: Callable[[str], str] = input,
) -> SheetColumnConfig:
    """Ask the user to map unmapped columns; skipped columns are excluded from output."""
    empty: SheetColumnConfig = {"overrides": {}, "skip_columns": set()}
    spec = _detect_lane_structure(rows)
    if spec is None:
        return empty

    overrides: dict[str, int] = {}
    skip_columns: set[int] = set()
    while True:
        spec = _detect_lane_structure(rows, column_overrides=overrides or None)
        if spec is None:
            break
        unmapped = find_unmapped_lane_columns(rows, spec, skip_columns)
        if not unmapped:
            break

        print(f"\n  Unmapped columns — {source_name} / '{sheet_name}':")
        for col_i, header, sample in unmapped:
            print(f"    Excel col {col_i + 1}: [{header}]  (example: {sample})")
        print("\n  Map a column to a field (number). Enter / 0 = skip column (omit from output):")
        for n, (field_key, field_label) in enumerate(MAPPABLE_LANE_FIELDS, 1):
            print(f"    {n}. {field_label} ({field_key})")

        for col_i, header, sample in unmapped:
            choice = input_fn(
                f"\n  Col {col_i + 1} [{header}] (e.g. {sample}) -> field number [0=skip]: "
            ).strip()
            if not choice or choice == "0":
                skip_columns.add(col_i)
                print("  -> skipped (not included)")
                continue
            if not choice.isdigit() or not (1 <= int(choice) <= len(MAPPABLE_LANE_FIELDS)):
                print("  Invalid choice, column skipped.")
                skip_columns.add(col_i)
                continue
            field_key = MAPPABLE_LANE_FIELDS[int(choice) - 1][0]
            overrides[field_key] = col_i
            print(f"  -> {field_key}")

    return {"overrides": overrides, "skip_columns": skip_columns}


def gather_column_overrides(
    path: Path,
    sheets: list[str] | None,
    *,
    input_fn: Callable[[str], str] = input,
) -> dict[str, SheetColumnConfig] | None:
    """Prompt for column mappings on sheets that still have unmapped columns."""
    per_sheet: dict[str, SheetColumnConfig] = {}
    for sheet in sheets_to_convert(path, sheets=sheets, auto_include=_is_price_tab):
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        spec = _detect_lane_structure(rows)
        if spec is None:
            continue
        if not find_unmapped_lane_columns(rows, spec):
            continue
        config = prompt_sheet_column_overrides(
            path.name, sheet, rows, input_fn=input_fn
        )
        if config["overrides"] or config["skip_columns"]:
            per_sheet[sheet] = config
    return per_sheet or None


def _parse_lane_sheet(
    rows: list[list[Any]],
    sheet_name: str,
    column_overrides: dict[str, int] | None = None,
    skip_columns: set[int] | None = None,
) -> pd.DataFrame:
    spec = _detect_lane_structure(rows, column_overrides=column_overrides)
    if spec is None:
        return pd.DataFrame()

    if skip_columns:
        spec["rate_labels"] = [
            label
            for label in spec.get("rate_labels", [])
            if label[0] not in skip_columns
        ]

    records: list[dict[str, Any]] = []
    data_start = spec.get("data_start") or (spec.get("subheader_idx") or spec["header_idx"]) + 1
    current_lane: dict[str, Any] = {}

    for row_idx in range(data_start, len(rows)):
        row = rows[row_idx]
        if not any(not is_blank(c) for c in row):
            continue

        lane_num = _cell(row, spec["lane_id"])
        lane_route = _cell(row, spec.get("lane_route_col"))
        lane_val = lane_route if not is_blank(lane_route) else lane_num
        origin_zip = _cell(row, spec["origin_zip"])
        dest_zip = _cell(row, spec["dest_zip"])

        header_tokens = ("lane id", "lane #", "lane number", "min price", "paid by")
        if isinstance(lane_val, str) and normalize_header(lane_val) in header_tokens:
            continue
        if isinstance(lane_num, str) and normalize_header(lane_num) in header_tokens:
            continue

        desc_col = spec.get("description")
        if desc_col is not None:
            desc_val = _cell(row, desc_col)
            if isinstance(desc_val, str):
                dv = normalize_header(desc_val)
                if dv in ("cost", "pallet(s) place"):
                    continue

        # Skip template rows: dash placeholders with no lane identifiers
        if spec.get("rate_labels") and is_blank(lane_val) and is_blank(lane_num):
            if is_blank(origin_zip) and is_blank(dest_zip):
                price_vals = [
                    _cell(row, col_i) for col_i, _, _, _ in spec["rate_labels"][:8]
                ]
                if price_vals and all(
                    is_blank(v) or str(v).strip() == "-" for v in price_vals
                ):
                    continue

        # Skip weight-break template rows (no lane, only dashes / break labels)
        if is_blank(lane_val) and is_blank(lane_num) and is_blank(origin_zip) and is_blank(dest_zip):
            if not current_lane:
                continue
            joined = " ".join(str(c) for c in row if not is_blank(c))
            if joined.replace("-", "").replace(".", "").isdigit() or not joined:
                continue

        # Update lane context when identifiers present
        has_route = (not is_blank(origin_zip) or not is_blank(_cell(row, spec["origin_city"]))) and (
            not is_blank(dest_zip) or not is_blank(_cell(row, spec["dest_city"]))
        )
        if not is_blank(lane_val) or not is_blank(lane_num) or has_route:
            current_lane = {
                "lane_id": lane_val,
                "lane_number": lane_num if lane_route and lane_num != lane_route else None,
                "paid_by": _cell(row, spec.get("paid_by")),
                "origin_name": _cell(row, spec.get("origin_name")),
                "origin_zip": origin_zip,
                "origin_city": _cell(row, spec["origin_city"]),
                "origin_country": _cell(row, spec["origin_country"]),
                "dest_name": _cell(row, spec.get("dest_name")),
                "dest_zip": dest_zip,
                "dest_city": _cell(row, spec["dest_city"]),
                "dest_country": _cell(row, spec["dest_country"]),
                "lane_description": _cell(row, spec["lane_description"]),
            }
            if not is_blank(_cell(row, spec["currency"])):
                current_lane["currency"] = _cell(row, spec["currency"])

        if not current_lane:
            continue

        currency = _cell(row, spec["currency"]) or current_lane.get("currency")
        row_description = _cell(row, spec.get("description"))

        # Cost component row (Base Price / Toll / Full Rate / Description column)
        cost_component = None
        if desc_col is not None:
            val = _cell(row, desc_col)
            if isinstance(val, str) and val.strip():
                t = normalize_header(val)
                if t not in ("cost", "pallet(s) place") and (
                    t in ("rate", "maut", "total cost", "total")
                    or "price" in t
                    or t in ("base price", "toll", "full rate")
                ):
                    cost_component = str(val).strip()
        if cost_component is None:
            for col_i in range(spec["price_start"] - 2, spec["price_start"] + 1):
                if col_i < 0 or col_i >= len(row):
                    continue
                val = row[col_i]
                if isinstance(val, str) and val.strip():
                    t = normalize_header(val)
                    if t in (
                        "base price",
                        "toll",
                        "full rate",
                        "cost",
                        "min price",
                    ) or "price" in t:
                        cost_component = str(val).strip()
                        break
        if cost_component is None and not is_blank(row_description):
            dv = normalize_header(str(row_description))
            if dv not in ("cost", "pallet(s) place"):
                cost_component = str(row_description).strip()

        for col_i, rate_label, rate_group, charge_type in spec["rate_labels"]:
            price = _cell(row, col_i)
            if is_blank(price) or str(price).strip() in ("-", ""):
                continue
            if isinstance(price, str) and normalize_header(price) in (
                "base price",
                "toll",
                "full rate",
            ):
                continue

            rec = {
                **current_lane,
                "currency": currency,
                "description": row_description,
                "cost_component": cost_component,
                "price_per": charge_type or None,
                "rate_column": rate_label,
                "rate_group": rate_group,
                "price": price,
                "row_number": row_idx + 1,
            }
            if not is_blank(lane_val):
                rec["lane_id"] = lane_val
            if not is_blank(lane_num) and lane_route:
                rec["lane_number"] = lane_num
            if not is_blank(origin_zip):
                rec["origin_zip"] = origin_zip
            records.append(rec)

    return pd.DataFrame(records)


def convert_file(
    path: Path,
    sheets: list[str] | None = None,
    sheet_configs: dict[str, SheetColumnConfig] | None = None,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for sheet in sheets_to_convert(path, sheets=sheets, auto_include=_is_price_tab):
        rows = read_sheet_rows(path, sheet, as_displayed=True)
        cfg = (sheet_configs or {}).get(sheet, {})
        df = _parse_lane_sheet(
            rows,
            sheet,
            column_overrides=cfg.get("overrides"),
            skip_columns=cfg.get("skip_columns"),
        )
        if df.empty:
            continue
        frames.append(
            add_metadata(
                df,
                source_file=path.name,
                layout="layout1",
                sheet_name=sheet,
            )
        )
    if not frames:
        return pd.DataFrame()
    from common import reorder_converted_df
    from converters.layout2 import LAYOUT1_DATA_COLUMNS

    return reorder_converted_df(pd.concat(frames, ignore_index=True), LAYOUT1_DATA_COLUMNS)
