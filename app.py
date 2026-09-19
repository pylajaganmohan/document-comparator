import io
import json
import os
import re
import difflib
from pathlib import Path
from collections import defaultdict
from itertools import product

import pandas as pd
import streamlit as st

# Optional PDF dependency
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import pytesseract
except Exception:
    pytesseract = None

# pypdfium2 renders PDF pages without requiring Poppler on Windows.
try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None

# Optional legacy fallback renderer.
try:
    from pdf2image import convert_from_bytes
except Exception:
    convert_from_bytes = None


APP_TITLE = "Document Comparator"
SETUP_DIR = Path("saved_setups")
SETUP_DIR.mkdir(exist_ok=True)


# ============================================================
# Utility functions
# ============================================================

def normalize_value(value, ignore_spaces=False, ignore_case=False):
    if pd.isna(value):
        text = ""
    else:
        text = str(value).strip()

    if ignore_spaces:
        text = re.sub(r"\s+", "", text)
    if ignore_case:
        text = text.lower()

    return text


def make_key(row, columns, ignore_spaces=False, ignore_case=False):
    return tuple(
        normalize_value(row[col], ignore_spaces, ignore_case)
        for col in columns
    )


def display_key(key):
    if not isinstance(key, tuple):
        return str(key)
    if len(key) == 1:
        return str(key[0])
    return " | ".join(str(x) for x in key)


def read_excel_file(uploaded_file, sheet_name):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def get_excel_sheets(uploaded_file):
    uploaded_file.seek(0)
    return pd.ExcelFile(uploaded_file).sheet_names


def read_csv_file(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file)


def get_file_type(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return "csv"
    if name.endswith((".xlsx", ".xls", ".xlsm")):
        return "excel"
    if name.endswith(".pdf"):
        return "pdf"
    return "unknown"


def load_dataframe(uploaded_file, sheet_name=None):
    file_type = get_file_type(uploaded_file)

    if file_type == "csv":
        return read_csv_file(uploaded_file)

    if file_type == "excel":
        if sheet_name is None:
            sheet_name = get_excel_sheets(uploaded_file)[0]
        return read_excel_file(uploaded_file, sheet_name)

    raise ValueError("Only Excel and CSV files can be converted to a DataFrame.")


def make_groups(df, key_columns, ignore_spaces=False, ignore_case=False):
    groups = defaultdict(list)

    for idx, row in df.iterrows():
        key = make_key(
            row,
            key_columns,
            ignore_spaces=ignore_spaces,
            ignore_case=ignore_case,
        )
        # Excel/CSV row number: header = row 1, first data row = row 2
        excel_row = idx + 2
        groups[key].append((idx, excel_row, row))

    return groups


def row_signature(row, source_columns, mapping, ignore_spaces=False, ignore_case=False):
    """
    Signature contains ONLY compared source columns.
    Match keys are deliberately NOT used as the whole-row signature.
    """
    return tuple(
        normalize_value(
            row[source_col],
            ignore_spaces=ignore_spaces,
            ignore_case=ignore_case,
        )
        for source_col in source_columns
    )


def changed_fields(source_row, target_row, compare_columns, mapping,
                   ignore_spaces=False, ignore_case=False,
                   numeric_tolerance=0.0, normalize_dates=False,
                   column_rules=None, value_aliases=None,
                   default_timezone="Asia/Kolkata"):
    changes = []

    for source_col in compare_columns:
        target_col = mapping[source_col]

        source_value = source_row[source_col]
        target_value = target_row[target_col]

        source_norm = normalize_value(
            source_value, ignore_spaces=ignore_spaces, ignore_case=ignore_case
        )
        target_norm = normalize_value(
            target_value, ignore_spaces=ignore_spaces, ignore_case=ignore_case
        )

        rule = (column_rules or {}).get(source_col, {})
        rs = rule.get("ignore_spaces", ignore_spaces)
        rc = rule.get("ignore_case", ignore_case)
        rt = float(rule.get("tolerance", numeric_tolerance) or 0)
        rd = bool(rule.get("normalize_dates", normalize_dates))
        if not compare_values(
            source_value, target_value, rs, rc, rt, rd,
            (value_aliases or {}).get(source_col, []),
            default_timezone,
        ):
            changes.append(
                {
                    "Column": source_col,
                    "Target Column": target_col,
                    "Source Value": "" if pd.isna(source_value) else str(source_value),
                    "Target Value": "" if pd.isna(target_value) else str(target_value),
                }
            )

    return changes


def pair_duplicate_rows(source_rows, target_rows, compare_columns, mapping,
                        ignore_spaces=False, ignore_case=False,
                        value_aliases=None, default_timezone="Asia/Kolkata"):
    """
    Match repeated keys without depending on row position.

    Strategy:
    1. Exact row matches first.
    2. For remaining rows, pair source/target rows using the smallest
       number of changed compared columns.
    """
    unmatched_source = list(source_rows)
    unmatched_target = list(target_rows)
    pairs = []

    # Exact matches first.
    target_by_signature = defaultdict(list)

    for t in unmatched_target:
        _, _, target_row = t
        sig = tuple(
            normalize_value(
                target_row[mapping[col]],
                ignore_spaces=ignore_spaces,
                ignore_case=ignore_case,
            )
            for col in compare_columns
        )
        target_by_signature[sig].append(t)

    still_source = []

    for s in unmatched_source:
        _, _, source_row = s
        sig = row_signature(
            source_row,
            compare_columns,
            mapping,
            ignore_spaces,
            ignore_case,
        )

        if target_by_signature[sig]:
            t = target_by_signature[sig].pop(0)
            pairs.append((s, t))
        else:
            still_source.append(s)

    unmatched_source = still_source

    used_target_ids = set()
    for bucket in target_by_signature.values():
        for t in bucket:
            used_target_ids.add(id(t))

    # Rebuild remaining target rows from original list, excluding exact matches.
    exact_target_ids = {
        id(t) for _, t in pairs
    }
    unmatched_target = [
        t for t in target_rows if id(t) not in exact_target_ids
    ]

    # Pair remaining rows by minimum number of changed fields.
    while unmatched_source and unmatched_target:
        best = None

        for si, s in enumerate(unmatched_source):
            for ti, t in enumerate(unmatched_target):
                _, _, sr = s
                _, _, tr = t
                changes = changed_fields(
                    sr, tr, compare_columns, mapping,
                    ignore_spaces, ignore_case,
                    value_aliases=value_aliases,
                    default_timezone=default_timezone,
                )
                cost = len(changes)

                candidate = (cost, si, ti, changes)

                if best is None or candidate[0] < best[0]:
                    best = candidate

        _, si, ti, changes = best
        s = unmatched_source.pop(si)
        t = unmatched_target.pop(ti)
        pairs.append((s, t))

    return pairs, unmatched_source, unmatched_target


def suggest_column_mappings(source_columns, target_columns):
    suggestions = {}
    clean = lambda x: re.sub(r"[^a-z0-9]", "", str(x).lower())
    tc = {t: clean(t) for t in target_columns}
    for s in source_columns:
        cs = clean(s)
        best, score = None, 0
        for t, ct in tc.items():
            v = difflib.SequenceMatcher(None, cs, ct).ratio()
            if cs in ct or ct in cs:
                v = max(v, 0.82)
            if cs.replace("id", "") == ct.replace("id", ""):
                v = max(v, 0.92)
            if v > score:
                best, score = t, v
        if best:
            suggestions[s] = (best, round(score * 100))
    return suggestions


def numeric_equal(a, b, tolerance):
    try:
        if pd.isna(a) or pd.isna(b):
            return False
        x = float(str(a).replace(",", "").strip())
        y = float(str(b).replace(",", "").strip())
        return abs(x - y) <= tolerance
    except Exception:
        return False


def normalize_datetime_candidates(value, default_timezone="Asia/Kolkata"):
    """
    Return possible normalized datetime values in UTC.

    Timezone-aware values are converted to UTC. Values without a timezone
    are interpreted using the configured default timezone.

    This is important for comparisons such as:
      2026-09-01 14:30:00+00:00
      01/09/2026 20:00
    when the timezone-less target value is intended to be IST.
    """
    if pd.isna(value):
        return set()

    text = str(value).strip()
    if not text:
        return set()

    candidates = []

    formats = [
        "%Y-%m-%d", "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %I:%M %p", "%Y-%m-%d %I:%M:%S %p",
        "%Y/%m/%d", "%Y/%m/%d %H:%M", "%Y/%m/%d %H:%M:%S",
        "%d/%m/%Y", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %I:%M %p", "%d/%m/%Y %I:%M:%S %p",
        "%m/%d/%Y", "%m/%d/%Y %H:%M", "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M %p", "%m/%d/%Y %I:%M:%S %p",
        "%d-%m-%Y", "%d-%m-%Y %H:%M", "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %I:%M %p", "%d-%m-%Y %I:%M:%S %p",
        "%m-%d-%Y", "%m-%d-%Y %H:%M", "%m-%d-%Y %H:%M:%S",
        "%m-%d-%Y %I:%M %p", "%m-%d-%Y %I:%M:%S %p",
    ]

    # First try explicit formats.
    for fmt in formats:
        try:
            candidates.append((pd.to_datetime(text, format=fmt), False))
        except Exception:
            pass

    # General parser fallback.
    if not candidates:
        for dayfirst in (False, True):
            try:
                parsed = pd.to_datetime(text, errors="raise", dayfirst=dayfirst)
                candidates.append((parsed, False))
            except Exception:
                pass

    # ISO strings with offsets/timezones are parsed separately.
    try:
        parsed = pd.to_datetime(text, errors="raise")
        candidates.append((parsed, True if getattr(parsed, "tzinfo", None) else False))
    except Exception:
        pass

    result = set()

    for dt, parser_had_timezone_hint in candidates:
        try:
            ts = pd.Timestamp(dt)

            if ts.tzinfo is not None:
                # Already timezone-aware: compare the actual instant.
                utc_ts = ts.tz_convert("UTC").tz_localize(None)
            else:
                # No timezone in the source text: interpret it in the
                # user-selected default timezone before converting to UTC.
                localized = ts.tz_localize(
                    default_timezone,
                    ambiguous="NaT",
                    nonexistent="NaT",
                )
                if pd.isna(localized):
                    continue
                utc_ts = localized.tz_convert("UTC").tz_localize(None)

            result.add(utc_ts)
        except Exception:
            pass

    return result


def date_equal(a, b, default_timezone="Asia/Kolkata"):
    left = normalize_datetime_candidates(a, default_timezone)
    right = normalize_datetime_candidates(b, default_timezone)

    if not left or not right:
        return False

    return bool(left.intersection(right))


def normalize_datetime_value(value, default_timezone="Asia/Kolkata"):
    candidates = normalize_datetime_candidates(value, default_timezone)
    if len(candidates) == 1:
        return next(iter(candidates))
    return None


def compare_values(a, b, ignore_spaces=False, ignore_case=False,
                   tolerance=0.0, normalize_dates=False, value_aliases=None,
                   default_timezone="Asia/Kolkata"):
    if tolerance > 0 and numeric_equal(a, b, tolerance):
        return True
    if normalize_dates and date_equal(a, b, default_timezone):
        return True

    source_norm = normalize_value(a, ignore_spaces, ignore_case)
    target_norm = normalize_value(b, ignore_spaces, ignore_case)

    if source_norm == target_norm:
        return True

    for alias in value_aliases or []:
        alias_source = normalize_value(alias.get("source", ""), ignore_spaces, ignore_case)
        alias_target = normalize_value(alias.get("target", ""), ignore_spaces, ignore_case)
        if source_norm == alias_source and target_norm == alias_target:
            return True

    return False


def difference_by_column(result_df):
    d = result_df[result_df["Status"] == "Value Difference"]
    if d.empty:
        return pd.DataFrame(columns=["Column", "Differences"])
    return d.groupby("Column").size().reset_index(name="Differences").sort_values(
        "Differences", ascending=False
    )


def duplicate_key_report(
    source_df,
    target_df,
    key_columns,
    mapping=None,
    ignore_spaces=False,
    ignore_case=False,
):
    # Source and Target column names may differ, so the Target side must
    # use the mapped Target columns for the same logical Match Key.
    mapping = mapping or {}
    target_key_columns = [mapping.get(col, col) for col in key_columns]

    sg = make_groups(source_df, key_columns, ignore_spaces, ignore_case)
    tg = make_groups(target_df, target_key_columns, ignore_spaces, ignore_case)
    keys = list(dict.fromkeys(list(sg.keys()) + list(tg.keys())))
    rows = []
    for k in keys:
        sc, tc = len(sg.get(k, [])), len(tg.get(k, []))
        if sc > 1 or tc > 1:
            rows.append({
                "Match Key": display_key(k),
                "Source Count": sc,
                "Target Count": tc,
                "Status": "Balanced" if sc == tc else "Count Mismatch",
            })
    return pd.DataFrame(rows)



def compare_dataframes(
    source_df,
    target_df,
    mapping,
    key_source_columns,
    compare_columns,
    ignore_spaces=False,
    ignore_case=False,
    numeric_tolerance=0.0,
    normalize_dates=False,
    column_rules=None,
    value_aliases=None,
    default_timezone="Asia/Kolkata",
):
    """
    Correct comparison model:

    - Match Key = ONLY the selected key source columns mapped to target columns.
    - Rows are grouped by key.
    - Row order is ignored.
    - Repeating keys are supported.
    - Extra rows have Source Row blank.
    - Missing rows have Target Row blank.
    """
    results = []

    source_groups = make_groups(
        source_df,
        key_source_columns,
        ignore_spaces,
        ignore_case,
    )

    target_key_columns = [mapping[col] for col in key_source_columns]

    target_groups = make_groups(
        target_df,
        target_key_columns,
        ignore_spaces,
        ignore_case,
    )

    all_keys = list(dict.fromkeys(
        list(source_groups.keys()) + list(target_groups.keys())
    ))

    for key in all_keys:
        source_rows = source_groups.get(key, [])
        target_rows = target_groups.get(key, [])

        key_text = display_key(key)

        # Missing in Target
        if source_rows and not target_rows:
            for _, source_excel_row, source_row in source_rows:
                results.append(
                    {
                        "Status": "Missing in Target",
                        "Match Key": key_text,
                        "Source Row": source_excel_row,
                        "Target Row": "",
                        "Column": "",
                        "Target Column": "",
                        "Source Value": "",
                        "Target Value": "",
                    }
                )
            continue

        # Extra in Target
        if target_rows and not source_rows:
            for _, target_excel_row, target_row in target_rows:
                results.append(
                    {
                        "Status": "Extra in Target",
                        "Match Key": key_text,
                        "Source Row": "",
                        "Target Row": target_excel_row,
                        "Column": "",
                        "Target Column": "",
                        "Source Value": "",
                        "Target Value": "",
                    }
                )
            continue

        # Both sides have this key.
        pairs, unmatched_source, unmatched_target = pair_duplicate_rows(
            source_rows,
            target_rows,
            compare_columns,
            mapping,
            ignore_spaces,
            ignore_case,
            value_aliases=value_aliases,
            default_timezone=default_timezone,
        )

        # Compare paired records.
        for source_item, target_item in pairs:
            _, source_excel_row, source_row = source_item
            _, target_excel_row, target_row = target_item

            changes = changed_fields(
                source_row,
                target_row,
                compare_columns,
                mapping,
                ignore_spaces,
                ignore_case,
                numeric_tolerance,
                normalize_dates,
                {},
                value_aliases,
            )

            if not changes:
                results.append(
                    {
                        "Status": "Matched",
                        "Match Key": key_text,
                        "Source Row": source_excel_row,
                        "Target Row": target_excel_row,
                        "Column": "",
                        "Target Column": "",
                        "Source Value": "",
                        "Target Value": "",
                    }
                )
            else:
                for change in changes:
                    results.append(
                        {
                            "Status": "Value Difference",
                            "Match Key": key_text,
                            "Source Row": source_excel_row,
                            "Target Row": target_excel_row,
                            **change,
                        }
                    )

        # Any remaining source rows are missing in target.
        for _, source_excel_row, source_row in unmatched_source:
            results.append(
                {
                    "Status": "Missing in Target",
                    "Match Key": key_text,
                    "Source Row": source_excel_row,
                    "Target Row": "",
                    "Column": "",
                    "Target Column": "",
                    "Source Value": "",
                    "Target Value": "",
                }
            )

        # Any remaining target rows are extra in target.
        for _, target_excel_row, target_row in unmatched_target:
            results.append(
                {
                    "Status": "Extra in Target",
                    "Match Key": key_text,
                    "Source Row": "",
                    "Target Row": target_excel_row,
                    "Column": "",
                    "Target Column": "",
                    "Source Value": "",
                    "Target Value": "",
                }
            )

    result_df = pd.DataFrame(results)

    if result_df.empty:
        result_df = pd.DataFrame(
            columns=[
                "Status", "Match Key", "Source Row", "Target Row",
                "Column", "Target Column", "Source Value", "Target Value"
            ]
        )

    order = {
        "Value Difference": 0,
        "Missing in Target": 1,
        "Extra in Target": 2,
        "Matched": 3,
    }
    result_df["_order"] = result_df["Status"].map(order).fillna(99)
    result_df = result_df.sort_values(
        by=["_order", "Match Key", "Source Row", "Target Row"],
        na_position="last"
    ).drop(columns="_order")

    summary = {
        "Source Rows": len(source_df),
        "Target Rows": len(target_df),
        "Matched Records": int((result_df["Status"] == "Matched").sum()),
        "Missing in Target": int((result_df["Status"] == "Missing in Target").sum()),
        "Extra in Target": int((result_df["Status"] == "Extra in Target").sum()),
        "Value Differences": int((result_df["Status"] == "Value Difference").sum()),
    }

    summary["Result"] = "PASS" if (
        summary["Missing in Target"] == 0
        and summary["Extra in Target"] == 0
        and summary["Value Differences"] == 0
    ) else "FAIL"

    return result_df, summary


# ============================================================
# PDF comparison
# ============================================================

def normalize_pdf_text(text, ignore_spaces=False, ignore_case=False):
    text = text or ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    if ignore_spaces:
        text = re.sub(r"\s+", "", text)

    if ignore_case:
        text = text.lower()

    return text.strip()


def extract_pdf(uploaded_file):
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed. Run: pip install pypdf")

    uploaded_file.seek(0)
    reader = PdfReader(uploaded_file)
    pages = []

    for page in reader.pages:
        pages.append(page.extract_text() or "")

    metadata = {}
    if reader.metadata:
        for key, value in reader.metadata.items():
            metadata[str(key)] = "" if value is None else str(value)

    return pages, metadata


def extract_pdf_with_ocr(uploaded_file):
    """
    Extract normal PDF text with pypdf first and OCR only pages that have no
    extractable text.

    pypdfium2 is preferred for rendering because it does not require Poppler.
    pdf2image remains a fallback for environments that already have Poppler.
    """
    if PdfReader is None:
        raise RuntimeError("pypdf is not installed. Run: pip install pypdf")

    if pytesseract is None:
        raise RuntimeError(
            "Tesseract OCR is not installed. Install pytesseract and the "
            "Tesseract OCR application on Windows."
        )

    uploaded_file.seek(0)
    raw = uploaded_file.read()
    reader = PdfReader(io.BytesIO(raw))

    # First use the PDF's text layer. This means OCR is not required for
    # ordinary text PDFs even when the OCR checkbox is enabled.
    pages = []
    needs_ocr = []

    for index, page in enumerate(reader.pages):
        extracted = page.extract_text() or ""
        pages.append(extracted)
        if not extracted.strip():
            needs_ocr.append(index)

    if needs_ocr:
        if pdfium is not None:
            pdf = pdfium.PdfDocument(raw)
            try:
                for index in needs_ocr:
                    page = pdf[index]
                    bitmap = page.render(scale=2.0)
                    image = bitmap.to_pil()
                    pages[index] = pytesseract.image_to_string(image)
                    page.close()
            finally:
                pdf.close()

        elif convert_from_bytes is not None:
            try:
                images = convert_from_bytes(
                    raw,
                    dpi=180,
                    first_page=min(needs_ocr) + 1,
                    last_page=max(needs_ocr) + 1,
                )
            except Exception as exc:
                raise RuntimeError(
                    "OCR rendering failed. Install pypdfium2 "
                    "(recommended) or install Poppler and add it to PATH."
                ) from exc

            # pdf2image returns a continuous range. Map only the pages that
            # need OCR within that range.
            first = min(needs_ocr)
            for offset, image in enumerate(images):
                page_index = first + offset
                if page_index in needs_ocr:
                    pages[page_index] = pytesseract.image_to_string(image)

        else:
            raise RuntimeError(
                "OCR rendering is unavailable. Install pypdfium2 "
                "(recommended, no Poppler required) or install Poppler."
            )

    meta = {}
    if reader.metadata:
        meta = {
            str(k): "" if v is None else str(v)
            for k, v in reader.metadata.items()
        }

    return pages, meta



def compare_pdfs(source_file, target_file, ignore_spaces=False,
                  ignore_case=False, compare_metadata=False, use_ocr=False):
    if use_ocr:
        source_pages, source_meta = extract_pdf_with_ocr(source_file)
        target_pages, target_meta = extract_pdf_with_ocr(target_file)
    else:
        source_pages, source_meta = extract_pdf(source_file)
        target_pages, target_meta = extract_pdf(target_file)

    rows = []

    max_pages = max(len(source_pages), len(target_pages))

    for page_number in range(max_pages):
        source_text = source_pages[page_number] if page_number < len(source_pages) else ""
        target_text = target_pages[page_number] if page_number < len(target_pages) else ""

        source_norm = normalize_pdf_text(
            source_text, ignore_spaces, ignore_case
        )
        target_norm = normalize_pdf_text(
            target_text, ignore_spaces, ignore_case
        )

        if page_number >= len(source_pages):
            status = "Extra in Target"
        elif page_number >= len(target_pages):
            status = "Missing in Target"
        elif source_norm == target_norm:
            status = "Matched"
        else:
            status = "Value Difference"

        rows.append(
            {
                "Status": status,
                "Page": page_number + 1,
                "Source Text": source_text[:5000],
                "Target Text": target_text[:5000],
            }
        )

    if compare_metadata:
        all_meta_keys = list(dict.fromkeys(
            list(source_meta.keys()) + list(target_meta.keys())
        ))

        for key in all_meta_keys:
            sv = source_meta.get(key, "")
            tv = target_meta.get(key, "")

            if normalize_pdf_text(str(sv), ignore_spaces, ignore_case) != normalize_pdf_text(
                str(tv), ignore_spaces, ignore_case
            ):
                rows.append(
                    {
                        "Status": "Metadata Difference",
                        "Page": "",
                        "Source Text": f"{key}: {sv}",
                        "Target Text": f"{key}: {tv}",
                    }
                )

    result_df = pd.DataFrame(rows)

    summary = {
        "Source Pages": len(source_pages),
        "Target Pages": len(target_pages),
        "Matched Pages": int((result_df["Status"] == "Matched").sum()),
        "Missing Pages": int((result_df["Status"] == "Missing in Target").sum()),
        "Extra Pages": int((result_df["Status"] == "Extra in Target").sum()),
        "Value Differences": int(
            (result_df["Status"].isin(["Value Difference", "Metadata Difference"])).sum()
        ),
    }

    summary["Result"] = "PASS" if (
        summary["Missing Pages"] == 0
        and summary["Extra Pages"] == 0
        and summary["Value Differences"] == 0
    ) else "FAIL"

    return result_df, summary


# ============================================================
# Setup/profile functions
# ============================================================

def setup_file_path(name):
    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", name)
    return SETUP_DIR / f"{safe_name}.json"


def save_setup(name, mapping, key_columns, selected_columns,
               ignore_spaces, ignore_case, composite_mode):
    data = {
        "name": name,
        "mapping": mapping,
        "key_columns": key_columns,
        "selected_columns": selected_columns,
        "ignore_spaces": ignore_spaces,
        "ignore_case": ignore_case,
        "composite_mode": composite_mode,
        "include_match_keys": st.session_state.get("ui_include_match_keys", False),
        "column_rules": st.session_state.get("column_rules", {}),
        "numeric_tolerance": st.session_state.get("numeric_tolerance", 0.0),
        "normalize_dates": st.session_state.get("normalize_dates", False),
        "value_aliases": st.session_state.get("value_aliases", {}),
        "default_timezone": st.session_state.get("default_timezone", "Asia/Kolkata"),
    }

    with open(setup_file_path(name), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def load_setup(name):
    path = setup_file_path(name)

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def delete_setup(name):
    path = setup_file_path(name)
    if path.exists():
        path.unlink()


def available_setups():
    return sorted(
        p.stem for p in SETUP_DIR.glob("*.json")
    )


# ============================================================
# Session state helpers
# ============================================================

def init_state():
    defaults = {
        "mapping": {},
        "match_keys": [],
        "selected_columns": [],
        "composite_mode": False,
        "ignore_spaces_value": False,
        "ignore_case_value": False,
        "pending_setup": None,
        "setup_message": "",
        "swap_files": False,
        "column_rules": {},
        "value_aliases": {},
        "default_timezone": "Asia/Kolkata",
        "numeric_tolerance": 0.0,
        "normalize_dates": False,
        "last_comparison_time": "",
        "comparison_history": [],
        # UI widget state must exist before the multiselect is rendered.
        "ui_composite_match_keys": [],
        "ui_composite_mode": False,
        "ui_single_match_key": "",
        "ui_selected_columns": [],
        "ui_ignore_spaces": False,
        "ui_ignore_case": False,
        "ui_numeric_tolerance": 0.0,
        "ui_normalize_dates": False,
        "ui_include_match_keys": False,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def apply_pending_setup(source_columns, target_columns):
    """
    Apply a saved profile before mapping widgets are instantiated.

    The important detail is that Streamlit widgets retain their own state.
    Therefore, loading a profile must explicitly initialize the widget keys
    before the selectboxes/multiselects are created.
    """
    profile = st.session_state.get("pending_setup")

    if not profile:
        return False

    saved_mapping = profile.get("mapping", {}) or {}

    # Keep only mappings that are valid for the currently uploaded files.
    valid_mapping = {
        source_col: target_col
        for source_col, target_col in saved_mapping.items()
        if source_col in source_columns and target_col in target_columns
    }

    saved_keys = [
        col
        for col in (profile.get("key_columns", []) or [])
        if col in source_columns and col in valid_mapping
    ]

    saved_selected = [
        col
        for col in (profile.get("selected_columns", []) or [])
        if col in source_columns and col in valid_mapping
    ]

    # If the profile did not contain selected columns, compare all mapped
    # columns by default.
    if not saved_selected:
        saved_selected = list(valid_mapping.keys())

    composite_mode = bool(profile.get("composite_mode", False))
    ignore_spaces = bool(profile.get("ignore_spaces", False))
    ignore_case = bool(profile.get("ignore_case", False))

    # Application state.
    st.session_state.mapping = valid_mapping
    st.session_state.match_keys = saved_keys
    st.session_state.selected_columns = saved_selected
    st.session_state.composite_mode = composite_mode
    st.session_state.ui_include_match_keys = bool(
        profile.get("include_match_keys", False)
    )
    st.session_state.ignore_spaces_value = ignore_spaces
    st.session_state.ignore_case_value = ignore_case
    st.session_state.column_rules = profile.get("column_rules", {}) or {}
    st.session_state.numeric_tolerance = float(profile.get("numeric_tolerance", 0) or 0)
    st.session_state.normalize_dates = bool(profile.get("normalize_dates", False))
    st.session_state.value_aliases = profile.get("value_aliases", {}) or {}
    st.session_state.default_timezone = profile.get(
        "default_timezone", "Asia/Kolkata"
    )

    # Clear mapping widget state for columns that are no longer mapped.
    # This prevents an old selectbox value from overriding the loaded profile.
    for source_col in source_columns:
        widget_key = f"ui_target_map_{source_col}"

        if source_col in valid_mapping:
            st.session_state[widget_key] = valid_mapping[source_col]
        else:
            st.session_state.pop(widget_key, None)

    # Initialize all other widget states BEFORE their widgets are created.
    st.session_state["ui_composite_mode"] = composite_mode

    single_key = (
        saved_keys[0]
        if saved_keys
        else (list(valid_mapping.keys())[0] if valid_mapping else "")
    )

    st.session_state["ui_single_match_key"] = single_key
    st.session_state["ui_composite_match_keys"] = saved_keys.copy()
    st.session_state["ui_selected_columns"] = saved_selected.copy()
    st.session_state["ui_ignore_spaces"] = ignore_spaces
    st.session_state["ui_ignore_case"] = ignore_case

    # Reset the setup selector so it doesn't keep an unrelated stale choice.
    st.session_state.pop("ui_setup_choice", None)

    setup_name = profile.get("name", "Unnamed")
    invalid_mapping_count = len(saved_mapping) - len(valid_mapping)

    if invalid_mapping_count:
        st.session_state.setup_message = (
            f"Loaded setup '{setup_name}'. "
            f"{invalid_mapping_count} saved mapping(s) could not be used "
            "because the current files do not contain the required column."
        )
    else:
        st.session_state.setup_message = (
            f"Loaded setup '{setup_name}'. "
            "The saved column mapping has been populated."
        )

    # Consume the pending profile.
    st.session_state.pending_setup = None

    return True


# ============================================================
# Excel/CSV UI
# ============================================================

def excel_comparison_ui(source_file, target_file):
    source_type = get_file_type(source_file)
    target_type = get_file_type(target_file)

    if source_type == "excel":
        source_sheets = get_excel_sheets(source_file)
        source_sheet = st.selectbox(
            "Source Sheet",
            source_sheets,
            key="ui_source_sheet",
        )
    else:
        source_sheet = None

    if target_type == "excel":
        target_sheets = get_excel_sheets(target_file)
        target_sheet = st.selectbox(
            "Target Sheet",
            target_sheets,
            key="ui_target_sheet",
        )
    else:
        target_sheet = None

    try:
        source_df = load_dataframe(source_file, source_sheet)
        target_df = load_dataframe(target_file, target_sheet)
    except Exception as e:
        st.error(f"Unable to read files: {e}")
        return

    source_df.columns = [str(c) for c in source_df.columns]
    target_df.columns = [str(c) for c in target_df.columns]

    source_columns = list(source_df.columns)
    target_columns = list(target_df.columns)

    a, b = st.columns(2)
    with a:
        st.markdown(f"**Source:** `{source_file.name}`")
        st.caption(f"{len(source_df):,} rows × {len(source_df.columns):,} columns")
    with b:
        st.markdown(f"**Target:** `{target_file.name}`")
        st.caption(f"{len(target_df):,} rows × {len(target_df.columns):,} columns")

    with st.expander("🔍 Auto Mapping Suggestions"):
        suggestions = suggest_column_mappings(source_columns, target_columns)
        if suggestions:
            st.dataframe(
                pd.DataFrame([
                    {"Source": s, "Suggested Target": t, "Confidence": f"{p}%"}
                    for s, (t, p) in suggestions.items()
                ]),
                hide_index=True,
                use_container_width=True,
            )
            if st.button("Apply Suggestions ≥ 75%", key="apply_auto_mapping"):
                for s, (t, p) in suggestions.items():
                    if p >= 75:
                        st.session_state.mapping[s] = t
                        st.session_state[f"ui_target_map_{s}"] = t
                st.rerun()
        else:
            st.info("No mapping suggestions available.")

    # Apply profile before mapping widgets are instantiated.
    apply_pending_setup(source_columns, target_columns)

    if st.session_state.setup_message:
        st.info(st.session_state.setup_message)
        st.session_state.setup_message = ""

    st.subheader("Saved Setup")

    setups = available_setups()

    if setups:
        setup_choice = st.selectbox(
            "Select saved setup",
            ["-- Select --"] + setups,
            key="ui_setup_choice",
        )

        col_a, col_b = st.columns(2)

        with col_a:
            if st.button(
                "Load Setup",
                key="load_setup_button",
                disabled=setup_choice == "-- Select --",
            ):
                try:
                    profile = load_setup(setup_choice)

                    # Do not directly modify mapping widgets here. Store the
                    # profile and rerun. It will be applied at the top of this
                    # function BEFORE any mapping widget is instantiated.
                    st.session_state.pending_setup = profile
                    st.rerun()

                except Exception as e:
                    st.error(f"Unable to load setup: {e}")

        with col_b:
            if st.button(
                "Delete Setup",
                key="delete_setup_button",
                disabled=setup_choice == "-- Select --",
            ):
                delete_setup(setup_choice)
                st.success(f"Deleted setup '{setup_choice}'.")
                st.rerun()
    else:
        st.caption("No saved setups yet.")

    st.divider()
    st.subheader("Column Mapping")

    # --------------------------------------------------------
    # Column Mapping
    # --------------------------------------------------------
    # Only required columns need to be mapped. Use the X button
    # to remove a mapping row.
    current_mapping = {
        s: t for s, t in st.session_state.get("mapping", {}).items()
        if s in source_columns and t in target_columns
    }
    st.session_state.mapping = current_mapping

    unmapped = [c for c in source_columns if c not in current_mapping]

    add1, add2 = st.columns([5, 1])
    with add1:
        selected_to_add = st.selectbox(
            "Add Source Column to Mapping",
            unmapped if unmapped else ["All source columns are mapped"],
            key="ui_add_mapping_column",
            disabled=not bool(unmapped),
        )
    with add2:
        st.write("")
        st.write("")
        if st.button("＋ Add", key="add_mapping_column_button",
                     disabled=not bool(unmapped)):
            target_default = (
                selected_to_add
                if selected_to_add in target_columns
                else (target_columns[0] if target_columns else "")
            )
            st.session_state.mapping[selected_to_add] = target_default
            st.session_state[f"ui_target_map_{selected_to_add}"] = target_default
            st.rerun()

    if not current_mapping:
        st.info("Add the Source columns you want to compare.")
    else:
        for source_col in list(current_mapping.keys()):
            c1, c2, c3 = st.columns([3, 5, 0.7])
            with c1:
                st.write(f"**{source_col}**")
            widget_key = f"ui_target_map_{source_col}"
            if widget_key not in st.session_state:
                st.session_state[widget_key] = current_mapping[source_col]
            with c2:
                target_value = st.selectbox(
                    "Target Column", target_columns, key=widget_key,
                    label_visibility="collapsed",
                )
            with c3:
                if st.button("✕", key=f"remove_mapping_{source_col}",
                             help=f"Remove {source_col} from mapping"):
                    st.session_state.mapping.pop(source_col, None)
                    st.session_state.match_keys = [
                        c for c in st.session_state.match_keys if c != source_col
                    ]
                    st.session_state.selected_columns = [
                        c for c in st.session_state.selected_columns if c != source_col
                    ]
                    st.session_state.pop(widget_key, None)
                    st.rerun()
            current_mapping[source_col] = target_value

    st.session_state.mapping = current_mapping

    mapped_columns = list(st.session_state.mapping.keys())

    # Column-specific rule controls were removed from this section to avoid
    # duplicating the same comparison options shown under "Columns to Compare".
    # Comparison behavior is now controlled from the single global rules area
    # below Columns to Compare.
    st.subheader("Match Key")

    if "ui_composite_mode" not in st.session_state:
        st.session_state.ui_composite_mode = st.session_state.composite_mode

    composite_mode = st.checkbox(
        "Use Composite Match Key",
        key="ui_composite_mode",
    )

    st.session_state.composite_mode = composite_mode

    mapped_source_columns = list(st.session_state.mapping.keys())
    if not mapped_source_columns:
        st.warning("Add at least one column to the mapping before selecting a Match Key.")
        return

    if composite_mode:
        default_keys = [
            c for c in st.session_state.match_keys
            if c in source_columns
        ]

        st.session_state.ui_composite_match_keys = [
            c for c in st.session_state.get("ui_composite_match_keys", [])
            if c in mapped_source_columns
        ]

        selected_keys = st.multiselect(
            "Select Source columns for Match Key",
            mapped_source_columns,
            key="ui_composite_match_keys",
        )

        if not selected_keys:
            st.warning("Select at least one Match Key column.")
            return

        match_keys = selected_keys
    else:
        default_key = (
            st.session_state.match_keys[0]
            if st.session_state.match_keys
            and st.session_state.match_keys[0] in source_columns
            else mapped_source_columns[0]
        )

        if (
            "ui_single_match_key" not in st.session_state
            or st.session_state.ui_single_match_key not in mapped_source_columns
        ):
            st.session_state.ui_single_match_key = default_key

        match_key = st.selectbox(
            "Select Source column for Match Key",
            mapped_source_columns,
            key="ui_single_match_key",
        )

        match_keys = [match_key]

    st.session_state.match_keys = match_keys

    st.subheader("Columns to Compare")

    # All mapped columns are selected automatically by default.
    # Match Key columns can optionally be included in value comparison.
    if "ui_include_match_keys" not in st.session_state:
        st.session_state.ui_include_match_keys = False

    include_match_keys = st.checkbox(
        "Include Match Key columns in comparison",
        key="ui_include_match_keys",
        help=(
            "When enabled, Match Key columns are also checked for value differences. "
            "Otherwise, Match Keys are used only to identify/group records."
        ),
    )

    default_compare = [
        c for c in mapped_source_columns
        if include_match_keys or c not in match_keys
    ]

    # Keep existing user selections, while automatically adding newly mapped
    # columns and removing columns that are no longer mapped.
    existing_selected = [
        c for c in st.session_state.get("ui_selected_columns", [])
        if c in mapped_source_columns
    ]

    # Newly mapped columns are automatically selected.
    selected_columns = list(dict.fromkeys(existing_selected + default_compare))

    # Match Key columns are excluded automatically unless the user explicitly
    # enables "Include Match Key columns in comparison".
    if not include_match_keys:
        selected_columns = [c for c in selected_columns if c not in match_keys]

    st.session_state.ui_selected_columns = selected_columns
    st.session_state.selected_columns = selected_columns

    if selected_columns:
        st.multiselect(
            "Selected columns",
            mapped_source_columns,
            key="ui_selected_columns",
            help="Mapped columns are selected automatically. Remove any column if you do not want to compare it.",
        )
    else:
        st.info(
            "No non-Match-Key columns are currently selected. "
            "Enable 'Include Match Key columns in comparison' if you also want to validate the Match Key values."
        )

    # Refresh the application state after the widget interaction.
    st.session_state.selected_columns = st.session_state.get(
        "ui_selected_columns", []
    )

    col1, col2 = st.columns(2)

    with col1:
        if "ui_ignore_spaces" not in st.session_state:
            st.session_state.ui_ignore_spaces = st.session_state.ignore_spaces_value

        ignore_spaces = st.checkbox(
            "Ignore spaces",
            key="ui_ignore_spaces",
        )

    with col2:
        if "ui_ignore_case" not in st.session_state:
            st.session_state.ui_ignore_case = st.session_state.ignore_case_value

        ignore_case = st.checkbox(
            "Ignore case",
            key="ui_ignore_case",
        )

    st.session_state.ignore_spaces_value = ignore_spaces
    st.session_state.ignore_case_value = ignore_case

    q1, q2 = st.columns(2)
    with q1:
        numeric_tolerance = st.number_input(
            "Global numeric tolerance",
            min_value=0.0,
            value=float(st.session_state.numeric_tolerance),
            step=0.01,
            key="ui_numeric_tolerance",
        )
    with q2:
        normalize_dates = st.checkbox(
            "Normalize date formats",
            value=bool(st.session_state.normalize_dates),
            key="ui_normalize_dates",
        )
    st.session_state.numeric_tolerance = numeric_tolerance
    st.session_state.normalize_dates = normalize_dates

    st.markdown("### Date / Timezone Handling")
    st.caption(
        "Timezone-aware timestamps are converted to UTC. Values without a timezone "
        "use the selected default timezone."
    )
    timezone_options = {
        "India Standard Time (IST) — Asia/Kolkata": "Asia/Kolkata",
        "UTC": "UTC",
        "Eastern Time — America/New_York": "America/New_York",
        "Central Time — America/Chicago": "America/Chicago",
        "Mountain Time — America/Denver": "America/Denver",
        "Pacific Time — America/Los_Angeles": "America/Los_Angeles",
        "UK Time — Europe/London": "Europe/London",
        "Central Europe — Europe/Berlin": "Europe/Berlin",
        "Singapore — Asia/Singapore": "Asia/Singapore",
    }
    current_tz = st.session_state.get("default_timezone", "Asia/Kolkata")
    current_label = next(
        (label for label, tz in timezone_options.items() if tz == current_tz),
        "India Standard Time (IST) — Asia/Kolkata",
    )
    selected_tz_label = st.selectbox(
        "Default timezone for values without timezone",
        list(timezone_options.keys()),
        index=list(timezone_options.keys()).index(current_label),
        key="ui_default_timezone",
    )
    st.session_state.default_timezone = timezone_options[selected_tz_label]

    with st.expander("🔄 Value Matching / Aliases"):
        st.caption(
            "Define intentional Source → Target value changes. "
            "Example: VisionLink → NEW VISIONLINK."
        )

        alias_col = st.selectbox(
            "Column",
            mapped_source_columns,
            key="ui_alias_column",
        )

        alias_source = st.text_input(
            "Old System Value",
            key="ui_alias_source",
            placeholder="Example: VisionLink",
        )

        alias_target = st.text_input(
            "New System Value",
            key="ui_alias_target",
            placeholder="Example: NEW VISIONLINK",
        )

        if st.button("＋ Add Value Rule", key="add_value_alias"):
            source_value = alias_source.strip()
            target_value = alias_target.strip()

            if not source_value or not target_value:
                st.warning("Enter both Old System Value and New System Value.")
            else:
                aliases = st.session_state.value_aliases.setdefault(alias_col, [])
                duplicate = any(
                    item.get("source") == source_value
                    and item.get("target") == target_value
                    for item in aliases
                )

                if duplicate:
                    st.warning("This value rule already exists.")
                else:
                    aliases.append({
                        "source": source_value,
                        "target": target_value,
                    })
                    # Do not modify ui_alias_source/ui_alias_target here.
                    # They are widget-owned session-state keys and cannot be
                    # changed after the text_input widgets are instantiated.
                    st.rerun()

        column_aliases = st.session_state.value_aliases.get(alias_col, [])

        if column_aliases:
            st.markdown("**Current Rules**")
            for alias_index, alias in enumerate(column_aliases):
                a1, a2, a3 = st.columns([4, 4, 1])
                with a1:
                    st.write(str(alias.get("source", "")))
                with a2:
                    st.write(str(alias.get("target", "")))
                with a3:
                    if st.button(
                        "✕",
                        key=f"remove_value_alias_{alias_col}_{alias_index}",
                        help="Remove this value rule",
                    ):
                        column_aliases.pop(alias_index)
                        st.rerun()
        else:
            st.info(f"No value rules configured for '{alias_col}'.")

    with st.expander("🧠 Fuzzy Match Suggestions"):
        st.caption("Advisory only; exact matching remains the comparison rule.")
        st.slider("Similarity threshold", 70, 100, 90, key="ui_fuzzy_threshold")

    # --------------------------------------------------------
    # Save / Update setup
    # --------------------------------------------------------
    st.subheader("Save / Update Setup")

    setup_name = st.text_input(
        "Setup name",
        key="ui_setup_name",
        placeholder="Example: Employee Comparison",
    )

    save_col, update_col = st.columns(2)

    with save_col:
        if st.button("Save New Setup", key="save_setup_button"):
            clean_name = setup_name.strip()

            if not clean_name:
                st.warning("Enter a setup name.")
            elif not selected_columns:
                st.warning("Select at least one column to compare.")
            elif not match_keys:
                st.warning("Select at least one Match Key.")
            elif setup_file_path(clean_name).exists():
                st.warning(
                    f"Setup '{clean_name}' already exists. "
                    "Use Update Selected Setup to modify an existing setup."
                )
            else:
                save_setup(
                    clean_name,
                    st.session_state.mapping,
                    match_keys,
                    selected_columns,
                    ignore_spaces,
                    ignore_case,
                    composite_mode,
                )
                st.success(f"Setup '{clean_name}' saved.")
                st.rerun()

    with update_col:
        selected_setup_for_update = st.selectbox(
            "Setup to update",
            ["-- Select --"] + available_setups(),
            key="ui_update_setup_choice",
        )

        if st.button(
            "Update Selected Setup",
            key="update_setup_button",
            disabled=selected_setup_for_update == "-- Select --",
        ):
            update_name = selected_setup_for_update

            if not selected_columns:
                st.warning("Select at least one column to compare.")
            elif not match_keys:
                st.warning("Select at least one Match Key.")
            else:
                save_setup(
                    update_name,
                    st.session_state.mapping,
                    match_keys,
                    selected_columns,
                    ignore_spaces,
                    ignore_case,
                    composite_mode,
                )
                st.success(f"Setup '{update_name}' updated successfully.")
                st.rerun()

    # --------------------------------------------------------
    # Compare
    # --------------------------------------------------------
    st.divider()

    if st.button("Compare Files", type="primary", key="compare_excel_button"):
        if not selected_columns:
            st.error("Select at least one column to compare.")
            return

        missing_mappings = [
            c for c in selected_columns
            if c not in st.session_state.mapping
            or st.session_state.mapping[c] not in target_columns
        ]

        if missing_mappings:
            st.error(
                "Missing target mapping for: "
                + ", ".join(missing_mappings)
            )
            return

        key_mapping_missing = [
            c for c in match_keys
            if c not in st.session_state.mapping
            or st.session_state.mapping[c] not in target_columns
        ]

        if key_mapping_missing:
            st.error(
                "Missing target mapping for Match Key: "
                + ", ".join(key_mapping_missing)
            )
            return

        with st.spinner("Comparing files..."):
            result_df, summary = compare_dataframes(
                source_df,
                target_df,
                st.session_state.mapping,
                match_keys,
                selected_columns,
                ignore_spaces,
                ignore_case,
                numeric_tolerance,
                normalize_dates,
                {},
                st.session_state.get("value_aliases", {}),
                st.session_state.get("default_timezone", "Asia/Kolkata"),
            )

        st.session_state.last_excel_result = result_df
        st.session_state.last_excel_summary = summary
        st.session_state.last_source_name = source_file.name
        st.session_state.last_target_name = target_file.name
        st.session_state.last_column_summary = difference_by_column(result_df)
        st.session_state.last_duplicate_summary = duplicate_key_report(
            source_df,
            target_df,
            match_keys,
            st.session_state.mapping,
            ignore_spaces,
            ignore_case,
        )
        st.session_state.last_comparison_time = pd.Timestamp.now().strftime("%d-%b-%Y %H:%M:%S")
        st.session_state.comparison_history.append({
            "Time": st.session_state.last_comparison_time,
            "Source": source_file.name,
            "Target": target_file.name,
            "Result": summary["Result"],
            "Matched": summary["Matched Records"],
            "Changed": summary["Value Differences"],
            "Missing": summary["Missing in Target"],
            "Extra": summary["Extra in Target"],
        })
        st.session_state.comparison_history = st.session_state.comparison_history[-10:]

    if (
        "last_excel_result" in st.session_state
        and st.session_state.last_excel_result is not None
        and "last_excel_summary" in st.session_state
        and st.session_state.last_excel_summary is not None
    ):
        show_excel_results(
            st.session_state.last_excel_result,
            st.session_state.last_excel_summary,
        )


def show_excel_results(result_df, summary):
    st.divider()
    st.subheader("Comparison Summary")

    cols = st.columns(7)

    metrics = [
        ("Source Rows", summary["Source Rows"]),
        ("Target Rows", summary["Target Rows"]),
        ("Matched", summary["Matched Records"]),
        ("Missing", summary["Missing in Target"]),
        ("Extra", summary["Extra in Target"]),
        ("Changed", summary["Value Differences"]),
        ("Result", summary["Result"]),
    ]

    for col, (label, value) in zip(cols, metrics):
        col.metric(label, value)

    # --------------------------------------------------------
    # Split result into clear sections.
    # Matched records are NOT mixed with differences.
    # --------------------------------------------------------
    filtered_result = result_df.copy()

    value_diff_df = filtered_result[
        filtered_result["Status"] == "Value Difference"
    ].copy()

    missing_df = filtered_result[
        filtered_result["Status"] == "Missing in Target"
    ].copy()

    extra_df = filtered_result[
        filtered_result["Status"] == "Extra in Target"
    ].copy()

    matched_df = filtered_result[
        filtered_result["Status"] == "Matched"
    ].copy()

    # --------------------------------------------------------
    # Value Differences
    # --------------------------------------------------------
    st.subheader(
        f"🔴 Value Differences ({len(value_diff_df)})"
    )

    if value_diff_df.empty:
        st.success("No value differences found.")
    else:
        value_display = value_diff_df[
            [
                "Match Key",
                "Source Row",
                "Target Row",
                "Column",
                "Target Column",
                "Source Value",
                "Target Value",
            ]
        ].copy()

        st.dataframe(
            value_display,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # Missing in Target
    # --------------------------------------------------------
    st.subheader(
        f"🟠 Missing in Target ({len(missing_df)})"
    )

    if missing_df.empty:
        st.success("No records are missing in Target.")
    else:
        missing_display = missing_df[
            ["Match Key", "Source Row", "Target Row"]
        ].copy()

        missing_display["Details"] = "Entire record missing in Target"

        st.dataframe(
            missing_display,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # Extra in Target
    # --------------------------------------------------------
    st.subheader(
        f"🟡 Extra in Target ({len(extra_df)})"
    )

    if extra_df.empty:
        st.success("No extra records found in Target.")
    else:
        extra_display = extra_df[
            ["Match Key", "Source Row", "Target Row"]
        ].copy()

        extra_display["Details"] = "Entire record extra in Target"

        st.dataframe(
            extra_display,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # Matched Records
    # --------------------------------------------------------
    with st.expander(
        f"🟢 Matched Records ({len(matched_df)})",
        expanded=False,
    ):
        if matched_df.empty:
            st.info("No completely matched records.")
        else:
            matched_display = matched_df[
                ["Match Key", "Source Row", "Target Row"]
            ].copy()

            st.dataframe(
                matched_display,
                use_container_width=True,
                hide_index=True,
            )

    # --------------------------------------------------------
    # --------------------------------------------------------
    # Professional QA Excel report
    # --------------------------------------------------------
    output = io.BytesIO()

    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    def _style_report_sheet(ws, title=None, freeze="A2", autofilter=True):
        ws.sheet_view.showGridLines = False

        if title:
            max_col = max(ws.max_column, 1)
            ws.insert_rows(1, 2)
            ws.merge_cells(
                start_row=1, start_column=1,
                end_row=1, end_column=max_col
            )
            title_cell = ws.cell(1, 1, title)
            title_cell.font = Font(size=16, bold=True)
            title_cell.alignment = Alignment(vertical="center")
            ws.row_dimensions[1].height = 28
            ws.row_dimensions[2].height = 8

        header_row = 3 if title else 1

        header_fill = PatternFill("solid", fgColor="1F2937")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="D1D5DB")

        for cell in ws[header_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )
            cell.border = Border(bottom=thin)

        ws.freeze_panes = freeze if not title else f"A{header_row + 1}"

        if autofilter and ws.max_row >= header_row:
            ws.auto_filter.ref = (
                f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
            )

        for row in ws.iter_rows(min_row=header_row + 1):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        for col_cells in ws.iter_cols(
            min_row=header_row,
            max_row=ws.max_row,
        ):
            if not col_cells:
                continue

            max_length = 0
            for cell in col_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))

            # A merged title row creates MergedCell objects that do not have
            # column_letter. Use the real column index instead.
            column_index = col_cells[0].column
            column_letter = get_column_letter(column_index)

            width = min(max(max_length + 3, 12), 50)
            ws.column_dimensions[column_letter].width = width

    def _style_status_cell(cell, result):
        if result == "PASS":
            cell.fill = PatternFill("solid", fgColor="C6EFCE")
            cell.font = Font(color="006100", bold=True)
        else:
            cell.fill = PatternFill("solid", fgColor="FFC7CE")
            cell.font = Font(color="9C0006", bold=True)
        cell.alignment = Alignment(horizontal="center")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # ---------------------------
        # Summary
        # ---------------------------
        summary_rows = [
            {"Metric": "Comparison Result", "Value": summary["Result"]},
            {"Metric": "Source Rows", "Value": summary["Source Rows"]},
            {"Metric": "Target Rows", "Value": summary["Target Rows"]},
            {"Metric": "Matched Records", "Value": summary["Matched Records"]},
            {"Metric": "Missing in Target", "Value": summary["Missing in Target"]},
            {"Metric": "Extra in Target", "Value": summary["Extra in Target"]},
            {"Metric": "Value Differences", "Value": summary["Value Differences"]},
            {
                "Metric": "Match Rate",
                "Value": (
                    f"{summary['Matched Records'] / summary['Source Rows'] * 100:.2f}%"
                    if summary["Source Rows"] else "0.00%"
                ),
            },
            {
                "Metric": "Comparison Time",
                "Value": st.session_state.get("last_comparison_time", ""),
            },
            {
                "Metric": "Source File",
                "Value": st.session_state.get("last_source_name", ""),
            },
            {
                "Metric": "Target File",
                "Value": st.session_state.get("last_target_name", ""),
            },
        ]
        summary_df = pd.DataFrame(summary_rows)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        ws = writer.book["Summary"]
        _style_report_sheet(ws, "Document Comparator — QA Comparison Report")
        result_cell = ws["B3"]
        _style_status_cell(result_cell, summary["Result"])

        # ---------------------------
        # Configuration
        # ---------------------------
        mapping = st.session_state.get("mapping", {})
        match_keys_cfg = st.session_state.get("match_keys", [])
        selected_cfg = st.session_state.get("selected_columns", [])
        aliases_cfg = st.session_state.get("value_aliases", {})
        default_tz = st.session_state.get("default_timezone", "Asia/Kolkata")

        config_rows = [
            {"Setting": "Source File", "Value": st.session_state.get("last_source_name", "")},
            {"Setting": "Target File", "Value": st.session_state.get("last_target_name", "")},
            {"Setting": "Match Key Type", "Value": "Composite" if len(match_keys_cfg) > 1 else "Single"},
            {"Setting": "Match Key Columns", "Value": ", ".join(match_keys_cfg)},
            {"Setting": "Columns Compared", "Value": ", ".join(selected_cfg)},
            {"Setting": "Ignore Spaces", "Value": "ON" if st.session_state.get("ignore_spaces_value", False) else "OFF"},
            {"Setting": "Ignore Case", "Value": "ON" if st.session_state.get("ignore_case_value", False) else "OFF"},
            {"Setting": "Numeric Tolerance", "Value": st.session_state.get("numeric_tolerance_value", 0.0)},
            {"Setting": "Normalize Date/Time", "Value": "ON" if st.session_state.get("normalize_dates_value", False) else "OFF"},
            {"Setting": "Default Timezone", "Value": default_tz},
        ]
        pd.DataFrame(config_rows).to_excel(
            writer, sheet_name="Configuration", index=False
        )

        alias_rows = []
        for source_col, rules in aliases_cfg.items():
            for rule in rules:
                alias_rows.append({
                    "Column": source_col,
                    "Old System Value": rule.get("source", ""),
                    "New System Value": rule.get("target", ""),
                })

        if not alias_rows:
            alias_rows = [{"Column": "", "Old System Value": "", "New System Value": ""}]

        pd.DataFrame(alias_rows).to_excel(
            writer, sheet_name="Value Aliases", index=False
        )

        # ---------------------------
        # Mapping
        # ---------------------------
        mapping_df = pd.DataFrame([
            {
                "Source Column": source,
                "Target Column": target,
                "Match Key": "Yes" if source in match_keys_cfg else "No",
                "Compared": "Yes" if source in selected_cfg else "No",
            }
            for source, target in mapping.items()
        ])
        mapping_df.to_excel(writer, sheet_name="Column Mapping", index=False)

        # ---------------------------
        # Result sheets
        # ---------------------------
        value_display = value_diff_df[
            [
                "Match Key", "Source Row", "Target Row", "Column",
                "Target Column", "Source Value", "Target Value",
            ]
        ].copy()
        value_display.to_excel(
            writer, sheet_name="Value Differences", index=False
        )

        missing_display = missing_df[
            ["Match Key", "Source Row", "Target Row"]
        ].copy()
        missing_display["Details"] = "Entire record missing in Target"
        missing_display.to_excel(
            writer, sheet_name="Missing in Target", index=False
        )

        extra_display = extra_df[
            ["Match Key", "Source Row", "Target Row"]
        ].copy()
        extra_display["Details"] = "Entire record extra in Target"
        extra_display.to_excel(
            writer, sheet_name="Extra in Target", index=False
        )

        matched_display = matched_df[
            ["Match Key", "Source Row", "Target Row"]
        ].copy()
        matched_display.to_excel(
            writer, sheet_name="Matched Records", index=False
        )

        # Differences by column.
        column_summary = st.session_state.get("last_column_summary")
        if isinstance(column_summary, pd.DataFrame) and not column_summary.empty:
            column_summary.to_excel(
                writer, sheet_name="Differences by Column", index=False
            )
        else:
            pd.DataFrame(
                columns=["Column", "Differences"]
            ).to_excel(
                writer, sheet_name="Differences by Column", index=False
            )

        # Duplicate/repeating key information.
        duplicate_summary = st.session_state.get("last_duplicate_summary")
        if isinstance(duplicate_summary, pd.DataFrame) and not duplicate_summary.empty:
            duplicate_summary.to_excel(
                writer, sheet_name="Duplicate Match Keys", index=False
            )
        else:
            pd.DataFrame(
                columns=["Match Key", "Source Count", "Target Count"]
            ).to_excel(
                writer, sheet_name="Duplicate Match Keys", index=False
            )

        # ---------------------------
        # Apply professional formatting
        # ---------------------------
        workbook = writer.book

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            _style_report_sheet(ws, sheet_name)

        # Summary-specific formatting.
        summary_ws = workbook["Summary"]
        for row in range(4, summary_ws.max_row + 1):
            metric = summary_ws.cell(row, 1).value
            if metric in {
                "Matched Records", "Missing in Target",
                "Extra in Target", "Value Differences"
            }:
                summary_ws.cell(row, 2).alignment = Alignment(horizontal="center")

        # Make result prominent.
        summary_ws["B3"].border = Border(
            left=Side(style="medium", color="808080"),
            right=Side(style="medium", color="808080"),
            top=Side(style="medium", color="808080"),
            bottom=Side(style="medium", color="808080"),
        )

        # Highlight non-empty difference sheets.
        if len(value_display) > 0:
            ws = workbook["Value Differences"]
            for cell in ws[3]:
                cell.fill = PatternFill("solid", fgColor="F4CCCC")
                cell.font = Font(color="9C0006", bold=True)

        if len(missing_display) > 0:
            ws = workbook["Missing in Target"]
            for cell in ws[3]:
                cell.fill = PatternFill("solid", fgColor="FCE5CD")
                cell.font = Font(color="9C6500", bold=True)

        if len(extra_display) > 0:
            ws = workbook["Extra in Target"]
            for cell in ws[3]:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
                cell.font = Font(color="7F6000", bold=True)

    if st.session_state.get("last_comparison_time"):
        st.caption(f"Last comparison: {st.session_state.last_comparison_time}")

    st.download_button(
        "Download Professional QA Excel Report",
        data=output.getvalue(),
        file_name="document_comparison_QA_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel_report_v23",
    )


# ============================================================
# PDF UI
# ============================================================

def pdf_comparison_ui(source_file, target_file):
    if PdfReader is None:
        st.error("pypdf is not installed. Run: pip install pypdf")
        return

    col1, col2 = st.columns(2)

    with col1:
        ignore_spaces = st.checkbox(
            "Ignore spaces",
            value=False,
            key="pdf_ignore_spaces",
        )

    with col2:
        ignore_case = st.checkbox(
            "Ignore case",
            value=False,
            key="pdf_ignore_case",
        )

    compare_metadata = st.checkbox(
        "Compare PDF metadata",
        value=False,
        key="pdf_compare_metadata",
    )
    use_ocr = st.checkbox(
        "Use OCR for scanned/image-only PDFs",
        value=False,
        key="pdf_use_ocr",
        help=(
            "Normal text PDFs do not need OCR. For scanned pages, Tesseract is "
            "required. pypdfium2 is used for PDF rendering and does not require Poppler."
        ),
    )

    if st.button("Compare PDFs", type="primary", key="compare_pdf_button"):
        try:
            with st.spinner("Comparing PDFs..."):
                result_df, summary = compare_pdfs(
                    source_file,
                    target_file,
                    ignore_spaces,
                    ignore_case,
                    compare_metadata,
                    use_ocr,
                )

            st.session_state.last_pdf_result = result_df
            st.session_state.last_pdf_summary = summary

        except Exception as e:
            st.error(f"PDF comparison failed: {e}")
            return

    if "last_pdf_result" in st.session_state:
        summary = st.session_state.last_pdf_summary
        result_df = st.session_state.last_pdf_result

        st.divider()
        st.subheader("PDF Comparison Summary")

        cols = st.columns(6)

        metrics = [
            ("Source Pages", summary["Source Pages"]),
            ("Target Pages", summary["Target Pages"]),
            ("Matched", summary["Matched Pages"]),
            ("Missing", summary["Missing Pages"]),
            ("Extra", summary["Extra Pages"]),
            ("Result", summary["Result"]),
        ]

        for col, (label, value) in zip(cols, metrics):
            col.metric(label, value)

        st.dataframe(
            result_df,
            use_container_width=True,
            hide_index=True,
        )

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(
                [
                    {"Metric": k, "Value": v}
                    for k, v in summary.items()
                ]
            ).to_excel(writer, sheet_name="Summary", index=False)

            result_df.to_excel(
                writer,
                sheet_name="PDF Differences",
                index=False,
            )

        st.download_button(
            "Download PDF Report",
            data=output.getvalue(),
            file_name="pdf_comparison_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_pdf_report",
        )


# ============================================================
# Main application
# ============================================================

def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="📄",
        layout="wide",
    )

    init_state()

    st.title("📄 Document Comparator")
    st.caption(
        "Compare Excel, CSV and PDF files with row-order-independent matching."
    )

    file_type = st.radio(
        "Comparison Type",
        ["Excel / CSV", "PDF"],
        horizontal=True,
        key="ui_file_type",
    )

    if file_type == "Excel / CSV":
        col1, col2 = st.columns(2)

        with col1:
            source_file = st.file_uploader(
                "Upload Source File",
                type=["xlsx", "xls", "xlsm", "csv"],
                key="source_file_uploader",
            )

        with col2:
            target_file = st.file_uploader(
                "Upload Target File",
                type=["xlsx", "xls", "xlsm", "csv"],
                key="target_file_uploader",
            )

        # Explicitly swap the two uploader roles. This is useful when
        # Source and Target were uploaded into the opposite placeholders.
        if source_file and target_file:
            if st.button(
                "⇄ Swap Source & Target",
                key="swap_source_target_button",
                help="Swap the roles of the two uploaded files.",
            ):
                st.session_state.swap_files = not st.session_state.swap_files
                st.session_state.pop("last_excel_result", None)
                st.session_state.pop("last_excel_summary", None)
                st.session_state.pop("last_column_summary", None)
                st.session_state.pop("last_duplicate_summary", None)
                st.rerun()

            if st.session_state.swap_files:
                source_file, target_file = target_file, source_file
                st.info(
                    f"Swapped roles are active: **{source_file.name}** is now "
                    f"treated as Source and **{target_file.name}** as Target."
                )

            excel_comparison_ui(source_file, target_file)
        else:
            st.info("Upload both Source and Target files to begin.")

    else:
        col1, col2 = st.columns(2)

        with col1:
            source_file = st.file_uploader(
                "Upload Source PDF",
                type=["pdf"],
                key="source_pdf_uploader",
            )

        with col2:
            target_file = st.file_uploader(
                "Upload Target PDF",
                type=["pdf"],
                key="target_pdf_uploader",
            )

        if source_file and target_file:
            if st.button(
                "⇄ Swap Source & Target",
                key="swap_pdf_source_target_button",
                help="Swap the roles of the two uploaded PDFs.",
            ):
                st.session_state.pop("last_pdf_result", None)
                st.session_state.pop("last_pdf_summary", None)
                st.rerun()

            pdf_comparison_ui(source_file, target_file)
        else:
            st.info("Upload both Source and Target PDFs to begin.")


if __name__ == "__main__":
    main()
