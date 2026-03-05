import copy as _copy
import difflib
import hashlib
import io
import math
import re
import statistics

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.utils import column_index_from_string

_SIZE_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*SQM$", re.IGNORECASE)


# ==========================================
# CORE LOGIC
# ==========================================

def parse_size_header(h):
    """Return numeric SQM size if header like '2.25 SQM', else None."""
    if not isinstance(h, str):
        return None
    m = _SIZE_RE.match(h.strip())
    return float(m.group(1)) if m else None


def to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return None
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s in {"", "-"}:
            return None
        s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return None
    return None


def split_name_address(raw_name):
    """
    Raw competitor string is typically:
      'Brand - Site, 123 Street, Suburb, State Postcode'
    We split on the first comma.
    """
    if not isinstance(raw_name, str) or not raw_name.strip():
        return None, None
    parts = raw_name.split(",", 1)
    name = parts[0].strip()
    addr = parts[1].strip() if len(parts) > 1 else ""
    return name, addr


def find_data_sheet(wb):
    """
    Return the worksheet that contains StoreTrack data by scanning all sheets
    for '12 mo. trailing avg.' in column C. Falls back to the active sheet.
    """
    for ws in wb.worksheets:
        for r in range(5, min(ws.max_row + 1, 200)):
            v = ws.cell(r, 3).value
            if isinstance(v, str) and v.strip().lower().startswith("12 mo. trailing avg"):
                return ws
    return wb.active


def extract_comps_from_raw(raw_file_bytes):
    """
    Reads the raw workbook and returns list of comps:
      {name, address, distance, values{size: val}}

    Assumptions:
      - Size headers row: 4, starting at column D
      - Competitor metric marker in column C: '12 mo. trailing avg.'
      - Name + distance stored in previous row (r-1) columns A and B
    """
    wb = openpyxl.load_workbook(raw_file_bytes, data_only=True)
    ws = find_data_sheet(wb)

    size_col_map = {}
    for c in range(4, ws.max_column + 1):
        sz = parse_size_header(ws.cell(4, c).value)
        if sz is not None:
            size_col_map[sz] = c

    comps = []
    for r in range(5, ws.max_row + 1):
        stat = ws.cell(r, 3).value
        if isinstance(stat, str) and stat.strip().lower().startswith("12 mo. trailing avg"):
            raw_name = ws.cell(r - 1, 1).value
            dist = to_float(ws.cell(r - 1, 2).value)
            name, address = split_name_address(raw_name)

            values = {}
            for sz, c in size_col_map.items():
                v = to_float(ws.cell(r, c).value)
                if v is not None:
                    values[sz] = v

            if values:
                comps.append({
                    "name": name,
                    "address": address,
                    "distance": dist,
                    "values": values,
                })

    return comps


def build_level_row_maps(ws):
    """
    Splits template size rows into Ground Floor and Upper Level purely by
    position. First contiguous group of numeric col-C values = GF,
    second group = UL.

    Returns:
        gf_map       – {sqm_float: row_int}
        ul_map       – {sqm_float: row_int}
        dup_warnings – list of strings for duplicate SQM sizes within a section
    """
    groups: list[list[tuple[float, int]]] = []
    current_group: list[tuple[float, int]] = []

    for r in range(12, ws.max_row + 1):
        sm = to_float(ws.cell(r, 3).value)
        if sm is not None:
            current_group.append((sm, r))
        elif current_group:
            groups.append(current_group)
            current_group = []

    if current_group:
        groups.append(current_group)

    def group_to_map(group):
        m: dict[float, int] = {}
        dups: list[str] = []
        for sm, r in group:
            if sm in m:
                dups.append(
                    f"{sm} SQM appears more than once in the template "
                    f"(rows {m[sm]} and {r}); first occurrence used."
                )
            else:
                m[sm] = r
        return m, dups

    gf_map, gf_dups = group_to_map(groups[0]) if len(groups) >= 1 else ({}, [])
    ul_map, ul_dups = group_to_map(groups[1]) if len(groups) >= 2 else ({}, [])

    return gf_map, ul_map, gf_dups + ul_dups


def fuzzy_match_name(name: str, candidates: dict, threshold: float = 0.85):
    """
    Find the best fuzzy match for `name` among the (normalised) keys of
    `candidates` using difflib.SequenceMatcher.

    Returns (best_key, score) if score >= threshold, else (None, 0.0).
    """
    best_key, best_score = None, 0.0
    for k in candidates:
        score = difflib.SequenceMatcher(None, name, k).ratio()
        if score > best_score:
            best_score = score
            best_key = k
    if best_score >= threshold:
        return best_key, best_score
    return None, 0.0


def align_comps(gf_comps, ul_comps):
    """
    Align GF and UL comp lists by competitor name using a two-pass strategy:
      1. Exact match after normalisation (strip + lowercase).
      2. Fuzzy match via SequenceMatcher (threshold 0.85) for unmatched names
         — generates a verification warning per fuzzy pair.

    Returns:
        slots    – list of {"gf": comp_or_None, "ul": comp_or_None}
        warnings – fuzzy-match and one-sided mismatch warning strings
    """
    def norm(name):
        return (name or "").strip().lower()

    gf_by_name = {norm(c["name"]): c for c in gf_comps}
    ul_by_name = {norm(c["name"]): c for c in ul_comps}

    # Pass 1: exact matches
    gf_to_ul: dict[str, str] = {}
    for k in gf_by_name:
        if k in ul_by_name:
            gf_to_ul[k] = k

    unmatched_gf = [k for k in gf_by_name if k not in gf_to_ul]
    unmatched_ul = set(ul_by_name) - set(gf_to_ul.values())

    # Pass 2: fuzzy matches for remaining unmatched names
    fuzzy_warnings: list[str] = []
    for gf_k in unmatched_gf:
        best_ul_k, score = fuzzy_match_name(gf_k, {k: None for k in unmatched_ul})
        if best_ul_k is not None:
            gf_to_ul[gf_k] = best_ul_k
            unmatched_ul.discard(best_ul_k)
            gf_display = gf_by_name[gf_k]["name"] or gf_k
            ul_display = ul_by_name[best_ul_k]["name"] or best_ul_k
            fuzzy_warnings.append(
                f"Fuzzy match: '{gf_display}' (GF) matched to '{ul_display}' (UL) "
                f"[{score:.0%} similarity] — verify these are the same facility."
            )

    # Build ordered slot list: GF order first, then UL-only extras
    ordered: list[tuple] = []
    seen: set[str] = set()
    for gf_k in gf_by_name:
        if gf_k not in seen:
            seen.add(gf_k)
            ordered.append((gf_k, gf_to_ul.get(gf_k)))

    matched_ul = set(gf_to_ul.values())
    for ul_k in ul_by_name:
        if ul_k not in matched_ul and ul_k not in seen:
            seen.add(ul_k)
            ordered.append((None, ul_k))

    slots: list[dict] = []
    mismatch_warnings: list[str] = []
    for gf_k, ul_k in ordered:
        gf = gf_by_name.get(gf_k) if gf_k else None
        ul = ul_by_name.get(ul_k) if ul_k else None
        display = (gf or ul)["name"] or (gf_k or ul_k)  # type: ignore[index]
        if gf and not ul and ul_comps:
            mismatch_warnings.append(
                f"'{display}' found in Ground Floor data only — Upper Level rates will be blank."
            )
        elif ul and not gf and gf_comps:
            mismatch_warnings.append(
                f"'{display}' found in Upper Level data only — Ground Floor rates will be blank."
            )
        slots.append({"gf": gf, "ul": ul})

    return slots, fuzzy_warnings + mismatch_warnings


def validate_floor_assignment(gf_comps, ul_comps):
    """
    Heuristic checks that GF and UL files were not uploaded to the wrong slots.

    Heuristic 1: GF mean rate is more than 10% below UL mean rate.
                 Ground Floor storage is conventionally priced higher.
    Heuristic 2: GF and UL rates are near-identical across ≥80% of
                 size/competitor intersections (same file uploaded twice).

    Returns a list of warning strings (empty if no issues detected).
    """
    warnings: list[str] = []
    if not gf_comps or not ul_comps:
        return warnings

    gf_vals = [v for c in gf_comps for v in c["values"].values()]
    ul_vals = [v for c in ul_comps for v in c["values"].values()]

    if gf_vals and ul_vals:
        gf_mean = statistics.mean(gf_vals)
        ul_mean = statistics.mean(ul_vals)
        if ul_mean > 0 and gf_mean < ul_mean * 0.90:
            warnings.append(
                f"Ground Floor mean rate (${gf_mean:.2f}) is more than 10% below "
                f"Upper Level (${ul_mean:.2f}). Ground Floor storage is typically "
                "priced higher — verify that files are in the correct uploaders."
            )

    gf_by_name = {(c["name"] or "").strip().lower(): c for c in gf_comps}
    ul_by_name = {(c["name"] or "").strip().lower(): c for c in ul_comps}
    matches = comparisons = 0
    for k in gf_by_name:
        if k in ul_by_name:
            for sz, gv in gf_by_name[k]["values"].items():
                uv = ul_by_name[k]["values"].get(sz)
                if uv is not None:
                    comparisons += 1
                    if abs(gv - uv) < 0.01:
                        matches += 1
    if comparisons >= 3 and matches / comparisons >= 0.80:
        warnings.append(
            f"Ground Floor and Upper Level rates are identical for {matches}/{comparisons} "
            "size/competitor intersections — the same file may have been uploaded to both inputs."
        )

    return warnings


def get_max_comp_slots(ws):
    """Derive how many competitor column slots fit in the sheet from column N onward."""
    first_ask_col = column_index_from_string("N")
    available = (ws.max_column - first_ask_col) // 2 + 1
    return max(1, min(available, 64))


def find_closest_size(sz: float, size_map: dict, tolerance: float = 0.05):
    """
    Find the template row and matched key for the closest SQM value within
    `tolerance`. Returns (row_int, matched_sz) on success, (None, None) if
    nothing is within tolerance.
    """
    if sz in size_map:
        return size_map[sz], sz
    best_key, best_diff = None, float("inf")
    for k in size_map:
        diff = abs(k - sz)
        if diff < best_diff:
            best_diff = diff
            best_key = k
    if best_key is not None and best_diff <= tolerance:
        return size_map[best_key], best_key
    return None, None


def _copy_cell_style(source_cell, target_cell):
    """Copy font, fill, alignment, and border from source to target cell."""
    if source_cell.has_style:
        target_cell.font = _copy.copy(source_cell.font)
        target_cell.fill = _copy.copy(source_cell.fill)
        target_cell.alignment = _copy.copy(source_cell.alignment)
        target_cell.border = _copy.copy(source_cell.border)


def fill_template(template_file_bytes, slots, max_comp_slots=None):
    """
    Writes competitor rates into the template using pre-aligned slots.

    Each slot is {"gf": comp_or_None, "ul": comp_or_None}.

    Sheet: 'Comps & Unit Mix'
      - Competitor name  → row 3,  columns N, P, R … (every 2 columns)
      - Competitor addr  → row 4,  same columns
      - Distance         → row 9,  same columns
      - Selected flag    → row 10, same columns  ('Yes')
      - Ground Floor rates → first contiguous SQM group in col C (rows 12+)
      - Upper Level rates  → second contiguous SQM group in col C (rows 12+)

    Rate cells receive currency formatting ($#,##0.00).
    Header rows 3/4/9 copy their style from the adjacent column M reference cell.
    Size matching uses a ±0.05 SQM tolerance to handle minor float differences.

    Returns:
        out       – BytesIO of the filled workbook
        n_written – number of competitor slots written
        warnings  – combined duplicate-SQM and size-match warning strings
    """
    wb = openpyxl.load_workbook(template_file_bytes)
    if "Comps & Unit Mix" not in wb.sheetnames:
        raise ValueError("Template workbook must contain a sheet named 'Comps & Unit Mix'.")

    ws = wb["Comps & Unit Mix"]
    gf_map, ul_map, dup_warnings = build_level_row_maps(ws)

    if max_comp_slots is None:
        max_comp_slots = get_max_comp_slots(ws)

    first_ask_col = column_index_from_string("N")
    subject_col = column_index_from_string("M")
    n_written = 0
    size_match_warnings: list[str] = []

    for i, slot in enumerate(slots[:max_comp_slots]):
        ask_col = first_ask_col + 2 * i
        gf_comp = slot["gf"]
        ul_comp = slot["ul"]
        identity = gf_comp or ul_comp
        comp_name = identity.get("name") or f"Competitor {i + 1}"

        # Copy style from column M reference cell for header rows
        for hr in (3, 4, 9):
            _copy_cell_style(ws.cell(hr, subject_col), ws.cell(hr, ask_col))

        ws.cell(3, ask_col).value = identity.get("name")
        ws.cell(4, ask_col).value = identity.get("address")
        if identity.get("distance") is not None:
            ws.cell(9, ask_col).value = identity["distance"]
        ws.cell(10, ask_col).value = "Yes"

        if gf_comp:
            for sz, val in gf_comp["values"].items():
                row, matched_sz = find_closest_size(sz, gf_map)
                if row is not None:
                    ws.cell(row, ask_col).value = val
                    ws.cell(row, ask_col).number_format = '"$"#,##0.00'
                    if matched_sz is not None and abs(matched_sz - sz) > 1e-9:
                        size_match_warnings.append(
                            f"GF — {comp_name}: {sz} SQM matched to {matched_sz} SQM template row."
                        )

        if ul_comp:
            for sz, val in ul_comp["values"].items():
                row, matched_sz = find_closest_size(sz, ul_map)
                if row is not None:
                    ws.cell(row, ask_col).value = val
                    ws.cell(row, ask_col).number_format = '"$"#,##0.00'
                    if matched_sz is not None and abs(matched_sz - sz) > 1e-9:
                        size_match_warnings.append(
                            f"UL — {comp_name}: {sz} SQM matched to {matched_sz} SQM template row."
                        )

        n_written += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, n_written, dup_warnings + size_match_warnings


# ==========================================
# STREAMLIT UI HELPERS
# ==========================================

def _file_signature(f) -> str | None:
    """Stable fingerprint for an uploaded file (name + size). Returns None if no file."""
    if f is None:
        return None
    return hashlib.md5((f.name + str(f.size)).encode()).hexdigest()


def _render_comp_preview(comps: list, label: str):
    """Render competitor data as a dataframe in a collapsible expander."""
    if not comps:
        return
    all_sizes = sorted({sz for c in comps for sz in c["values"]})
    rows = []
    for c in comps:
        row: dict = {
            "Name": c["name"] or "Unknown",
            "Address": c["address"] or "—",
            "Distance (km)": c["distance"] if c["distance"] is not None else "N/A",
        }
        for sz in all_sizes:
            val = c["values"].get(sz)
            row[f"{sz} SQM"] = f"${val:.2f}" if val is not None else "—"
        rows.append(row)
    with st.expander(f"Preview — {label} ({len(comps)} competitor(s))"):
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def _render_slot_editor(slots: list) -> list:
    """
    Let the analyst include/exclude and reorder competitors before generating.
    Renders a st.data_editor with Include (checkbox) and Order (number) editable.
    Returns the filtered and reordered slot list.
    """
    rows = []
    for i, slot in enumerate(slots, 1):
        identity = slot["gf"] or slot["ul"]
        rows.append({
            "Include": True,
            "Order": i,
            "Name": identity.get("name") or "Unknown",
            "Distance (km)": (
                identity["distance"] if identity.get("distance") is not None else "N/A"
            ),
            "GF Sizes (SQM)": (
                ", ".join(str(sz) for sz in sorted(slot["gf"]["values"]))
                if slot["gf"] else "—"
            ),
            "UL Sizes (SQM)": (
                ", ".join(str(sz) for sz in sorted(slot["ul"]["values"]))
                if slot["ul"] else "—"
            ),
        })

    df = pd.DataFrame(rows)
    with st.expander("Select and Order Competitors", expanded=True):
        st.caption(
            "Tick/untick **Include** to add or remove competitors. "
            "Edit **Order** to change the column position in the template."
        )
        edited = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Include": st.column_config.CheckboxColumn("Include", default=True),
                "Order": st.column_config.NumberColumn("Order", min_value=1, step=1),
                "Name": st.column_config.TextColumn("Name", disabled=True),
                "Distance (km)": st.column_config.TextColumn("Distance (km)", disabled=True),
                "GF Sizes (SQM)": st.column_config.TextColumn("GF Sizes (SQM)", disabled=True),
                "UL Sizes (SQM)": st.column_config.TextColumn("UL Sizes (SQM)", disabled=True),
            },
            num_rows="fixed",
        )

    included = edited[edited["Include"]].sort_values("Order")
    return [slots[i] for i in included.index.tolist()]


def _render_market_summary(slots: list, gf_map: dict, ul_map: dict):
    """
    Display min/max/mean/median per SQM size across all competitors and flag
    any value more than 2 standard deviations from the mean as an outlier.
    """
    st.subheader("Market Rate Summary")

    def gather_vals(size_map: dict, level_key: str) -> dict[float, list[float]]:
        data: dict[float, list[float]] = {}
        for sz in size_map:
            for slot in slots:
                comp = slot.get(level_key)
                if not comp:
                    continue
                v = comp["values"].get(sz)
                if v is None:
                    for k, kv in comp["values"].items():
                        if abs(k - sz) <= 0.05:
                            v = kv
                            break
                if v is not None:
                    data.setdefault(sz, []).append(v)
        return data

    def build_stats_df(data: dict[float, list[float]]) -> tuple[pd.DataFrame, list[str]]:
        rows, outlier_msgs = [], []
        for sz in sorted(data):
            vals = data[sz]
            avg = statistics.mean(vals)
            med = statistics.median(vals)
            rows.append({
                "Size (SQM)": sz,
                "Count": len(vals),
                "Min ($)": f"${min(vals):.2f}",
                "Max ($)": f"${max(vals):.2f}",
                "Mean ($)": f"${avg:.2f}",
                "Median ($)": f"${med:.2f}",
            })
            if len(vals) >= 3:
                try:
                    sd = statistics.stdev(vals)
                    for v in vals:
                        if abs(v - avg) > 2 * sd:
                            outlier_msgs.append(
                                f"{sz} SQM: ${v:.2f} is more than 2 standard deviations "
                                f"from the mean (${avg:.2f} ± ${sd:.2f})"
                            )
                except statistics.StatisticsError:
                    pass
        return pd.DataFrame(rows), outlier_msgs

    gf_data = gather_vals(gf_map, "gf")
    ul_data = gather_vals(ul_map, "ul")
    all_outliers: list[str] = []

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Ground Floor**")
        if gf_data:
            df, msgs = build_stats_df(gf_data)
            st.dataframe(df, use_container_width=True, hide_index=True)
            all_outliers.extend(f"GF — {m}" for m in msgs)
        else:
            st.info("No Ground Floor rate data available.")

    with col2:
        st.markdown("**Upper Level**")
        if ul_data:
            df, msgs = build_stats_df(ul_data)
            st.dataframe(df, use_container_width=True, hide_index=True)
            all_outliers.extend(f"UL — {m}" for m in msgs)
        else:
            st.info("No Upper Level rate data available.")

    for msg in all_outliers:
        st.warning(f"Outlier detected: {msg}")


# ==========================================
# STREAMLIT APP
# ==========================================

def main():
    st.set_page_config(
        page_title="Market Rent Transposer",
        layout="wide",
        page_icon="🏢",
    )
    st.title("Market Rent Transposer")
    st.write(
        "Upload the Ground Floor and Upper Level StoreTrack exports plus the "
        "Market Rent Analysis template. The app will extract competitor rates "
        "from each file and insert them into the matching floor-level rows of "
        "the template."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("1. Ground Floor Raw Data")
        gf_file = st.file_uploader(
            "Upload Ground Floor StoreTrack export (.xlsx)",
            type=["xlsx"],
            key="gf",
            help="Raw workbook with size headers in row 4 (column D onward) "
                 "and '12 mo. trailing avg.' rows per competitor — Ground Floor pricing.",
        )

    with col2:
        st.subheader("2. Upper Level Raw Data")
        ul_file = st.file_uploader(
            "Upload Upper Level StoreTrack export (.xlsx)",
            type=["xlsx"],
            key="ul",
            help="Same format as the Ground Floor file — Upper Level pricing.",
        )

    with col3:
        st.subheader("3. Template")
        template_file = st.file_uploader(
            "Upload Market Rent Analysis template (.xlsx)",
            type=["xlsx"],
            key="tmpl",
            help="Must contain a sheet named 'Comps & Unit Mix'. "
                 "Ground Floor rows are the first contiguous SQM block in column C; "
                 "Upper Level rows are the second.",
        )

    st.markdown("---")

    if not template_file:
        st.info("Please upload the template file to continue.")
        return
    if not gf_file and not ul_file:
        st.info("Please upload at least one raw data file (Ground Floor or Upper Level).")
        return

    # ---- Stale session state invalidation ----
    # Clear any previous result when the uploaded files change.
    current_sig = (
        _file_signature(gf_file),
        _file_signature(ul_file),
        _file_signature(template_file),
    )
    if st.session_state.get("file_sig") != current_sig:
        for key in ("result", "n_written", "dup_warnings"):
            st.session_state.pop(key, None)
        st.session_state["file_sig"] = current_sig

    # ---- Wrap uploads in BytesIO immediately ----
    gf_bytes = io.BytesIO(gf_file.read()) if gf_file else None
    ul_bytes = io.BytesIO(ul_file.read()) if ul_file else None
    tmpl_bytes = io.BytesIO(template_file.read())

    # ---- Extract comps ----
    gf_comps: list = []
    ul_comps: list = []

    if gf_bytes:
        with st.spinner("Reading Ground Floor data..."):
            try:
                gf_comps = extract_comps_from_raw(gf_bytes)
            except Exception as e:
                st.error(f"Could not read Ground Floor file: {e}")
                return
        if not gf_comps:
            st.warning(
                "No competitors found in the Ground Floor file. "
                "Check that row 4 has SQM headers and column C has '12 mo. trailing avg.' rows."
            )
        else:
            st.success(f"Ground Floor: found **{len(gf_comps)}** competitor(s).")
            _render_comp_preview(gf_comps, "Ground Floor")

    if ul_bytes:
        with st.spinner("Reading Upper Level data..."):
            try:
                ul_comps = extract_comps_from_raw(ul_bytes)
            except Exception as e:
                st.error(f"Could not read Upper Level file: {e}")
                return
        if not ul_comps:
            st.warning(
                "No competitors found in the Upper Level file. "
                "Check that row 4 has SQM headers and column C has '12 mo. trailing avg.' rows."
            )
        else:
            st.success(f"Upper Level: found **{len(ul_comps)}** competitor(s).")
            _render_comp_preview(ul_comps, "Upper Level")

    if not gf_comps and not ul_comps:
        st.error("No competitor data could be extracted from either file.")
        return

    # ---- Align competitors by name (exact + fuzzy) ----
    slots, align_warnings = align_comps(gf_comps, ul_comps)
    for w in align_warnings:
        st.warning(w)

    # ---- Floor assignment validation ----
    for w in validate_floor_assignment(gf_comps, ul_comps):
        st.warning(w)

    # ---- Peek at template: extract row maps and max slots ----
    gf_map: dict = {}
    ul_map: dict = {}
    max_slots = 16
    try:
        tmpl_bytes.seek(0)
        wb_peek = openpyxl.load_workbook(tmpl_bytes, data_only=True)
        if "Comps & Unit Mix" in wb_peek.sheetnames:
            ws_peek = wb_peek["Comps & Unit Mix"]
            max_slots = get_max_comp_slots(ws_peek)
            gf_map, ul_map, _ = build_level_row_maps(ws_peek)
        tmpl_bytes.seek(0)
    except Exception:
        tmpl_bytes.seek(0)

    st.markdown("---")

    # ---- Competitor selection & reordering ----
    slots = _render_slot_editor(slots)

    if not slots:
        st.warning("No competitors selected. Tick at least one competitor to generate the report.")
        return

    if len(slots) > max_slots:
        st.warning(
            f"{max_slots} of {len(slots)} competitors will be written — "
            f"the template only has {max_slots} competitor column slot(s). "
            "Deselect extras in the table above."
        )

    # ---- Market rate summary + outlier detection ----
    if gf_map or ul_map:
        st.markdown("---")
        _render_market_summary(slots, gf_map, ul_map)

    # ---- Generate ----
    st.markdown("---")
    if st.button("Generate Filled Template", type="primary"):
        with st.spinner("Filling template..."):
            try:
                result, n_written, all_warnings = fill_template(tmpl_bytes, slots)
            except ValueError as e:
                st.error(str(e))
                return
            except Exception as e:
                st.error(f"Error filling template: {e}")
                return

        st.session_state["result"] = result
        st.session_state["n_written"] = n_written
        st.session_state["dup_warnings"] = all_warnings

    # Render download button outside the generate block so that clicking it
    # does not re-trigger generation on the script rerun Streamlit performs.
    if "result" in st.session_state:
        for w in st.session_state["dup_warnings"]:
            st.warning(w)
        st.success(f"Template filled with **{st.session_state['n_written']}** competitor(s).")
        st.download_button(
            label="Download Filled Template",
            data=st.session_state["result"],
            file_name="Market_Rent_Analysis_Filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
