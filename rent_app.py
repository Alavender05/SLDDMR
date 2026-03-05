import io
import math
import re

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

    Assumptions (matches your current raw template):
      - Size headers row: 4, starting at column D
      - Competitor metric marker in column C: '12 mo. trailing avg.'
      - Name + distance stored in previous row (r-1) columns A and B
    """
    wb = openpyxl.load_workbook(raw_file_bytes, data_only=True)
    ws = find_data_sheet(wb)

    # Map size -> column index
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
                comps.append(
                    {
                        "name": name,
                        "address": address,
                        "distance": dist,
                        "values": values,
                    }
                )

    return comps


def build_level_row_maps(ws):
    """
    Splits template size rows into Ground Floor and Upper Level purely by
    position — no label scanning required.

    Scans column C from row 12 onward and collects contiguous groups of rows
    that contain a numeric SQM value (gaps of non-numeric rows separate groups).
    The FIRST group becomes the Ground Floor map; the SECOND group becomes the
    Upper Level map.  This matches whichever file the user placed in the
    'Ground Floor' uploader vs the 'Upper Level' uploader.

    Returns:
        gf_map       – {sqm_float: row_int} for the first (Ground Floor) section
        ul_map       – {sqm_float: row_int} for the second (Upper Level) section
        dup_warnings – list of strings for any duplicate SQM sizes within a section
    """
    # Collect contiguous groups of (sqm, row) pairs
    groups: list[list[tuple[float, int]]] = []
    current_group: list[tuple[float, int]] = []

    for r in range(12, ws.max_row + 1):
        sm = to_float(ws.cell(r, 3).value)
        if sm is not None:
            current_group.append((sm, r))
        else:
            if current_group:
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


def align_comps(gf_comps, ul_comps):
    """
    Align Ground Floor and Upper Level comp lists by competitor name so that
    the same facility ends up in the same column slot regardless of ordering.

    Returns:
        slots    – list of {"gf": comp_or_None, "ul": comp_or_None}
        warnings – list of strings for competitors that appear in only one file
    """
    def norm(name):
        return (name or "").strip().lower()

    gf_by_name = {norm(c["name"]): c for c in gf_comps}
    ul_by_name = {norm(c["name"]): c for c in ul_comps}

    # Preserve order: GF order first, then any UL-only extras
    seen: dict[str, bool] = {}
    ordered_keys: list[str] = []
    for c in gf_comps:
        k = norm(c["name"])
        if k not in seen:
            seen[k] = True
            ordered_keys.append(k)
    for c in ul_comps:
        k = norm(c["name"])
        if k not in seen:
            seen[k] = True
            ordered_keys.append(k)

    slots: list[dict] = []
    warnings: list[str] = []
    for k in ordered_keys:
        gf = gf_by_name.get(k)
        ul = ul_by_name.get(k)
        display = (gf or ul)["name"] or k  # type: ignore[index]
        if gf and not ul and ul_comps:
            warnings.append(f"'{display}' found in Ground Floor data only — Upper Level rates will be blank for this competitor.")
        elif ul and not gf and gf_comps:
            warnings.append(f"'{display}' found in Upper Level data only — Ground Floor rates will be blank for this competitor.")
        slots.append({"gf": gf, "ul": ul})

    return slots, warnings


def get_max_comp_slots(ws):
    """Derive how many competitor column slots fit in the sheet from column N onward."""
    first_ask_col = column_index_from_string("N")
    available = (ws.max_column - first_ask_col) // 2 + 1
    return max(1, min(available, 64))


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

    Returns:
        out          – BytesIO of the filled workbook
        n_written    – number of competitor slots written
        dup_warnings – list of duplicate-SQM warning strings from the template
    """
    wb = openpyxl.load_workbook(template_file_bytes)
    if "Comps & Unit Mix" not in wb.sheetnames:
        raise ValueError("Template workbook must contain a sheet named 'Comps & Unit Mix'.")

    ws = wb["Comps & Unit Mix"]
    gf_map, ul_map, dup_warnings = build_level_row_maps(ws)

    if max_comp_slots is None:
        max_comp_slots = get_max_comp_slots(ws)

    first_ask_col = column_index_from_string("N")
    n_written = 0

    for i, slot in enumerate(slots[:max_comp_slots]):
        ask_col = first_ask_col + 2 * i
        gf_comp = slot["gf"]
        ul_comp = slot["ul"]

        identity = gf_comp or ul_comp
        ws.cell(3, ask_col).value = identity.get("name")
        ws.cell(4, ask_col).value = identity.get("address")
        if identity.get("distance") is not None:
            ws.cell(9, ask_col).value = identity["distance"]
        ws.cell(10, ask_col).value = "Yes"

        if gf_comp:
            for sz, val in gf_comp["values"].items():
                row = gf_map.get(sz)
                if row is not None:
                    ws.cell(row, ask_col).value = val

        if ul_comp:
            for sz, val in ul_comp["values"].items():
                row = ul_map.get(sz)
                if row is not None:
                    ws.cell(row, ask_col).value = val

        n_written += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, n_written, dup_warnings


# ==========================================
# STREAMLIT APP
# ==========================================

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

    # Template is always required; at least one raw file is required
    if not template_file:
        st.info("Please upload the template file to continue.")
        return
    if not gf_file and not ul_file:
        st.info("Please upload at least one raw data file (Ground Floor or Upper Level).")
        return

    # Wrap uploaded files in BytesIO immediately so they can be read multiple
    # times within the same script run without exhausting the upload buffer.
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

    # ---- Align competitors by name ----
    slots, align_warnings = align_comps(gf_comps, ul_comps)
    for w in align_warnings:
        st.warning(w)

    # ---- Slot count warning (peek at template without consuming tmpl_bytes) ----
    try:
        tmpl_bytes.seek(0)
        wb_peek = openpyxl.load_workbook(tmpl_bytes, data_only=True)
        if "Comps & Unit Mix" in wb_peek.sheetnames:
            max_slots = get_max_comp_slots(wb_peek["Comps & Unit Mix"])
            if len(slots) > max_slots:
                st.warning(
                    f"{max_slots} of {len(slots)} competitors will be written — "
                    f"the template only has {max_slots} competitor column slot(s). "
                    "Extra competitors will be ignored."
                )
        tmpl_bytes.seek(0)
    except Exception:
        tmpl_bytes.seek(0)

    # ---- Generate ----
    if st.button("Generate Filled Template", type="primary"):
        with st.spinner("Filling template..."):
            try:
                result, n_written, dup_warnings = fill_template(tmpl_bytes, slots)
            except ValueError as e:
                st.error(str(e))
                return
            except Exception as e:
                st.error(f"Error filling template: {e}")
                return

        st.session_state["result"] = result
        st.session_state["n_written"] = n_written
        st.session_state["dup_warnings"] = dup_warnings

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
