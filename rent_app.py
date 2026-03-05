import io
import math
import re

import openpyxl
import streamlit as st
from openpyxl.utils import column_index_from_string


# ==========================================
# CORE LOGIC
# ==========================================

def parse_size_header(h):
    """Return numeric SQM size if header like '2.25 SQM', else None."""
    if not isinstance(h, str):
        return None
    h = h.strip()
    m = re.match(r"^(\d+(?:\.\d+)?)\s*SQM$", h, flags=re.IGNORECASE)
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
    ws = wb.active

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
        gf_map  – {sqm_float: row_int} for the first (Ground Floor) section
        ul_map  – {sqm_float: row_int} for the second (Upper Level) section
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

    gf_map: dict[float, int] = {}
    ul_map: dict[float, int] = {}

    if len(groups) >= 1:
        for sm, r in groups[0]:
            gf_map[sm] = r
    if len(groups) >= 2:
        for sm, r in groups[1]:
            ul_map[sm] = r

    return gf_map, ul_map


def fill_template(template_file_bytes, gf_comps, ul_comps, max_comp_slots=16):
    """
    Writes Ground Floor and Upper Level competitor rates into the template.

    Sheet: 'Comps & Unit Mix'
      - Competitor name  → row 3,  columns N, P, R … (every 2 columns)
      - Competitor addr  → row 4,  same columns
      - Distance         → row 9,  same columns
      - Selected flag    → row 10, same columns  ('Yes')
      - Ground Floor rates → first contiguous SQM group in col C (rows 12+)
      - Upper Level rates  → second contiguous SQM group in col C (rows 12+)

    Competitor identity (name / address / distance) is taken from whichever
    level has data for that slot; Ground Floor takes priority.
    """
    wb = openpyxl.load_workbook(template_file_bytes)
    if "Comps & Unit Mix" not in wb.sheetnames:
        raise ValueError("Template workbook must contain a sheet named 'Comps & Unit Mix'.")

    ws = wb["Comps & Unit Mix"]

    gf_map, ul_map = build_level_row_maps(ws)

    if not gf_map and not ul_map:
        # Fallback: no section labels found – treat all size rows as a single
        # pool and write GF comps (preserves original single-level behaviour).
        fallback_map: dict[float, int] = {}
        for r in range(12, ws.max_row + 1):
            sm = to_float(ws.cell(r, 3).value)
            if sm is not None:
                fallback_map[sm] = r
        gf_map = fallback_map

    first_ask_col = column_index_from_string("N")
    n_slots = max(len(gf_comps), len(ul_comps))

    for i in range(min(n_slots, max_comp_slots)):
        ask_col = first_ask_col + 2 * i

        gf_comp = gf_comps[i] if i < len(gf_comps) else None
        ul_comp = ul_comps[i] if i < len(ul_comps) else None

        # Header info: prefer GF, fall back to UL
        identity = gf_comp or ul_comp
        ws.cell(3, ask_col).value = identity.get("name")
        ws.cell(4, ask_col).value = identity.get("address")
        if identity.get("distance") is not None:
            ws.cell(9, ask_col).value = identity["distance"]
        ws.cell(10, ask_col).value = "Yes"

        # Ground Floor rates
        if gf_comp:
            for sz, val in gf_comp["values"].items():
                row = gf_map.get(sz)
                if row is not None:
                    ws.cell(row, ask_col).value = val

        # Upper Level rates
        if ul_comp:
            for sz, val in ul_comp["values"].items():
                row = ul_map.get(sz)
                if row is not None:
                    ws.cell(row, ask_col).value = val

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ==========================================
# STREAMLIT APP
# ==========================================

def _render_comp_preview(comps: list, label: str):
    """Render a collapsible competitor table for one level."""
    with st.expander(f"Preview — {label} ({len(comps)} competitor(s))"):
        for i, c in enumerate(comps, 1):
            sizes = ", ".join(
                f"{sz} SQM: ${v:.2f}" for sz, v in sorted(c["values"].items())
            )
            dist_str = f"{c['distance']} km" if c["distance"] is not None else "N/A"
            st.markdown(
                f"**{i}. {c['name'] or 'Unknown'}**  \n"
                f"Address: {c['address'] or '—'}  \n"
                f"Distance: {dist_str}  \n"
                f"Rates: {sizes}"
            )


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
            help="The raw workbook with size headers in row 4 (column D onward) "
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
                 "Ground Floor and Upper Level sections are detected automatically "
                 "by looking for those labels in column A.",
        )

    st.markdown("---")

    # At least one raw file and the template are required to proceed
    have_raw = gf_file or ul_file
    if not have_raw or not template_file:
        missing = []
        if not gf_file:
            missing.append("Ground Floor raw data")
        if not ul_file:
            missing.append("Upper Level raw data")
        if not template_file:
            missing.append("template")
        st.info(f"Please upload: {', '.join(missing)}.")
        return

    # ---- Extract comps ----
    gf_comps: list = []
    ul_comps: list = []

    if gf_file:
        with st.spinner("Reading Ground Floor data..."):
            try:
                gf_comps = extract_comps_from_raw(io.BytesIO(gf_file.read()))
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

    if ul_file:
        with st.spinner("Reading Upper Level data..."):
            try:
                ul_comps = extract_comps_from_raw(io.BytesIO(ul_file.read()))
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

    # ---- Generate ----
    if st.button("Generate Filled Template", type="primary"):
        with st.spinner("Filling template..."):
            try:
                result = fill_template(
                    io.BytesIO(template_file.read()), gf_comps, ul_comps
                )
            except ValueError as e:
                st.error(str(e))
                return
            except Exception as e:
                st.error(f"Error filling template: {e}")
                return

        st.success("Template filled successfully.")
        st.download_button(
            label="Download Filled Template",
            data=result,
            file_name="Market_Rent_Analysis_Filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
