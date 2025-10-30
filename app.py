import io
import re
import streamlit as st
from typing import List, Set, Optional
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Excel Image In-Cell Preview", layout="wide")

HTTP_URL_RE = re.compile(r'^(https?://)?([A-Za-z0-9\.\-]+\.[A-Za-z]{2,})(/.*)$', re.IGNORECASE)
LIKELY_CDN_RE = re.compile(r'^(cdn\.|media\.|images\.|static\.)', re.IGNORECASE)
DEFAULT_SCHEME = "https://"

def is_url_like(s: str) -> bool:
    if not isinstance(s, str): return False
    s = s.strip()
    if not s: return False
    if s.lower().startswith(('http://','https://')): return True
    if HTTP_URL_RE.match(s): return True
    if LIKELY_CDN_RE.match(s): return True
    return False

def normalize_url(s: str, default_scheme: str = DEFAULT_SCHEME) -> Optional[str]:
    if not s: return None
    s = s.strip().strip('"').strip("'")
    if s.lower().startswith(('http://','https://')): return s
    m = HTTP_URL_RE.match(s)
    if m:
        scheme, host, path = m.group(1), m.group(2), m.group(3)
        scheme = scheme or default_scheme
        return f"{scheme}{host}{path}"
    if LIKELY_CDN_RE.match(s):
        if '/' in s:
            host, path = s.split('/', 1)
            return f"{default_scheme}{host}/{path}"
        return f"{default_scheme}{s}"
    return None

def px_to_col_width(px:int)->float: return round(px/7.0,2)
def px_to_row_height(px:int)->float: return round(px*0.75,2)

def detect_url_columns(ws, header_row:int=1)->List[int]:
    urls_in_col = []
    max_col = ws.max_column
    max_row = ws.max_row
    for col_idx in range(1, max_col + 1):
        for row_idx in range(header_row + 1, min(max_row, header_row + 50) + 1):
            v = ws.cell(row=row_idx, column=col_idx).value
            if isinstance(v, str) and is_url_like(v):
                urls_in_col.append(col_idx)
                break
    return urls_in_col

def columns_by_names(ws, names:List[str], header_row:int=1)->List[int]:
    name_map = {}
    for col_idx in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=col_idx).value
        if isinstance(v, str):
            name_map[v.strip()] = col_idx
    cols = []
    for n in names:
        if n in name_map:
            cols.append(name_map[n])
    return cols

def iter_target_cells(ws, target_cols:Set[int], header_row:int):
    for row in range(header_row + 1, ws.max_row + 1):
        for c in target_cols:
            yield ws.cell(row=row, column=c)

def place_image_formula(cell, url: str, w_px: int, h_px: int, keep_note: bool = True):
    if keep_note:
        cell.comment = Comment(f"Original URL:\n{url}", "PreviewBot")
    cell.value = f'=IMAGE("{url}",,3,{w_px},{h_px})'

def adjust_dimensions(ws, col_indices:Set[int], row_height_px:int):
    for c in col_indices:
        ws.column_dimensions[get_column_letter(c)].width = px_to_col_width(row_height_px)
    target_height_pt = px_to_row_height(row_height_px)
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = target_height_pt

st.title("Excel Image In-Cell Preview")
st.caption("Preview images *inside the same cells* in your masterfile using Excel's IMAGE() — keep SKU, SEO, Title, GTIN columns unchanged.")

uploaded = st.file_uploader("Upload your Excel masterfile (.xlsx)", type=["xlsx"])

if uploaded:
    # Load workbook to memory
    data = uploaded.read()
    wb = load_workbook(io.BytesIO(data))
    sheets = wb.sheetnames

    st.sidebar.header("Settings")
    sheet_mode = st.sidebar.radio("Which sheets?", ["One sheet", "All sheets"], horizontal=True)
    if sheet_mode == "One sheet":
        sheet_name = st.sidebar.selectbox("Sheet to process", sheets, index=0)
        target_sheets = [sheet_name]
    else:
        target_sheets = sheets

    header_row = st.sidebar.number_input("Header row number", min_value=1, value=1, step=1)
    width = st.sidebar.number_input("Image width (px)", min_value=40, value=140, step=10)
    height = st.sidebar.number_input("Image height (px)", min_value=40, value=140, step=10)
    keep_notes = st.sidebar.checkbox("Keep original URL as a cell note", value=True)
    create_adjacent = st.sidebar.checkbox("Create preview in NEW adjacent column(s) (keep URLs intact)", value=False)

    st.write("### Column Selection")
    ws0 = wb[target_sheets[0]]
    # Suggest likely column names from header row
    header_vals = []
    for c in range(1, ws0.max_column + 1):
        v = ws0.cell(row=header_row, column=c).value
        if isinstance(v, str):
            header_vals.append(v.strip())
        else:
            header_vals.append(f"Col {c}")

    # Auto-detect URL columns on the first sheet for convenience
    auto_cols_idx = detect_url_columns(ws0, header_row=header_row)
    auto_cols_names = [header_vals[i-1] if i-1 < len(header_vals) else f"Col {i}" for i in auto_cols_idx]

    selected_by_name = st.multiselect(
        "Pick URL columns by header (auto-detected suggestions pre-selected)",
        options=header_vals,
        default=auto_cols_names
    )

    # Helper: map headers to indices
    def headers_to_indices(ws, names:List[str])->Set[int]:
        return set(columns_by_names(ws, names, header_row=header_row))

    # Preview counts per sheet
    st.write("### Preview")
    info_rows = []
    for s in target_sheets:
        ws = wb[s]
        targets = headers_to_indices(ws, selected_by_name) if selected_by_name else set(detect_url_columns(ws, header_row=header_row))
        count = 0
        for cell in iter_target_cells(ws, targets, header_row=header_row):
            v = cell.value
            if isinstance(v, str) and is_url_like(v):
                count += 1
        info_rows.append((s, len(targets), count))
    st.dataframe({"Sheet": [r[0] for r in info_rows],
                  "Target columns": [r[1] for r in info_rows],
                  "URL cells found": [r[2] for r in info_rows]},
                 use_container_width=True)

    if st.button("Process & prepare download"):
        changed_sheets = 0
        for s in target_sheets:
            ws = wb[s]
            targets = headers_to_indices(ws, selected_by_name) if selected_by_name else set(detect_url_columns(ws, header_row=header_row))
            if not targets:
                continue

            # If creating adjacent preview cols, insert new columns next to each target
            preview_targets = set()
            if create_adjacent:
                # Insert to the right of each target column, accounting for shifts
                inserted = 0
                for col_idx in sorted(targets):
                    insert_at = col_idx + 1 + inserted
                    ws.insert_cols(insert_at)
                    # Copy header name with suffix
                    header_cell = ws.cell(row=header_row, column=col_idx + inserted)
                    new_header_cell = ws.cell(row=header_row, column=insert_at)
                    base = header_cell.value if isinstance(header_cell.value, str) else f"Col {col_idx}"
                    new_header_cell.value = f"{base}_preview"
                    preview_targets.add(insert_at)
                    inserted += 1
            else:
                preview_targets = targets

            # Resize rows/cols for the preview columns only
            adjust_dimensions(ws, preview_targets, row_height_px=max(width, height))

            # Write IMAGE() formulas into chosen/created columns
            processed = 0
            for cell in iter_target_cells(ws, preview_targets, header_row=header_row):
                # Determine the source URL: if adjacent preview, read from left neighbor
                src_cell = cell if not create_adjacent else ws.cell(row=cell.row, column=cell.column - 1)
                val = src_cell.value
                if isinstance(val, str) and is_url_like(val):
                    url = normalize_url(val) or val.strip()
                    place_image_formula(cell, url, width, height, keep_note=keep_notes and not create_adjacent)
                    processed += 1

            changed_sheets += 1

        # Save to buffer
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.success("Done! Download your updated Excel file below.")
        st.download_button(
            label="⬇️ Download updated masterfile",
            data=buf,
            file_name="master_with_image_previews.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Upload your .xlsx masterfile to begin.")
