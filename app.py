import io
import re
import requests
import streamlit as st
from typing import List, Set, Optional
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Pattern Walmart/Target Image Preview", layout="wide")

# --- URL helpers ---
HTTP_URL_RE = re.compile(r'^(https?://)?([A-Za-z0-9\.\-]+\.[A-Za-z]{2,})(/.*)$', re.IGNORECASE)
LIKELY_CDN_RE = re.compile(r'^(cdn\.|media\.|images\.|static\.)', re.IGNORECASE)
DEFAULT_SCHEME = "https://"

def is_url_like(s: str) -> bool:
    if not isinstance(s, str): return False
    s = s.strip()
    if not s: return False
    if s.lower().startswith(("http://", "https://")): return True
    if HTTP_URL_RE.match(s): return True
    if LIKELY_CDN_RE.match(s): return True
    return False

def normalize_url(s: str, default_scheme: str = DEFAULT_SCHEME) -> Optional[str]:
    if not s: return None
    s = s.strip().strip('"').strip("'")
    if s.lower().startswith(("http://", "https://")): return s
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

# --- Excel helpers ---
def px_to_col_width(px:int)->float: return round(px/7.0, 2)
def px_to_row_height(px:int)->float: return round(px*0.75, 2)

def detect_url_columns(ws, header_row:int=1)->List[int]:
    hits = []
    for c in range(1, ws.max_column + 1):
        for r in range(header_row + 1, min(ws.max_row, header_row + 50) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and is_url_like(v):
                hits.append(c); break
    return hits

def columns_by_names(ws, names:List[str], header_row:int=1)->List[int]:
    name_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str): name_map[v.strip()] = c
    return [name_map[n] for n in names if n in name_map]

def iter_target_cells(ws, target_cols:Set[int], header_row:int):
    for r in range(header_row + 1, ws.max_row + 1):
        for c in target_cols:
            yield ws.cell(row=r, column=c)

def adjust_dimensions(ws, col_indices:Set[int], row_height_px:int):
    for c in col_indices:
        ws.column_dimensions[get_column_letter(c)].width = px_to_col_width(row_height_px)
    target_h_pt = px_to_row_height(row_height_px)
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = target_h_pt

def place_image_formula(cell, url: str, w_px: int, h_px: int, keep_note: bool):
    if keep_note: cell.comment = Comment(f"Original URL:\n{url}", "PreviewBot")
    cell.value = f'=IMAGE("{url}",,3,{w_px},{h_px})'

def place_anchor_image(ws, cell, url: str, w_px: int, h_px: int, keep_note: bool):
    resp = requests.get(url, timeout=25)
    resp.raise_for_status()
    data = io.BytesIO(resp.content)
    img = XLImage(data)
    img.width = w_px
    img.height = h_px
    img.anchor = cell.coordinate
    ws.add_image(img)
    if keep_note: cell.comment = Comment(f"Original URL:\n{url}", "PreviewBot")

# --- UI ---
st.title("Pattern Walmart & Target — Image Preview in Excel")
st.caption("Preview images inside the same file next to SKU / SEO / Title / GTIN. Supports http/https & CDN links. Choose a mode: IMAGE() formula, ANCHOR (embed), or AUTO.")

uploaded = st.file_uploader("Upload your Excel masterfile (.xlsx)", type=["xlsx"])

if uploaded:
    data = uploaded.read()
    wb = load_workbook(io.BytesIO(data))
    sheets = wb.sheetnames

    st.sidebar.header("Settings")
    mode = st.sidebar.selectbox("Mode", ["AUTO (try IMAGE, else embed)", "IMAGE (Excel formula)", "ANCHOR (embed pictures)"], index=0)
    header_row = st.sidebar.number_input("Header row number", min_value=1, value=1, step=1)
    width = st.sidebar.number_input("Image width (px)", min_value=40, value=140, step=10)
    height = st.sidebar.number_input("Image height (px)", min_value=40, value=140, step=10)
    keep_notes = st.sidebar.checkbox("Keep original URL as cell note", value=True)
    create_adjacent = st.sidebar.checkbox("Create preview in NEW adjacent column(s)", value=False)

    sheet_mode = st.sidebar.radio("Sheets to process", ["One sheet", "All sheets"], horizontal=True)
    if sheet_mode == "One sheet":
        sheet_name = st.sidebar.selectbox("Which sheet?", sheets, index=0)
        target_sheets = [sheet_name]
    else:
        target_sheets = sheets

    ws0 = wb[target_sheets[0]]
    headers = []
    for c in range(1, ws0.max_column + 1):
        v = ws0.cell(row=header_row, column=c).value
        headers.append(v.strip() if isinstance(v, str) else f"Col {c}")

    auto_cols_idx = detect_url_columns(ws0, header_row=header_row)
    auto_cols_names = [headers[i-1] if i-1 < len(headers) else f"Col {i}" for i in auto_cols_idx]

    selected_by_name = st.multiselect(
        "Pick URL columns by header (auto-detected suggestions pre-selected)",
        options=headers,
        default=auto_cols_names
    )

    def headers_to_indices(ws, names:List[str])->Set[int]:
        return set(columns_by_names(ws, names, header_row=header_row))

    rows = []
    for s in target_sheets:
        ws = wb[s]
        targets = headers_to_indices(ws, selected_by_name) if selected_by_name else set(detect_url_columns(ws, header_row=header_row))
        url_cells = 0
        for cell in iter_target_cells(ws, targets, header_row=header_row):
            v = cell.value
            if isinstance(v, str) and is_url_like(v): url_cells += 1
        rows.append((s, len(targets), url_cells))
    st.dataframe({"Sheet": [r[0] for r in rows], "Target columns": [r[1] for r in rows], "URL cells found": [r[2] for r in rows]}, use_container_width=True)

    if st.button("Process & Prepare Download", type="primary"):
        changed = 0
        for s in target_sheets:
            ws = wb[s]
            targets = headers_to_indices(ws, selected_by_name) if selected_by_name else set(detect_url_columns(ws, header_row=header_row))
            if not targets: continue
            preview_targets = targets
            adjust_dimensions(ws, preview_targets, row_height_px=max(width, height))

            chosen = "auto" if mode.startswith("AUTO") else ("image" if mode.startswith("IMAGE") else "anchor")

            processed = 0
            for cell in iter_target_cells(ws, preview_targets, header_row=header_row):
                val = cell.value
                if not (isinstance(val, str) and is_url_like(val)): continue
                url = normalize_url(val) or val.strip()
                try:
                    if chosen == "image":
                        place_image_formula(cell, url, width, height, keep_note=keep_notes)
                    elif chosen == "anchor":
                        place_anchor_image(ws, cell, url, width, height, keep_note=keep_notes)
                    else:
                        try:
                            place_image_formula(cell, url, width, height, keep_note=keep_notes)
                        except Exception:
                            place_anchor_image(ws, cell, url, width, height, keep_note=keep_notes)
                    processed += 1
                except Exception as e:
                    cell.comment = Comment(f"Preview failed; kept value.\n{url}\nError: {e}", "PreviewBot")

            changed += 1

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        st.success("Done! Download your updated workbook below.")
        st.download_button("⬇️ Download updated masterfile", data=out, file_name="master_with_previews.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload your .xlsx masterfile to begin.")
