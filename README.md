# Streamlit: Excel Image In-Cell Preview

A browser app to convert image-URL cells to Excel `IMAGE()` previews **in the same file**—so you can view Image + SKU + SEO + Title + GTIN side-by-side.

## Deploy on Streamlit Cloud
1. Create a new repo and add:
   - `app.py`
   - `requirements.txt`
   - (optional) `README.md`
2. In Streamlit Cloud, point to `app.py` as the entry file.

## How it works
- Upload your `.xlsx` masterfile.
- Choose sheet(s), pick URL columns (or auto-detect), pick size and options.
- Click **Process & prepare download** → download the updated workbook.
- By default, URLs are replaced by `IMAGE()` previews in those same cells.
- Toggle **“Create preview in NEW adjacent column(s)”** to keep URLs untouched and show previews in new columns.

## Notes
- Works with http/https and CDN-like links (e.g., `cdn.amplifi.pattern.com/...`).
- Uses Excel `IMAGE()` formulas; open the downloaded file in Microsoft 365 Excel or Excel on the web for best results.
