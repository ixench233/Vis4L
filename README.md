# Vis4VL

This repository hosts a web-based interactive browser for the Vis4VL survey.

The site is adapted from the ZJUTVIS `MPMSurvey` browser. The original website data lived in `data/`; the current Vis4VL browser data is generated from the Excel workbook in `vis_data/`.

## Data Build

Run the converter whenever `vis_data/*.xlsx` changes:

```powershell
python scripts\build_vis4vl_data.py
```

The script generates:

- `data/content.json`
- `data/categories.json`
- `bibtex/*.bib`
- `thumbs100/*.png`
- `thumbs200/*.png`

The converter imports all reviewed rows in the Excel sheet, extracts paper thumbnails from embedded images, derives paper IDs from BibTeX keys, and marks colored-title papers as `Agent` or `RAG` based on the paper title.

## Local Preview

```powershell
node server.js
```

Then open <http://localhost:8000/>.
