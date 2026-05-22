# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env   # then fill in the paths
```

The `.env` file must define:
- `TEMPLATE_DOC_PATH` — path to `Templates/Template Report.docx`
- `OUTPUT_REPORT_FOLDER_PATH` — path to `Output Report/` folder
- `INPUT_DOC_PATH` — path to the filled-in input form in `Report Input Forms/`
- `EXTRACTED_DATA_PATH` — path to `Reports Extracted Data/`
- `INPUT_IMAGES_FOLDER_PATH` — path to a folder of photos to auto-insert into the input form (used by `populate_input_pictures.py`)
- `REPORT_INPUT_FORMS_FOLDER` — path to `Report Input Forms/` (output destination for `populate_input_pictures.py`)
- `INPUT_FORM_TEMPLATE_PATH` — path to `Templates/Template Report Inputs Form.docx`

## Running the Program

**Stage 0 (optional) — auto-populate the input form's picture column from a folder of images:**
```bash
python populate_input_pictures.py
```
Produces a fresh copy of the input-form template in `REPORT_INPUT_FORMS_FOLDER` with the Picture column pre-filled. The user then opens that file, fills in Description/Caption text and the 4-column metadata, and points `INPUT_DOC_PATH` at it.

**Stage 1 — extract data from a filled-in input form:**
```bash
python wordextraction.py
```

**Stage 2 — generate a report from extracted data:**
```bash
python main.py
```
`main.py` prompts interactively for a Report ID (integer).

There are no tests or lint commands configured in this project.

## Architecture

The system is a two-stage pipeline that converts a filled-in Word form into a formatted inspection report, with an optional pre-stage for bulk image insertion.

### Stage 0 (optional): `populate_input_pictures.py`

Standalone helper that pre-populates the Picture column of a fresh input-form copy from a folder of photos, so the user doesn't have to paste images one-by-one before filling out the form.

- Reads `INPUT_IMAGES_FOLDER_PATH`, `REPORT_INPUT_FORMS_FOLDER`, `INPUT_FORM_TEMPLATE_PATH` from `.env`.
- Copies the input-form template to `<REPORT_INPUT_FORMS_FOLDER>/<images_folder_name>_<YYYYMMDD_HHMMSS>.docx`.
- Clears all data rows in the second table (rows 1–end; row 0 header preserved): strips every child of each `<w:tc>` except `<w:tcPr>` and appends a fresh empty `<w:p>`. This wipes dummy text and inline `<w:drawing>` while preserving cell width/margins.
- Iterates folder images (natural-sorted by filename, filtered to `.jpg .jpeg .png .gif .bmp .tif .tiff`) and inserts each into column 2 of a successive row, centered, at `width=Inches(3.6)` with height auto-scaled. Adds new rows past the template's 38 data rows if needed.
- Does **not** import or use `utils.py` — that module targets the output report; the input form has its own structure.

### Stage 1: `wordextraction.py`

Reads `INPUT_DOC_PATH` (a copy of `Templates/Template Report Inputs Form.docx` filled in by the user) and writes to `EXTRACTED_DATA_PATH`:
- Parses a **4-column table** → writes one row to `report_info.csv` with all text metadata (customer, dates, drawings, specs, section text, etc.)
- Parses a **3-column table** → for each row, extracts a description, caption, and embedded image; writes to `picture_info.csv` and saves images as `report_XXXX/image_00N.jpeg`
- Assigns sequential Report IDs via `get_next_report_id()`, which reads the max existing ID from `report_info.csv`

### Stage 2: `main.py` + `utils.py`

Reads the CSVs for a chosen Report ID, copies the template, and populates it entirely in memory using a **single python-docx Document object** opened once and saved once at the end.

The document is built in this order:
1. `update_document_properties` — sets core + custom XML properties
2. `insert_formatted_text_after_header` × 3 — Introduction, Entrance Meeting, Conclusions
3. `add_formatted_bullets` × 2 — Drawings Used, Specifications Used
4. Image loop (one iteration per pair of images):
   - `add_table_with_images` — creates 1- or 2-column table, inserts images
   - `add_caption_paragraph` × 1–2 — writes SEQ field XML with a named bookmark (`_Ref_Fig_N`) directly after the table; returns `(bookmark_name, xml_element)` so the second caption chains off the first
   - `add_bullets_above_tables` — inserts placeholder bullets above the table; returns the list of `Paragraph` objects
   - `build_cross_reference_bullet` × 1–2 — rewrites each placeholder bullet in-place as a REF field (pointing to the caption bookmark) followed by " shows \<description\>"
5. `remove_empty_paragraphs_after_table` / `remove_first_empty_paragraph_above_text` — cleanup

### Key design decisions

**No win32com.** All document manipulation uses python-docx with direct OOXML (`lxml`) element construction. Captions use `SEQ Figure \* ARABIC` field codes; cross-references use `REF _Ref_Fig_N \h \* Charformat` field codes — both hand-crafted as XML. Word auto-updates these fields when the document is opened.

**Single document session.** All `utils.py` functions accept a `Document` object (`doc`) and do not open or save the file internally. Only `main.py` calls `doc.save()`, once, at the very end.

**Figure counter.** `main.py` owns `figure_counter` (starts at 1, increments per image). It is passed explicitly to `add_caption_paragraph` and `build_cross_reference_bullet` to keep captions and cross-references in sync.

**Custom document properties** are accessed via `doc.part.package.iter_parts()` to find `/docProps/custom.xml`, then updated via lxml element traversal.

### `utils.py` function reference

| Function | Purpose |
|---|---|
| `update_document_properties(doc, report_data)` | Core + custom Word properties |
| `insert_formatted_text_after_header(doc, header, text)` | Inserts + formats a paragraph after a named heading |
| `add_formatted_bullets(doc, header, items, is_drawing)` | Replaces template bullet with formatted list; prefixes drawings with `"Equipment Drawing: "` |
| `add_table_with_images(doc, header, counter, cols, img1, img2)` | Creates image table, returns `Table` |
| `add_caption_paragraph(after_elem, caption, fig_idx)` | Inserts SEQ caption paragraph, returns `(bm_name, xml_elem)` |
| `build_cross_reference_bullet(para, desc, bm_name, fig_idx)` | Rewrites bullet as bold REF field + description |
| `add_bullets_above_tables(doc, table, cols)` | Inserts placeholder bullets, returns `[Paragraph, ...]` |
| `remove_empty_paragraphs_after_table(doc)` | Cleans blank paragraphs after image tables |
| `remove_first_empty_paragraph_above_text(doc, text)` | Removes leading blank line above a named paragraph |
