# Report Automation

Automates the generation of professional Word inspection reports for Maverick Applied Science. A user fills out a data-entry form, and the system produces a formatted, publication-ready `.docx` report.

## Requirements

- Python 3.x
- Dependencies: `pip install -r requirements.txt`

## Setup

Copy `.env.example` to `.env` and update the paths to match your local directories:

```
TEMPLATE_DOC_PATH          → path to Templates/Template Report.docx
OUTPUT_REPORT_FOLDER_PATH  → path to Output Report/ folder
INPUT_DOC_PATH             → path to your filled-in input form
EXTRACTED_DATA_PATH        → path to Reports Extracted Data/ folder
INPUT_IMAGES_FOLDER_PATH   → path to a folder of photos for Stage 1 (optional)
REPORT_INPUT_FORMS_FOLDER  → path to Report Input Forms/ folder
INPUT_FORM_TEMPLATE_PATH   → path to Templates/Template Report Inputs Form.docx
```

## How to Use

The core workflow has two stages, with an optional pre-stage for bulk image insertion.

### Stage 0 (optional) — Auto-Populate the Picture Column

If you already have a folder of inspection photos, run the helper to produce a fresh copy of the input form with the Picture column pre-filled. Set `INPUT_IMAGES_FOLDER_PATH` to your photo folder, then:

```bash
python populate_input_pictures.py
```

This drops a new file in `Report Input Forms/` named `<folder_name>_<YYYYMMDD_HHMMSS>.docx`, with one image per row (natural-sorted by filename, scaled to 3.6" wide, height proportional). Open that file and continue with Stage 1.

### Stage 1 — Fill Out the Input Form

> **Important:** Never edit `Templates/Template Report Inputs Form.docx` directly. Always make a copy first (or use Stage 0).

If you skipped Stage 0, make a copy of `Templates/Template Report Inputs Form.docx`, give it a descriptive name (e.g., `Acme Corp Boiler Inspection 2026-05.docx`), and save it to `Report Input Forms/`. Fill in the copy:

- **4-column table** — customer name, address, contact, subject, PO number, job ID, inspection dates, drawings used, specifications used, and text for the Introduction, Entrance Meeting, and Conclusions sections.
- **3-column table** — one row per image: description, caption, and an embedded image (already inserted if you ran Stage 0).

Save the completed form to `Report Input Forms/` and set `INPUT_DOC_PATH` to its path.

### Stage 2 — Extract Data

Run the extraction script to parse the form and save the data:

```bash
python wordextraction.py
```

This creates:
- `Reports Extracted Data/report_info.csv` — report metadata
- `Reports Extracted Data/picture_info.csv` — image metadata
- `Reports Extracted Data/report_XXXX/` — extracted image files

Each run assigns the next sequential Report ID.

### Stage 3 — Generate the Report

Run the main script and enter the Report ID when prompted:

```bash
python main.py
```

The output report is saved to `Output Report/Report_XXXX_CustomerName.docx`.

## Project Structure

```
report-automation/
├── main.py                          # Report generation orchestrator
├── utils.py                         # Document manipulation helpers
├── wordextraction.py                # Input form parser
├── populate_input_pictures.py       # (Optional) Pre-fills input form Picture column from a folder
├── requirements.txt
├── .env                             # Local path configuration (not tracked)
├── .env.example                     # Template for .env
│
├── Templates/
│   ├── Template Report.docx         # Master output template
│   ├── Template Report Inputs Form.docx  # Data-entry form template
│   └── Example Report.docx          # Reference example
│
├── Report Input Forms/              # Place completed input forms here
├── Reports Extracted Data/          # CSVs and extracted images (auto-generated)
└── Output Report/                   # Final reports (auto-generated)
```

## How It Works

All document manipulation is done entirely with **python-docx** and direct OOXML construction via **lxml** — no Microsoft Word or COM automation required. Figure captions use `SEQ Figure \* ARABIC` field codes and cross-references use `REF _Ref_Fig_N \h \* Charformat` field codes, both hand-crafted as XML elements. Word auto-updates these fields when the document is opened.

### `wordextraction.py` — Data Extraction Flow

```mermaid
flowchart TD
    A([User fills out a copy of the Input Form]) --> B([Run wordextraction.py])
    B --> C[Get next Report ID from report_info.csv]
    C --> D[Create report_XXXX/ folder in Extracted Data]
    D --> E[Open Input Form .docx]

    E --> F[Parse 4-column text table]
    E --> G[Parse 3-column picture table]

    F --> H["Extract: Customer · Address · Contact · Subject<br/>PO · Job ID · Dates · Drawings · Specifications<br/>Introduction · Entrance Meeting · Conclusions"]
    H --> I[(Append row to report_info.csv)]

    G --> J[Extract Description & Caption per row]
    G --> K[Pull embedded image from Word XML]
    K --> L[Save as image_00N.jpeg in report_XXXX/]
    J --> M[(Append row to picture_info.csv with filename)]
    L --> M
```

### `main.py` — Report Generation Flow

```mermaid
flowchart TD
    A([Run main.py]) --> B[User enters Report ID]
    B --> C[Read report_info.csv]
    B --> D[Read picture_info.csv]
    C & D --> E[Copy Template Report.docx · open as Document object]

    E --> F["Update document properties<br/>via lxml custom XML"]
    F --> G["Insert text sections<br/>Introduction · Entrance Meeting · Conclusions"]
    G --> H["Build bullet lists<br/>Drawings Used · Specifications Used"]

    H --> I{For each image pair}

    subgraph loop["  Image Processing Loop  "]
        J["Create 1- or 2-column image table"]
        K["Insert SEQ caption field XML inside each cell<br/>e.g. Figure 1: description — returns bookmark name"]
        L["Insert placeholder bullets above table"]
        M["Rewrite bullets as REF field XML<br/>e.g. Figure 1 shows description"]
        J --> K --> L --> M
    end

    I --> J
    M --> I
    I -->|All images done| O["Clean up empty paragraphs"]
    O --> P([doc.save → Report_XXXX_CustomerName.docx])
```
