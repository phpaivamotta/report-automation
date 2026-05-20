# Report Automation

Automates the generation of professional Word inspection reports for Maverick Applied Science. A user fills out a data-entry form, and the system produces a formatted, publication-ready `.docx` report.

## Requirements

- Windows with Microsoft Word installed
- Python 3.x
- Dependencies: `pip install -r requirements.txt`

## Setup

Copy `.env.example` to `.env` and update the four paths to match your local directories:

```
TEMPLATE_DOC_PATH          → path to Templates/Template Report.docx
OUTPUT_REPORT_FOLDER_PATH  → path to Output Report/ folder
INPUT_DOC_PATH             → path to your filled-in input form
EXTRACTED_DATA_PATH        → path to Reports Extracted Data/ folder
```

## How to Use

The workflow has two stages.

### Stage 1 — Fill Out the Input Form

> **Important:** Never edit `Templates/Template Report Inputs Form.docx` directly. Always make a copy first.

Make a copy of `Templates/Template Report Inputs Form.docx`, give it a descriptive name (e.g., `Acme Corp Boiler Inspection 2026-05.docx`), and save it to `Report Input Forms/`. Fill in the copy:

- **4-column table** — customer name, address, contact, subject, PO number, job ID, inspection dates, drawings used, specifications used, and text for the Introduction, Entrance Meeting, and Conclusions sections.
- **3-column table** — one row per image: description, caption, and an embedded image.

Save the completed form to `Report Input Forms/` under any name you like.

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

The system uses two Python libraries with complementary roles:

- **python-docx** — structural edits: inserting text sections, bullet lists, and image tables into the document XML.
- **win32com** (Word COM automation) — advanced formatting that python-docx cannot do natively: figure captions with auto-numbering, cross-references (e.g., "**Figure 1** shows [description]"), and bulk paragraph font/spacing formatting.

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
    C & D --> E[Copy Template Report.docx to Output Report/]

    E --> F["win32com: Set document properties<br/>Title · Author · Subject · Keywords"]
    F --> G["python-docx: Insert text sections<br/>Introduction · Entrance Meeting · Conclusions"]
    G --> H["python-docx: Build bullet lists<br/>Drawings Used · Specifications Used"]

    H --> I{For each image pair}

    subgraph loop["  Image Processing Loop  "]
        J["python-docx: Create 1- or 2-column image table<br/>and insert images"]
        K["win32com: Add auto-numbered figure caption<br/>e.g. Figure 1: description"]
        L["python-docx: Insert placeholder bullet above table"]
        M["win32com: Insert cross-reference<br/>e.g. Figure 1 shows description"]
        N["win32com: Format bullet — Calibri 12pt · bold figure ref"]
        J --> K --> L --> M --> N
    end

    I --> J
    N --> I
    I -->|All images done| O[python-docx: Clean up empty paragraphs]
    O --> P([Output: Report_XXXX_CustomerName.docx])
```
