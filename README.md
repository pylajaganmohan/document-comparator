# Document Comparator v24

A Streamlit-based document comparison utility for Excel, CSV, and PDF files.

## v24 Highlights

- Excel / CSV comparison
- Row-order-independent matching
- Source → Target column mapping
- Automatic population of mapped columns in "Columns to Compare"
- Single or composite Match Keys
- Repeating / duplicate Match Key support
- Different Source and Target column names
- Ignore spaces
- Ignore case
- Numeric tolerance
- Date/time normalization
- Timezone-aware datetime comparison
- Configurable default timezone for timezone-less datetime values
- Source → Target value aliases
  - Example: `VisionLink → NEW VISIONLINK`
  - Multiple aliases can be configured per column
- Saved comparison setups
- Load / update / delete saved setups
- Source / Target swap
- Missing / Extra / Changed / Matched result sections
- Match-rate and difference summaries
- Professional QA Excel report
- PDF text comparison
- PDF metadata comparison
- Optional OCR for scanned/image-only PDFs
- Poppler-free PDF rendering through `pypdfium2`
- Tesseract OCR support for scanned PDFs

## Requirements

Python 3.10+ is recommended.

Install dependencies:

```powershell
pip install -r requirements.txt
```

## Run the application

From the folder containing `app.py`:

```powershell
streamlit run app.py
```

The application will open in your browser.

## Excel / CSV comparison workflow

1. Select **Excel / CSV**.
2. Upload Source and Target files.
3. Map Source columns to Target columns.
4. The mapped Source columns are automatically populated in **Columns to Compare**.
5. Select a Match Key.
6. Use a Composite Match Key when one column is not unique.
7. Configure comparison rules if required:
   - Ignore spaces
   - Ignore case
   - Numeric tolerance
   - Date/time normalization
   - Default timezone
8. Add value aliases when an old and new system intentionally use different values.

Example:

```text
old value -> NEW VALUE
Jagan  -> Jagan Mohan
Old Portal -> NEW PORTAL
```

9. Click **Compare Files**.
10. Review:
    - Value Differences
    - Missing in Target
    - Extra in Target
    - Matched Records
11. Download the **Professional QA Excel Report**.

## Date / Time and Timezone Handling

When date/time normalization is enabled:

- Different date formats can be compared.
- Timezone-aware timestamps are normalized to UTC.
- Values without a timezone use the configured **Default timezone**.

Example:

```text
2026-09-01 14:30:00+00:00
=
01/09/2026 20:00 IST
```

The default timezone is:

```text
Asia/Kolkata
```

Change it when your timezone-less target/source values belong to another timezone.

## Value Aliases

Value aliases are intended for intentional old-system/new-system differences.

Example:

```text
Column: Application

VisionLink -> NEW VISIONLINK
FleetView  -> NEW FLEETVIEW
Old Portal -> NEW PORTAL
```

These values are treated as equivalent only for the configured column and configured mapping.

The alias configuration is also included in the Excel QA report.

## Composite Match Keys

Use a composite key when one column can repeat.

Example:

```text
Emp_ID + Record_Type
```

This allows:

```text
101 + A
101 + B
104 + A
104 + B
```

to be treated as separate records.

## Professional QA Excel Report

The downloaded report contains:

- Summary
- Configuration
- Value Aliases
- Column Mapping
- Value Differences
- Missing in Target
- Extra in Target
- Matched Records
- Differences by Column
- Duplicate Match Keys

The report also includes:

- PASS / FAIL highlighting
- Filters
- Freeze panes
- Automatic column sizing
- Comparison timestamp
- Source and Target filenames
- Match rate
- Comparison configuration

## PDF Comparison

The PDF comparator supports:

- Page count comparison
- Page-by-page text comparison
- Ignore spaces
- Ignore case
- Optional PDF metadata comparison
- Optional OCR for scanned/image-only PDFs

### OCR

Normal text PDFs do not require OCR.

For scanned PDFs, install Tesseract OCR separately on Windows. The Python package `pytesseract` is only the Python interface to Tesseract.

The application uses `pypdfium2` for PDF rendering, so **Poppler is not required for the normal OCR path**.

If Tesseract is not installed or is not available on PATH, OCR will not work.

## Test Data

A comprehensive test pair should validate:

- Different Source/Target column names
- Composite keys
- Duplicate keys
- Row-order independence
- Ignore spaces
- Ignore case
- Numeric tolerance
- Date/time format normalization
- Timezone conversion
- Value aliases
- Missing records
- Extra records
- Ignored columns

## Troubleshooting

### Streamlit command not found

Try:

```powershell
python -m streamlit run app.py
```

### Missing Python package

Run:

```powershell
pip install -r requirements.txt
```

### OCR does not work

Install Tesseract OCR on Windows and make sure the Tesseract executable is available on PATH.

### PDF says Poppler is required

Make sure you are running the `app.py` and have installed:

```powershell
pip install pypdfium2
```

v24 uses `pypdfium2` as the preferred PDF rendering path.

## Recommended folder structure

```text
DocumentComparator/
│
├── app.py
├── requirements.txt
└── README.md
```

## Start

```powershell
cd path\to\DocumentComparator
pip install -r requirements.txt
python -m streamlit run app.py
```
