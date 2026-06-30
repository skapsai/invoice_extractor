<<<<<<< HEAD
# Invoice Extractor

A local, offline Python tool that extracts key invoice data from PDF files
(text-based or scanned) and writes the results to a formatted Excel spreadsheet.


---

## Table of Contents

1. [Folder Structure](#folder-structure)
2. [Prerequisites](#prerequisites)
3. [Install Python Packages](#install-python-packages)
4. [Install Tesseract OCR](#install-tesseract-ocr)
5. [Install Poppler](#install-poppler)
6. [Configure the Script](#configure-the-script)
7. [Place PDFs and Run](#place-pdfs-and-run)
8. [Excel Output](#excel-output)
9. [Extending for New Vendors](#extending-for-new-vendors)

---

## Folder Structure

```
project_folder/
│
├── invoice_extractor.py     ← Main script
├── requirements.txt         ← Python dependencies
├── README.md                ← This file
│
├── invoices/                ← ✅ Place your PDF invoices here
├── extracted_text/          ← Raw text / OCR output saved here (auto-created)
├── output/                  ← Excel output saved here (auto-created)
├── processed/               ← Successfully processed PDFs moved here (auto-created)
└── error/                   ← Failed PDFs moved here (auto-created)
```

---

## Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.9 or higher |
| Tesseract OCR | 5.x (Windows installer) |
| Poppler | Latest (Windows binary) |

---

## Install Python Packages

Open a terminal in the project folder and run:

```bash
pip install -r requirements.txt
```

This installs:
- `pdfplumber` – text extraction from text-based PDFs
- `pytesseract` – Python wrapper for Tesseract OCR
- `pdf2image` – converts PDF pages to images for OCR
- `pillow` – image handling (required by pdf2image)
- `pandas` – data manipulation and DataFrame handling
- `openpyxl` – Excel file creation and formatting

---

## Install Tesseract OCR

Tesseract is required for OCR on scanned / image-only PDFs.

### Windows

1. Download the installer from:
   **https://github.com/UB-Mannheim/tesseract/wiki**
2. Run the installer and note the installation path
   (default: `C:\Program Files\Tesseract-OCR\tesseract.exe`)
3. The script is pre-configured for this default path.
   If you installed it elsewhere, update this line in `invoice_extractor.py`:

```python
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### Linux / macOS

```bash
# Ubuntu / Debian
sudo apt install tesseract-ocr

# macOS (Homebrew)
brew install tesseract
```

Then remove or comment out the `tesseract_cmd` line in the script.

---

## Install Poppler

Poppler is required by `pdf2image` to convert PDF pages to images.

### Windows

1. Download the latest Windows binary from:
   **https://github.com/oschwartz10612/poppler-windows/releases**
2. Extract the ZIP, e.g., to `C:\poppler-24.07.0\`
3. **Option A – Add to PATH (recommended)**:
   - Add `C:\poppler-24.07.0\Library\bin` to your Windows `PATH` environment variable.
   - Restart your terminal.
4. **Option B – Set path in script**:
   - Edit this line in `invoice_extractor.py`:
   ```python
   POPPLER_PATH: Optional[str] = r"C:\poppler-24.07.0\Library\bin"
   ```

### Linux / macOS

```bash
# Ubuntu / Debian
sudo apt install poppler-utils

# macOS (Homebrew)
brew install poppler
```

No code change needed; Poppler will be on the system PATH automatically.

---

## Configure the Script

Open `invoice_extractor.py` and review the **CONFIGURATION** section at the top:

```python
# Tesseract path (Windows)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Poppler path (set to None if Poppler is already on PATH)
POPPLER_PATH: Optional[str] = None

# OCR resolution – higher DPI = better quality but slower
OCR_DPI = 300

# Minimum characters to consider a PDF text-based (vs. scanned)
MIN_TEXT_LENGTH = 100
```

---

## Place PDFs and Run

1. Copy all your PDF invoice files into the `invoices/` folder.
2. Open a terminal in the project folder.
3. Run:

```bash
python invoice_extractor.py
```

The script will:
- Process each PDF in `invoices/`
- Use `pdfplumber` for text-based PDFs
- Automatically fall back to Tesseract OCR for scanned PDFs
- Save raw extracted text to `extracted_text/` for debugging
- Move successfully processed PDFs to `processed/`
- Move failed PDFs to `error/`
- Print a summary in the terminal when finished

---

## Excel Output

The Excel file is saved at:

```
output/extracted_invoice_records.xlsx
```

### Columns

| Column | Description |
|---|---|
| Source File | Original PDF filename |
| Invoice Date | Date from invoice header |
| Invoice Number | Invoice / Bill / Tax Invoice number |
| PO Number | Purchase Order / Buyer Order number |
| Item | Description of goods or service |
| Serial Number | Serial / Batch / IMEI number |
| Quantity | Number of units |
| Unit | Unit of measure (e.g., NOS, PCS) |
| Basic Price excluding GST | Taxable / assessable value |
| Extraction Status | `Success`, `Partial`, or `Failed` |
| Remarks | Notes on missing fields, OCR method, serial number source, etc. |

### Formatting

- Bold, dark-blue header row with white text
- Alternating row shading for readability
- Frozen top row for easy scrolling
- Auto-adjusted column widths (capped at 50 characters)
- Text wrap enabled on all cells

---

## Extending for New Vendors

The script is designed to be modular. To add vendor-specific extraction rules:

### 1. Add new regex patterns

Open `invoice_extractor.py` and add patterns near the top under the
**REGEX PATTERNS** section:

```python
# Example: Vendor XYZ uses "Ref#" for invoice number
VENDOR_XYZ_INV_PATTERN = re.compile(r"Ref#\s*([A-Z0-9\-]+)", re.IGNORECASE)
```

### 2. Create a vendor-specific extractor function

```python
def extract_vendor_xyz(text: str) -> dict:
    """Vendor XYZ specific header extraction."""
    return {
        "Invoice Date": regex_first(INVOICE_DATE_PATTERN, text),
        "Invoice Number": regex_first(VENDOR_XYZ_INV_PATTERN, text),
        "PO Number": regex_first(PO_NUMBER_PATTERN, text),
    }
```

### 3. Hook into `process_pdf()`

In the `process_pdf()` function, add a vendor detection block before the
generic `extract_invoice_header()` call:

```python
if "Vendor XYZ" in text or "VENDOR XYZ GSTIN" in text:
    header = extract_vendor_xyz(text)
else:
    header = extract_invoice_header(text)
```

### 4. Adjust table column keywords if needed

Update the `TABLE_HEADERS` dict if the vendor uses non-standard column headings.

---

## Troubleshooting

| Issue | Fix |
|---|---|
| `TesseractNotFoundError` | Set `tesseract_cmd` to the correct Tesseract path |
| `PDFInfoNotInstalledError` | Install Poppler and add it to PATH |
| Empty columns in Excel | Check `extracted_text/` for raw text; refine regex patterns |
| Wrong data extracted | Add vendor-specific patterns (see Extending section above) |
| PDF moved to `error/` folder | Check terminal output for the error message |
=======
# invoice_extractor
>>>>>>> 597729f79233db3c0e30cda316390e4992ca72e4
