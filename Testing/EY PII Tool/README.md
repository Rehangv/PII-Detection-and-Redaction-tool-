# PII Detection and Redaction Tool

A Streamlit-based web application that scans uploaded documents for personally identifiable information (PII), displays the results in an interactive table, and lets you download redacted copies of the original files.

Built as part of an EY internship project.

---

## What it does

- Scans **Excel (.xlsx)**, **Word (.docx)**, and **PDF (.pdf)** files
- Detects both **US** and **Indian** PII types
- Groups related PII into **pairs** (e.g. Name + Email, Name + DOB)
- Runs **OCR** on embedded images (driver licenses, scanned pages) using Tesseract
- Lets you select which PII types to redact and downloads clean copies

---

## PII types detected

**US / International**
- SSN, US Passport, Driver's License (all 50 states)
- Phone, Email, Address, Date of Birth
- Credit/Debit Card (Visa, Mastercard, AmEx)
- IP Address, MRN (Medical Record Number), Insurance ID
- Bank Account Number, Routing Number
- Gender, Age, State, City

**Indian**
- Aadhaar Number, PAN Card, IFSC Code
- Indian Phone (+91), Voter ID

**Pairs (linked detections)**
- Name + Email
- Name + Date of Birth
- Name + Company ID

---

## Project structure

```
your-folder/
│
├── app.py                      # Streamlit UI — file upload, scan, results display, redaction
├── pi_detection_functions.py   # All detection logic, patterns, OCR, redaction functions
├── requirements.txt            # Python dependencies
├── logo.png                    # App logo shown in the sidebar
├── req.txt                     # (optional) alternate requirements reference
└── .streamlit/                 # Streamlit config (theme, page settings)
```

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name
```

### 2. Create a virtual environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Mac / Linux
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Install spaCy language model

The tool uses spaCy for person name extraction.

```bash
python -m spacy download en_core_web_sm
```

### 5. Install Tesseract OCR

Required only for scanning images embedded in documents.

- **Windows** — download from [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)
- **Mac** — `brew install tesseract`
- **Linux** — `sudo apt install tesseract-ocr`

If you're on Windows, update the path in `pi_detection_functions.py` (line ~36):

```python
_win_tesseract = r"C:\path\to\tesseract.exe"
```

### 6. Run the app

```bash
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

---

## How to use

1. **Upload files** — drag and drop one or more `.xlsx`, `.docx`, or `.pdf` files
2. **Click Scan** — the tool extracts text, runs detection, and shows a results table
3. **Review results** — each row shows the file name, page/sheet, PII type, detected value, and confidence score
4. **Select PII types** — use the dropdown to choose which types to redact
5. **Redact and download** — the tool creates redacted copies where selected PII is replaced with `[REDACTED]`

---

## How detection works

**Text extraction**
Each file type has its own scanner. The Excel scanner reads rows and column headers. The DOCX scanner processes paragraphs and tables. The PDF scanner extracts text page by page and normalises common formatting artifacts (like split words across lines).

**Pair detection**
When a name appears alongside an email, DOB, or company ID in the same block of text, the tool groups them into a pair like `Name, email: James Carter, jcarter@example.com`. This avoids duplicate rows — if a value is already in a pair, it won't appear again as a standalone entry.

**OCR**
If a page has no extractable text (image-only), the tool renders it and passes it through Tesseract. Pages with sufficient selectable text are skipped to avoid duplication.

**Confidence scoring**
Every detected value gets a confidence score from 0–100. Values below 40 are dropped. The score considers context — for example, a 9-digit number near the word "Passport" scores higher as a passport number than the same number in isolation.

**Redaction**
- Excel — cell values are replaced in-place
- DOCX — text runs are replaced while preserving formatting
- PDF — a white rectangle is drawn over each match with `[REDACTED]` in black text. For values that split across table cell lines (e.g. `253-75-` on one line and `9741` on the next), the tool searches for each fragment separately.

---

## Notes

- All PII detection is local — no data is sent to any external server
- The tool is designed for synthetic/test data validation; always review results before using on real documents
- OCR accuracy depends on image quality; blurry or low-resolution scans may miss some values
- The spaCy model is loaded once and cached for the session

---

## Dependencies

| Library | Purpose |
|---|---|
| streamlit | Web UI |
| pandas, openpyxl | Excel reading and processing |
| python-docx | Word document parsing |
| PyPDF2, pdfplumber | PDF text extraction |
| pymupdf (fitz) | PDF redaction (drawing boxes) |
| pytesseract, Pillow | OCR on embedded images |
| pypdfium2 | Rendering PDF pages as images for OCR |
| spaCy | Person name extraction (NER) |
| python-dateutil | Date parsing and validation |