# Chepi

Chepi is a proof of concept for automating product uploads to [OrderEAT](https://ordereat.com/). It turns unstructured product information—such as WhatsApp messages, menus, PDFs, spreadsheets, CSV files, or plain text—into a normalized Excel file ready for OrderEAT's bulk-upload workflow.

> This repository is an experimental prototype, not a production service.

## How it works

1. A user uploads one or more files or pastes product information into the web interface.
2. Chepi extracts the available text, using OCR as a fallback for scanned PDFs.
3. Claude structures and normalizes each product's title, description, category, price, and availability dates.
4. The user reviews a preview and downloads the generated `.xlsx` file.

## Supported inputs

- PDF documents, including scanned documents through OCR
- Excel workbooks (`.xlsx`)
- CSV and TSV files
- Plain-text files
- Text pasted directly into the form

## Requirements

- Python 3.10+
- An [Anthropic API key](https://console.anthropic.com/)
- [Poppler](https://poppler.freedesktop.org/) for rendering scanned PDFs
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract), including Spanish language data, for OCR

On macOS, the system dependencies can be installed with Homebrew:

```bash
brew install poppler tesseract tesseract-lang
```

On Debian or Ubuntu:

```bash
sudo apt-get update
sudo apt-get install poppler-utils tesseract-ocr tesseract-ocr-spa
```

## Local setup

Clone the repository and enter its directory:

```bash
git clone https://github.com/ramiro-mart/chepi.git
cd chepi
```

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install the Python dependencies:

```bash
pip install -r requirements.txt
```

Create the local environment file and add your Anthropic API key:

```bash
cp .env.example .env
```

```dotenv
ANTHROPIC_API_KEY=your-api-key-here
```

Start the application:

```bash
uvicorn app:app --reload
```

Open [http://localhost:8000](http://localhost:8000) in your browser.

## Output format

Chepi generates an Excel workbook with the following columns:

| Column | Description |
| --- | --- |
| `TÍTULO` | Normalized product name |
| `DESCRIPCIÓN` | Product details, or `-` when no details are available |
| `CATEGORÍA` | Source category or an inferred OrderEAT category |
| `PRECIO` | Numeric price, or `REVISAR` when the source has no price |
| `FECHAS DISPONIBLES` | Availability date in `DD/MM/YYYY` format when applicable |

## Prototype limitations

- Uploaded and generated files are stored in the operating system's temporary directory and are not persisted.
- Input is truncated after 80,000 characters.
- The application has no authentication, rate limiting, background jobs, or production deployment configuration.
- AI-generated product data can be incomplete or incorrect and should be reviewed before upload.
- API and model availability may change; the configured Anthropic model may need to be updated over time.

## Tech stack

- FastAPI and Uvicorn
- Anthropic Python SDK
- OpenPyXL
- pdfplumber, pdf2image, and Tesseract OCR
