# Deltek Open Plan 8.5 Developer's Guide

This repository contains a high-quality markdown conversion of the Deltek Open Plan 8.5 Developer's Guide using advanced OCR technology.

## 🚀 Quick Start

### Prerequisites

1. Python 3.9 or higher
2. An OpenRouter API key ([Get one here](https://openrouter.ai/keys))

### Setup

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd Deltek-OPP-Docs
   ```

2. Install dependencies:
   ```bash
   pip3 install -r requirements.txt
   ```

3. Configure your API key:
   ```bash
   cp .env.example .env
   # Edit .env and add your OpenRouter API key
   ```

### Running the Conversion

**Test with a small sample (recommended first):**
```bash
python3 mistral_ocr_converter.py --max-pages 10
```

**Full conversion:**
```bash
python3 mistral_ocr_converter.py
```

**Custom options:**
```bash
python3 mistral_ocr_converter.py \
  --max-pages 50 \
  --pages-per-chunk 10 \
  --output CustomFilename.md \
  --output-dir custom_output
```

## 📚 Documentation

The converted markdown documentation will be available in the [`docs_mistral/`](docs_mistral/) directory after running the conversion.

## 🔧 Technology Stack

This conversion uses:
- **PyMuPDF (fitz)** - For PDF text extraction
- **OpenRouter API** - For accessing Claude Sonnet 4
- **Claude Sonnet 4** - For intelligent markdown formatting and cleanup
- **Rich** - For beautiful terminal progress display

## ✨ Features

- ✅ **Reliable Text Extraction** - Uses PyMuPDF for accurate text extraction from PDFs
- ✅ **Clean Output** - Claude automatically removes headers, footers, and page numbers
- ✅ **Table Formatting** - Preserves table structures in markdown format
- ✅ **Code Preservation** - Maintains code examples and technical formatting
- ✅ **GitHub-Compatible** - Outputs GitHub Flavored Markdown
- ✅ **Large Document Support** - Processes documents in chunks of 50 pages
- ✅ **Progress Tracking** - Beautiful terminal UI with progress bars

## 📊 Conversion Details

- **Source Document**: Deltek Open Plan 8.5 Developer's Guide (600 pages)
- **Processing Method**: Text extraction + Claude formatting in 50-page chunks
- **Output Format**: GitHub Flavored Markdown
- **Estimated Cost**: ~$2-3 USD for full conversion (600 pages in 12 chunks)
- **Estimated Time**: 5-10 minutes for complete document

## 🔒 Security

- API keys are stored in `.env` files which are git-ignored
- Never commit your `.env` file to version control
- Use `.env.example` as a template for configuration

## 📝 License

The converted documentation maintains all original copyright and licensing terms from Deltek, Inc.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

---

*This conversion tool was created to make the Deltek Open Plan Developer's Guide more accessible and easier to navigate for developers.*
