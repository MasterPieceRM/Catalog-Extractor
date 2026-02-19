# PDF Catalog Product Extractor

A comprehensive solution for extracting product information from PDF catalogs using layout analysis, OCR, and LLM-powered structuring. Features a Streamlit QC/Review UI that guarantees 100% coverage (no missing blocks).

## 🌟 Features

- **Layout/OCR-based Extraction**: Uses PyMuPDF for native PDF text extraction and EasyOCR/Tesseract for scanned pages
- **Schema-driven LLM Structuring**: Pydantic models for consistent product data structure
- **Coverage Tracking**: Block-level tracking ensures no content is missed
- **QC/Review UI**: Streamlit interface for reviewing and correcting extracted products
- **Batch Processing**: Process entire catalogs with parallel processing support
- **Export Options**: JSON and CSV export with customizable fields

## 📁 Project Structure

```
RAG_project/
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── .env                      # Environment configuration
├── src/
│   ├── __init__.py
│   ├── config.py             # Configuration settings
│   ├── pdf_processor.py      # PDF loading and page extraction
│   ├── ocr_extractor.py      # OCR and layout analysis
│   ├── schemas.py            # Pydantic product schemas
│   ├── llm_extractor.py      # LLM-based product extraction
│   ├── coverage_tracker.py   # Block coverage tracking
│   ├── batch_processor.py    # Batch processing module
│   └── ui_components.py      # Additional UI components
├── data/
│   ├── uploads/              # Uploaded PDF files
│   ├── outputs/              # Exported results
│   └── cache/                # Processing cache
└── offers/                   # Sample catalogs
```

## 🚀 Quick Start

### 1. Install Dependencies

```bash
# Create virtual environment (optional but recommended)
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install requirements
pip install -r requirements.txt

# For Tesseract OCR (optional)
# Ubuntu/Debian: sudo apt-get install tesseract-ocr
# macOS: brew install tesseract
```

### 2. Configure Environment

Edit `.env` file:

```env
API_KEY=your_openrouter_api_key
LLM_MODEL=openai/gpt-4o-mini
LLM_BASE_URL=https://openrouter.ai/api/v1
OCR_ENGINE=easyocr
```

### 3. Run the Application

```bash
streamlit run app.py
```

Open http://localhost:8501 in your browser.

## 📖 Usage Guide

### Extraction Workflow

1. **Upload PDF**: Use the sidebar to upload a product catalog PDF
2. **Load Document**: Click "Load/Reload PDF" to process the document
3. **View Pages**: Navigate through pages and see extracted text blocks
4. **Extract Products**: Click "Extract Products from This Page" to run LLM extraction
5. **Review Products**: Switch to Review mode to validate extracted products
6. **Export**: Export reviewed products to JSON or CSV

### Coverage Tracking

The system tracks every text block in the PDF:

- 🔵 **Unprocessed**: Not yet processed
- 🟢 **Extracted**: Successfully extracted to a product
- ✅ **Reviewed**: Manually reviewed and confirmed
- ⚫ **Ignored**: Intentionally ignored (headers, footers, etc.)
- 🔴 **Failed**: Extraction attempted but failed

The UI shows coverage percentage and flags any unprocessed blocks.

### Batch Processing (CLI)

For processing entire catalogs without the UI:

```bash
python -m src.batch_processor path/to/catalog.pdf --format json --ocr
```

Options:
- `--format`: Output format (json or csv)
- `--ocr`: Enable OCR for scanned pages
- `--parallel`: Enable parallel processing

## 🔧 Configuration

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `API_KEY` | OpenRouter/OpenAI API key | Required |
| `LLM_MODEL` | LLM model to use | `openai/gpt-4o-mini` |
| `LLM_BASE_URL` | LLM API base URL | `https://openrouter.ai/api/v1` |
| `OCR_ENGINE` | OCR engine (tesseract/easyocr/paddleocr) | `easyocr` |

### Processing Settings (src/config.py)

- `DPI`: Resolution for PDF rendering (default: 300)
- `MIN_BLOCK_CONFIDENCE`: Minimum OCR confidence (default: 0.5)
- `MAX_TOKENS_PER_CHUNK`: Token limit for LLM calls (default: 4000)

## 📊 Product Schema

Extracted products include:

```python
Product(
    product_id="unique_id",
    sku="SKU123",
    name="Product Name",
    brand="Brand",
    category="Category",
    description="Full description",
    features=["Feature 1", "Feature 2"],
    specifications={"Material": "Steel", "Weight": "1kg"},
    price=Price(amount=99.99, currency="EUR"),
    dimensions=Dimensions(length=10, width=5, height=2, unit="cm"),
    page_number=1,
    extraction_confidence=0.85,
    needs_review=True
)
```

## 🔍 How It Works

### 1. PDF Processing
- Load PDF with PyMuPDF
- Extract native text blocks with bounding boxes
- Render pages to images for OCR if needed

### 2. Layout Analysis
- Detect column structure
- Group nearby blocks into product regions
- Merge PDF and OCR results

### 3. LLM Extraction
- Chunk text to fit token limits
- Send to LLM with extraction prompt
- Parse JSON response into Product objects

### 4. Coverage Tracking
- Track every block's status
- Ensure no content is missed
- Flag blocks needing attention

### 5. QC/Review
- Display products for validation
- Allow manual corrections
- Mark reviewed products as complete

## 🛠️ Extending

### Custom Product Schema

Edit `src/schemas.py` to add fields:

```python
class Product(BaseModel):
    # Add your custom fields
    custom_field: Optional[str] = None
```

### Custom Auto-Ignore Rules

Edit `src/coverage_tracker.py` `AutoIgnoreRules` class:

```python
IGNORE_PATTERNS = [
    r'^your_pattern_here',
]
```

### Custom OCR Engine

Implement the extraction method in `src/ocr_extractor.py`:

```python
def _extract_custom(self, image, page_num):
    # Your custom OCR logic
    pass
```

## 📝 License

MIT License

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request
