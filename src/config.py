"""
Configuration settings for the PDF Catalog Extractor
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Paths
PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"
CACHE_DIR = DATA_DIR / "cache"

# Create directories if they don't exist
for dir_path in [DATA_DIR, UPLOAD_DIR, OUTPUT_DIR, CACHE_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# API Configuration
API_KEY = os.getenv("API_KEY", "")
LLM_MODEL = os.getenv("LLM_MODEL", "openai/gpt-4o-mini")
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "https://openrouter.ai/api/v1")
LLM_VISION_ENABLED = os.getenv("LLM_VISION_ENABLED", "false").lower() == "true"

# OCR Configuration
# Options: tesseract, easyocr, paddleocr
OCR_ENGINE = os.getenv("OCR_ENGINE", "easyocr")
OCR_LANGUAGES = ["en", "fr", "de"]  # Languages for OCR

# Processing Configuration
DPI = 300  # DPI for PDF to image conversion
MIN_BLOCK_CONFIDENCE = 0.5  # Minimum confidence for text blocks
MAX_TOKENS_PER_CHUNK = 4000  # Max tokens for LLM processing

# UI Configuration
PAGE_TITLE = "PDF Catalog Product Extractor"
PAGE_ICON = "📦"
