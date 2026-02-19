"""
OCR and Layout Extraction Module
Handles OCR for images and scanned PDFs, plus layout analysis
"""
import numpy as np
from PIL import Image
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Any
from enum import Enum
import logging

from src.config import OCR_ENGINE, OCR_LANGUAGES, MIN_BLOCK_CONFIDENCE

logger = logging.getLogger(__name__)


class OCREngine(Enum):
    TESSERACT = "tesseract"
    EASYOCR = "easyocr"
    PADDLEOCR = "paddleocr"


@dataclass
class OCRBlock:
    """Represents an OCR-detected text block"""
    block_id: str
    text: str
    bbox: Tuple[float, float, float, float]  # x0, y0, x1, y1
    confidence: float
    page_num: int = 0
    block_type: str = "text"  # text, title, table, list
    extracted: bool = False

    def to_dict(self) -> dict:
        return {
            "block_id": self.block_id,
            "text": self.text,
            "bbox": list(self.bbox),
            "confidence": self.confidence,
            "page_num": self.page_num,
            "block_type": self.block_type,
            "extracted": self.extracted
        }


@dataclass
class LayoutRegion:
    """Represents a detected layout region (product area, table, etc.)"""
    region_id: str
    bbox: Tuple[float, float, float, float]
    region_type: str  # product, header, footer, sidebar, table
    blocks: List[OCRBlock] = field(default_factory=list)
    confidence: float = 1.0

    def get_combined_text(self) -> str:
        """Get all text from blocks in this region"""
        return "\n".join(b.text for b in self.blocks if b.text)


class OCRProcessor:
    """Handles OCR extraction using multiple engines"""

    def __init__(self, engine: str = OCR_ENGINE, languages: List[str] = None):
        self.engine_name = engine
        self.languages = languages or OCR_LANGUAGES
        self._engine = None
        self._init_engine()

    def _init_engine(self):
        """Initialize the OCR engine"""
        if self.engine_name == "easyocr":
            try:
                import easyocr
                self._engine = easyocr.Reader(self.languages, gpu=False)
                logger.info("EasyOCR engine initialized")
            except ImportError:
                logger.warning(
                    "EasyOCR not available, falling back to tesseract")
                self.engine_name = "tesseract"

        elif self.engine_name == "paddleocr":
            try:
                from paddleocr import PaddleOCR
                self._engine = PaddleOCR(use_angle_cls=True, lang='en')
                logger.info("PaddleOCR engine initialized")
            except ImportError:
                logger.warning(
                    "PaddleOCR not available, falling back to tesseract")
                self.engine_name = "tesseract"

        if self.engine_name == "tesseract":
            try:
                import pytesseract
                self._engine = pytesseract
                logger.info("Tesseract engine initialized")
            except ImportError:
                raise ImportError(
                    "No OCR engine available. Install pytesseract, easyocr, or paddleocr")

    def extract_text(self, image: Image.Image, page_num: int = 0) -> List[OCRBlock]:
        """Extract text blocks from an image using configured OCR engine"""
        if self.engine_name == "easyocr":
            return self._extract_easyocr(image, page_num)
        elif self.engine_name == "paddleocr":
            return self._extract_paddleocr(image, page_num)
        else:
            return self._extract_tesseract(image, page_num)

    def _extract_easyocr(self, image: Image.Image, page_num: int) -> List[OCRBlock]:
        """Extract using EasyOCR"""
        import numpy as np

        # Convert PIL Image to numpy array
        img_array = np.array(image)

        # Run OCR
        results = self._engine.readtext(img_array)

        blocks = []
        for idx, (bbox, text, conf) in enumerate(results):
            if conf < MIN_BLOCK_CONFIDENCE:
                continue

            # Convert bbox format from [[x1,y1],[x2,y2],[x3,y3],[x4,y4]] to (x0,y0,x1,y1)
            x_coords = [p[0] for p in bbox]
            y_coords = [p[1] for p in bbox]
            normalized_bbox = (min(x_coords), min(y_coords),
                               max(x_coords), max(y_coords))

            blocks.append(OCRBlock(
                block_id=f"ocr_p{page_num}_b{idx}",
                text=text,
                bbox=normalized_bbox,
                confidence=conf,
                page_num=page_num
            ))

        return blocks

    def _extract_paddleocr(self, image: Image.Image, page_num: int) -> List[OCRBlock]:
        """Extract using PaddleOCR"""
        import numpy as np

        img_array = np.array(image)
        results = self._engine.ocr(img_array, cls=True)

        blocks = []
        if results and results[0]:
            for idx, line in enumerate(results[0]):
                bbox_points, (text, conf) = line

                if conf < MIN_BLOCK_CONFIDENCE:
                    continue

                # Convert bbox format
                x_coords = [p[0] for p in bbox_points]
                y_coords = [p[1] for p in bbox_points]
                normalized_bbox = (min(x_coords), min(
                    y_coords), max(x_coords), max(y_coords))

                blocks.append(OCRBlock(
                    block_id=f"ocr_p{page_num}_b{idx}",
                    text=text,
                    bbox=normalized_bbox,
                    confidence=conf,
                    page_num=page_num
                ))

        return blocks

    def _extract_tesseract(self, image: Image.Image, page_num: int) -> List[OCRBlock]:
        """Extract using Tesseract"""
        # Get detailed OCR data
        data = self._engine.image_to_data(
            image, output_type=self._engine.Output.DICT)

        blocks = []
        current_block_num = -1
        current_text = []
        current_bbox = None
        current_conf = []

        for i in range(len(data['text'])):
            block_num = data['block_num'][i]
            text = data['text'][i].strip()
            conf = float(data['conf'][i]) / \
                100.0 if data['conf'][i] != -1 else 0

            if block_num != current_block_num:
                # Save previous block
                if current_text and current_bbox:
                    avg_conf = sum(current_conf) / \
                        len(current_conf) if current_conf else 0
                    if avg_conf >= MIN_BLOCK_CONFIDENCE:
                        blocks.append(OCRBlock(
                            block_id=f"ocr_p{page_num}_b{current_block_num}",
                            text=" ".join(current_text),
                            bbox=current_bbox,
                            confidence=avg_conf,
                            page_num=page_num
                        ))

                # Start new block
                current_block_num = block_num
                current_text = []
                current_bbox = None
                current_conf = []

            if text and conf > 0:
                current_text.append(text)
                current_conf.append(conf)

                x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                new_bbox = (x, y, x + w, y + h)

                if current_bbox is None:
                    current_bbox = new_bbox
                else:
                    current_bbox = (
                        min(current_bbox[0], new_bbox[0]),
                        min(current_bbox[1], new_bbox[1]),
                        max(current_bbox[2], new_bbox[2]),
                        max(current_bbox[3], new_bbox[3])
                    )

        # Don't forget the last block
        if current_text and current_bbox:
            avg_conf = sum(current_conf) / \
                len(current_conf) if current_conf else 0
            if avg_conf >= MIN_BLOCK_CONFIDENCE:
                blocks.append(OCRBlock(
                    block_id=f"ocr_p{page_num}_b{current_block_num}",
                    text=" ".join(current_text),
                    bbox=current_bbox,
                    confidence=avg_conf,
                    page_num=page_num
                ))

        return blocks


class LayoutAnalyzer:
    """Analyzes page layout to identify product regions and structure"""

    def __init__(self):
        self.grid_detection_threshold = 0.1  # 10% tolerance for grid alignment

    def detect_layout_regions(self, blocks: List[OCRBlock],
                              page_width: float,
                              page_height: float) -> List[LayoutRegion]:
        """Detect layout regions (columns, grids, product areas) from blocks"""
        if not blocks:
            return []

        # Analyze column structure
        columns = self._detect_columns(blocks, page_width)

        # Analyze row structure within columns
        regions = []
        for col_idx, column_blocks in enumerate(columns):
            col_regions = self._detect_product_regions(column_blocks, col_idx)
            regions.extend(col_regions)

        return regions

    def _detect_columns(self, blocks: List[OCRBlock],
                        page_width: float) -> List[List[OCRBlock]]:
        """Detect column structure in the page"""
        if not blocks:
            return [[]]

        # Get x-center of each block
        x_centers = [(b.bbox[0] + b.bbox[2]) / 2 for b in blocks]

        # Simple column detection: divide page into sections
        # and group blocks by their x-center

        # Try to detect 1, 2, or 3 column layouts
        best_columns = [blocks]  # Default: single column
        best_score = float('inf')

        for num_cols in [1, 2, 3, 4]:
            col_width = page_width / num_cols
            column_groups = [[] for _ in range(num_cols)]

            for block, x_center in zip(blocks, x_centers):
                col_idx = min(int(x_center / col_width), num_cols - 1)
                column_groups[col_idx].append(block)

            # Score: variance in block distribution across columns
            counts = [len(c) for c in column_groups if c]
            if counts:
                variance = np.var(counts) if len(counts) > 1 else 0
                # Prefer fewer empty columns
                empty_penalty = (num_cols - len(counts)) * 10
                score = variance + empty_penalty

                if score < best_score:
                    best_score = score
                    best_columns = [c for c in column_groups if c]

        return best_columns

    def _detect_product_regions(self, blocks: List[OCRBlock],
                                column_idx: int) -> List[LayoutRegion]:
        """Detect individual product regions within a column"""
        if not blocks:
            return []

        # Sort blocks by vertical position
        sorted_blocks = sorted(blocks, key=lambda b: b.bbox[1])

        # Group blocks that are close together vertically
        regions = []
        current_region_blocks = [sorted_blocks[0]]
        region_idx = 0

        vertical_gap_threshold = 50  # pixels

        for block in sorted_blocks[1:]:
            prev_block = current_region_blocks[-1]
            gap = block.bbox[1] - prev_block.bbox[3]

            if gap > vertical_gap_threshold:
                # Create region from current blocks
                region = self._create_region(current_region_blocks,
                                             f"region_c{column_idx}_r{region_idx}")
                regions.append(region)
                region_idx += 1
                current_region_blocks = [block]
            else:
                current_region_blocks.append(block)

        # Don't forget the last region
        if current_region_blocks:
            region = self._create_region(current_region_blocks,
                                         f"region_c{column_idx}_r{region_idx}")
            regions.append(region)

        return regions

    def _create_region(self, blocks: List[OCRBlock], region_id: str) -> LayoutRegion:
        """Create a layout region from a group of blocks"""
        # Calculate bounding box
        x0 = min(b.bbox[0] for b in blocks)
        y0 = min(b.bbox[1] for b in blocks)
        x1 = max(b.bbox[2] for b in blocks)
        y1 = max(b.bbox[3] for b in blocks)

        # Determine region type based on content
        region_type = self._classify_region(blocks)

        # Calculate average confidence
        avg_conf = sum(b.confidence for b in blocks) / len(blocks)

        return LayoutRegion(
            region_id=region_id,
            bbox=(x0, y0, x1, y1),
            region_type=region_type,
            blocks=blocks,
            confidence=avg_conf
        )

    def _classify_region(self, blocks: List[OCRBlock]) -> str:
        """Classify the type of region based on content"""
        combined_text = " ".join(b.text.lower() for b in blocks)

        # Simple heuristics for region classification
        if any(word in combined_text for word in ['price', '€', '$', '£', 'eur', 'usd']):
            return "product"
        elif any(word in combined_text for word in ['page', 'catalog', 'index', 'contents']):
            return "header"
        elif len(blocks) == 1 and len(blocks[0].text) < 50:
            return "title"
        else:
            return "product"  # Default to product region

    def merge_pdf_and_ocr_blocks(self, pdf_blocks: List[dict],
                                 ocr_blocks: List[OCRBlock],
                                 overlap_threshold: float = 0.5) -> List[dict]:
        """Merge PDF native blocks with OCR blocks, preferring PDF when available"""
        merged = []
        used_ocr_blocks = set()

        for pdf_block in pdf_blocks:
            if pdf_block.get("type") == "image" or pdf_block.get("needs_ocr"):
                # Find overlapping OCR blocks
                overlapping = self._find_overlapping_ocr(pdf_block, ocr_blocks,
                                                         overlap_threshold)
                if overlapping:
                    # Replace image block with OCR text
                    combined_text = "\n".join(b.text for b in overlapping)
                    pdf_block["text"] = combined_text
                    pdf_block["needs_ocr"] = False
                    pdf_block["ocr_source"] = [b.block_id for b in overlapping]
                    used_ocr_blocks.update(b.block_id for b in overlapping)

            merged.append(pdf_block)

        # Add OCR blocks that didn't overlap with PDF blocks
        for ocr_block in ocr_blocks:
            if ocr_block.block_id not in used_ocr_blocks:
                merged.append(ocr_block.to_dict())

        return merged

    def _find_overlapping_ocr(self, pdf_block: dict,
                              ocr_blocks: List[OCRBlock],
                              threshold: float) -> List[OCRBlock]:
        """Find OCR blocks that overlap with a PDF block"""
        if not pdf_block.get("bbox"):
            return []

        pdf_bbox = pdf_block["bbox"]
        overlapping = []

        for ocr_block in ocr_blocks:
            overlap = self._calculate_overlap(pdf_bbox, ocr_block.bbox)
            ocr_area = (ocr_block.bbox[2] - ocr_block.bbox[0]) * \
                (ocr_block.bbox[3] - ocr_block.bbox[1])

            if ocr_area > 0 and overlap / ocr_area >= threshold:
                overlapping.append(ocr_block)

        return overlapping

    def _calculate_overlap(self, bbox1: List[float],
                           bbox2: Tuple[float, float, float, float]) -> float:
        """Calculate overlap area between two bounding boxes"""
        x0 = max(bbox1[0], bbox2[0])
        y0 = max(bbox1[1], bbox2[1])
        x1 = min(bbox1[2], bbox2[2])
        y1 = min(bbox1[3], bbox2[3])

        if x1 <= x0 or y1 <= y0:
            return 0.0

        return (x1 - x0) * (y1 - y0)
