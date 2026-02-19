"""
PDF Processing Module
Handles PDF loading, page rendering, and basic text extraction
"""
import fitz  # PyMuPDF
from PIL import Image
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional, Tuple
import io
import hashlib

from src.config import DPI, CACHE_DIR


@dataclass
class PDFPage:
    """Represents a single PDF page with its content"""
    page_num: int
    width: float
    height: float
    image: Optional[Image.Image] = None
    text_content: str = ""
    blocks: List[dict] = field(default_factory=list)

    @property
    def dimensions(self) -> Tuple[float, float]:
        return (self.width, self.height)


@dataclass
class PDFDocument:
    """Represents a complete PDF document"""
    file_path: Path
    file_hash: str
    total_pages: int
    pages: List[PDFPage] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)

    @property
    def is_loaded(self) -> bool:
        return len(self.pages) == self.total_pages


class PDFProcessor:
    """Handles PDF loading and processing"""

    def __init__(self, dpi: int = DPI):
        self.dpi = dpi
        self.zoom = dpi / 72  # 72 is the base PDF resolution

    def get_file_hash(self, file_path: Path) -> str:
        """Generate MD5 hash of PDF file for caching"""
        with open(file_path, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()

    def load_document(self, file_path: Path) -> PDFDocument:
        """Load a PDF document and extract basic info"""
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"PDF file not found: {file_path}")

        file_hash = self.get_file_hash(file_path)

        with fitz.open(file_path) as doc:
            metadata = dict(doc.metadata) if doc.metadata else {}
            total_pages = len(doc)

            pdf_doc = PDFDocument(
                file_path=file_path,
                file_hash=file_hash,
                total_pages=total_pages,
                metadata=metadata
            )

        return pdf_doc

    def extract_page(self, file_path: Path, page_num: int,
                     render_image: bool = True) -> PDFPage:
        """Extract a single page from PDF with optional image rendering"""
        with fitz.open(file_path) as doc:
            if page_num >= len(doc):
                raise ValueError(f"Page {page_num} does not exist in document")

            page = doc[page_num]
            rect = page.rect

            # Extract text blocks with positions
            blocks = []
            text_dict = page.get_text(
                "dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

            for block_idx, block in enumerate(text_dict.get("blocks", [])):
                if block.get("type") == 0:  # Text block
                    block_data = {
                        "block_id": f"p{page_num}_b{block_idx}",
                        "bbox": block.get("bbox", []),
                        "lines": [],
                        "text": "",
                        "extracted": False,  # Track if processed by LLM
                        "confidence": 1.0
                    }

                    full_text = []
                    for line in block.get("lines", []):
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span.get("text", "")
                        block_data["lines"].append({
                            "bbox": line.get("bbox", []),
                            "text": line_text.strip()
                        })
                        full_text.append(line_text)

                    block_data["text"] = "\n".join(full_text).strip()
                    if block_data["text"]:  # Only add non-empty blocks
                        blocks.append(block_data)

                elif block.get("type") == 1:  # Image block
                    blocks.append({
                        "block_id": f"p{page_num}_img{block_idx}",
                        "bbox": block.get("bbox", []),
                        "type": "image",
                        "extracted": False,
                        "needs_ocr": True
                    })

            # Get plain text
            text_content = page.get_text("text")

            # Render page to image if requested
            image = None
            if render_image:
                mat = fitz.Matrix(self.zoom, self.zoom)
                pix = page.get_pixmap(matrix=mat)
                image = Image.open(io.BytesIO(pix.tobytes("png")))

            return PDFPage(
                page_num=page_num,
                width=rect.width,
                height=rect.height,
                image=image,
                text_content=text_content,
                blocks=blocks
            )

    def extract_all_pages(self, file_path: Path,
                          render_images: bool = True) -> PDFDocument:
        """Extract all pages from a PDF document"""
        pdf_doc = self.load_document(file_path)

        for page_num in range(pdf_doc.total_pages):
            page = self.extract_page(file_path, page_num, render_images)
            pdf_doc.pages.append(page)

        return pdf_doc

    def get_page_image(self, file_path: Path, page_num: int) -> Image.Image:
        """Get just the image for a specific page"""
        page = self.extract_page(file_path, page_num, render_image=True)
        return page.image

    def get_all_blocks(self, pdf_doc: PDFDocument) -> List[dict]:
        """Get all text blocks across all pages"""
        all_blocks = []
        for page in pdf_doc.pages:
            for block in page.blocks:
                block_copy = block.copy()
                block_copy["page_num"] = page.page_num
                all_blocks.append(block_copy)
        return all_blocks

    def get_unprocessed_blocks(self, pdf_doc: PDFDocument) -> List[dict]:
        """Get blocks that haven't been processed yet"""
        return [b for b in self.get_all_blocks(pdf_doc) if not b.get("extracted")]

    def mark_blocks_extracted(self, pdf_doc: PDFDocument,
                              block_ids: List[str]) -> None:
        """Mark specific blocks as extracted/processed"""
        for page in pdf_doc.pages:
            for block in page.blocks:
                if block["block_id"] in block_ids:
                    block["extracted"] = True


# Utility functions
def merge_nearby_blocks(blocks: List[dict],
                        vertical_threshold: float = 20) -> List[dict]:
    """Merge blocks that are close together vertically (likely same product)"""
    if not blocks:
        return blocks

    # Sort by vertical position
    sorted_blocks = sorted(blocks, key=lambda b: (b.get("page_num", 0),
                                                  b["bbox"][1] if b.get("bbox") else 0))

    merged = []
    current_group = [sorted_blocks[0]]

    for block in sorted_blocks[1:]:
        prev_block = current_group[-1]

        # Check if same page and close vertically
        same_page = block.get("page_num") == prev_block.get("page_num")
        prev_bottom = prev_block["bbox"][3] if prev_block.get("bbox") else 0
        curr_top = block["bbox"][1] if block.get("bbox") else 0
        close_vertically = abs(curr_top - prev_bottom) < vertical_threshold

        if same_page and close_vertically:
            current_group.append(block)
        else:
            # Merge current group
            merged.append(_merge_block_group(current_group))
            current_group = [block]

    # Don't forget the last group
    if current_group:
        merged.append(_merge_block_group(current_group))

    return merged


def _merge_block_group(blocks: List[dict]) -> dict:
    """Merge a group of blocks into one"""
    if len(blocks) == 1:
        return blocks[0]

    # Combine bounding boxes
    all_bboxes = [b["bbox"] for b in blocks if b.get("bbox")]
    if all_bboxes:
        merged_bbox = [
            min(bb[0] for bb in all_bboxes),  # x0
            min(bb[1] for bb in all_bboxes),  # y0
            max(bb[2] for bb in all_bboxes),  # x1
            max(bb[3] for bb in all_bboxes),  # y1
        ]
    else:
        merged_bbox = []

    # Combine text
    merged_text = "\n".join(b.get("text", "") for b in blocks)

    return {
        "block_id": blocks[0]["block_id"] + "_merged",
        "bbox": merged_bbox,
        "text": merged_text,
        "page_num": blocks[0].get("page_num"),
        "extracted": False,
        "source_blocks": [b["block_id"] for b in blocks]
    }
