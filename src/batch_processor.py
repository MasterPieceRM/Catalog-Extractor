"""
Batch Processing Module
Handles automated extraction of entire catalogs
"""
import json
import logging
from pathlib import Path
from typing import List, Optional, Callable
from dataclasses import dataclass
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from src.pdf_processor import PDFProcessor, PDFDocument
from src.ocr_extractor import OCRProcessor, LayoutAnalyzer
from src.llm_extractor import LLMExtractor, ProductRefiner
from src.coverage_tracker import CoverageTracker, AutoIgnoreRules, ensure_complete_coverage
from src.schemas import Product, CatalogExtractionResult, ExtractionResult
from src.config import OUTPUT_DIR

logger = logging.getLogger(__name__)


@dataclass
class BatchProgress:
    """Progress tracking for batch processing"""
    total_pages: int = 0
    processed_pages: int = 0
    total_products: int = 0
    current_page: int = 0
    status: str = "idle"
    errors: List[str] = None

    def __post_init__(self):
        if self.errors is None:
            self.errors = []

    @property
    def progress_percentage(self) -> float:
        if self.total_pages == 0:
            return 0.0
        return (self.processed_pages / self.total_pages) * 100


class BatchExtractor:
    """
    Batch processor for extracting products from entire catalogs.
    Ensures complete coverage and handles errors gracefully.
    """

    def __init__(self,
                 use_ocr: bool = False,
                 parallel: bool = False,
                 max_workers: int = 4):
        self.pdf_processor = PDFProcessor()
        self.llm_extractor = LLMExtractor()
        self.product_refiner = ProductRefiner(self.llm_extractor)
        self.coverage_tracker = CoverageTracker()

        self.use_ocr = use_ocr
        self.parallel = parallel
        self.max_workers = max_workers

        if use_ocr:
            self.ocr_processor = OCRProcessor()
            self.layout_analyzer = LayoutAnalyzer()

        self.progress = BatchProgress()
        self.progress_callback: Optional[Callable] = None

    def set_progress_callback(self, callback: Callable[[BatchProgress], None]):
        """Set a callback function for progress updates"""
        self.progress_callback = callback

    def _update_progress(self, **kwargs):
        """Update progress and call callback"""
        for key, value in kwargs.items():
            if hasattr(self.progress, key):
                setattr(self.progress, key, value)

        if self.progress_callback:
            self.progress_callback(self.progress)

    def process_catalog(self, file_path: Path,
                        auto_ignore: bool = True) -> CatalogExtractionResult:
        """
        Process an entire catalog and extract all products.

        Args:
            file_path: Path to the PDF file
            auto_ignore: Whether to auto-ignore non-product blocks

        Returns:
            CatalogExtractionResult with all extracted products
        """
        file_path = Path(file_path)
        logger.info(f"Starting batch extraction of {file_path}")

        self._update_progress(status="loading")

        # Load PDF
        pdf_doc = self.pdf_processor.extract_all_pages(file_path)

        self._update_progress(
            total_pages=pdf_doc.total_pages,
            status="processing"
        )

        # Register all blocks
        for page in pdf_doc.pages:
            self.coverage_tracker.register_blocks(page.blocks, page.page_num)

        # Auto-ignore non-product blocks
        if auto_ignore:
            ignored = AutoIgnoreRules.auto_ignore_blocks(self.coverage_tracker)
            logger.info(f"Auto-ignored {ignored} non-product blocks")

        # Create result container
        result = CatalogExtractionResult(
            file_name=file_path.name,
            file_hash=pdf_doc.file_hash,
            total_pages=pdf_doc.total_pages
        )

        # Process pages
        if self.parallel:
            page_results = self._process_pages_parallel(pdf_doc)
        else:
            page_results = self._process_pages_sequential(pdf_doc)

        result.page_results = page_results

        # Refine and deduplicate products
        all_products = result.get_all_products()
        refined_products = self.product_refiner.merge_duplicate_products(
            all_products)

        # Validate products
        for product in refined_products:
            is_valid, issues = self.product_refiner.validate_product(product)
            if not is_valid:
                product.needs_review = True
                logger.warning(
                    f"Product {product.product_id} has issues: {issues}")

        self._update_progress(
            status="complete",
            total_products=len(refined_products)
        )

        # Final coverage check
        is_complete, missing = ensure_complete_coverage(
            self.coverage_tracker, refined_products
        )

        if not is_complete:
            logger.warning(f"{len(missing)} blocks still unprocessed")
            self.progress.errors.append(f"{len(missing)} blocks not extracted")

        return result

    def _process_pages_sequential(self, pdf_doc: PDFDocument) -> List[ExtractionResult]:
        """Process pages one by one"""
        results = []

        for page in pdf_doc.pages:
            self._update_progress(
                current_page=page.page_num,
                processed_pages=page.page_num
            )

            try:
                result = self._process_single_page(page)
                results.append(result)

                self._update_progress(
                    total_products=self.progress.total_products +
                    len(result.products)
                )

            except Exception as e:
                logger.error(f"Error processing page {page.page_num}: {e}")
                self.progress.errors.append(f"Page {page.page_num}: {str(e)}")
                results.append(ExtractionResult(
                    page_num=page.page_num,
                    errors=[str(e)]
                ))

        self._update_progress(processed_pages=pdf_doc.total_pages)
        return results

    def _process_pages_parallel(self, pdf_doc: PDFDocument) -> List[ExtractionResult]:
        """Process pages in parallel"""
        results = [None] * pdf_doc.total_pages

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_page = {
                executor.submit(self._process_single_page, page): page.page_num
                for page in pdf_doc.pages
            }

            for future in as_completed(future_to_page):
                page_num = future_to_page[future]

                try:
                    result = future.result()
                    results[page_num] = result

                    self._update_progress(
                        processed_pages=self.progress.processed_pages + 1,
                        total_products=self.progress.total_products +
                        len(result.products)
                    )

                except Exception as e:
                    logger.error(f"Error processing page {page_num}: {e}")
                    self.progress.errors.append(f"Page {page_num}: {str(e)}")
                    results[page_num] = ExtractionResult(
                        page_num=page_num,
                        errors=[str(e)]
                    )

        return results

    def _process_single_page(self, page) -> ExtractionResult:
        """Process a single page"""
        blocks = page.blocks

        # Apply OCR if enabled
        if self.use_ocr and page.image:
            ocr_blocks = self.ocr_processor.extract_text(
                page.image, page.page_num)
            blocks = self.layout_analyzer.merge_pdf_and_ocr_blocks(
                blocks,
                ocr_blocks
            )

        # Extract products
        result = self.llm_extractor.extract_from_page(blocks, page.page_num)

        # Update coverage tracker
        for product in result.products:
            self.coverage_tracker.mark_extracted(
                product.source_blocks,
                product.product_id
            )

        # Mark failed blocks
        if result.unmatched_blocks:
            self.coverage_tracker.mark_failed(result.unmatched_blocks)

        return result

    def export_results(self, result: CatalogExtractionResult,
                       output_path: Optional[Path] = None,
                       format: str = "json") -> Path:
        """Export extraction results to file"""
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{result.file_name}_{timestamp}"
            output_path = OUTPUT_DIR / f"{filename}.{format}"

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if format == "json":
            data = {
                "file_name": result.file_name,
                "file_hash": result.file_hash,
                "total_pages": result.total_pages,
                "total_products": result.total_products,
                "products_needing_review": result.products_needing_review,
                "average_coverage": result.average_coverage,
                "products": [p.to_dict() for p in result.get_all_products()],
                "coverage_report": self.coverage_tracker.export_coverage_report()
            }

            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

        elif format == "csv":
            import pandas as pd

            products = result.get_all_products()
            rows = []
            for p in products:
                row = {
                    'product_id': p.product_id,
                    'sku': p.sku,
                    'ean': p.ean,
                    'reference': p.reference,
                    'name': p.name,
                    'brand': p.brand,
                    'category': p.category,
                    'subcategory': p.subcategory,
                    'description': p.description,
                    'price': p.price.amount if p.price else None,
                    'currency': p.price.currency.value if p.price else None,
                    'features': "; ".join(p.features),
                    'page': p.page_number + 1,
                    'confidence': p.extraction_confidence,
                    'needs_review': p.needs_review
                }
                rows.append(row)

            df = pd.DataFrame(rows)
            df.to_csv(output_path, index=False, encoding='utf-8')

        logger.info(f"Results exported to {output_path}")
        return output_path

    def get_coverage_report(self) -> dict:
        """Get the current coverage report"""
        return self.coverage_tracker.export_coverage_report()


def process_catalog_cli(file_path: str,
                        output_format: str = "json",
                        use_ocr: bool = False,
                        parallel: bool = False):
    """
    Command-line interface for batch processing.

    Usage:
        python -m src.batch_processor catalog.pdf --format json --ocr --parallel
    """
    from tqdm import tqdm

    extractor = BatchExtractor(use_ocr=use_ocr, parallel=parallel)

    # Create progress bar
    pbar = None

    def progress_callback(progress: BatchProgress):
        nonlocal pbar
        if pbar is None and progress.total_pages > 0:
            pbar = tqdm(total=progress.total_pages, desc="Processing pages")
        if pbar:
            pbar.n = progress.processed_pages
            pbar.refresh()

    extractor.set_progress_callback(progress_callback)

    # Process
    result = extractor.process_catalog(Path(file_path))

    if pbar:
        pbar.close()

    # Export
    output_path = extractor.export_results(result, format=output_format)

    # Print summary
    print(f"\n{'='*50}")
    print(f"Extraction Complete!")
    print(f"{'='*50}")
    print(f"Total Pages: {result.total_pages}")
    print(f"Total Products: {result.total_products}")
    print(f"Products Needing Review: {result.products_needing_review}")
    print(f"Average Coverage: {result.average_coverage:.1f}%")
    print(f"Output: {output_path}")

    coverage = extractor.get_coverage_report()
    if not coverage['is_complete']:
        print(
            f"\n⚠️ Warning: {coverage['by_status']['unprocessed']} blocks unprocessed")
        print(f"   {coverage['by_status']['failed']} blocks failed")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Batch process PDF catalogs")
    parser.add_argument("file", help="Path to PDF file")
    parser.add_argument("--format", choices=["json", "csv"], default="json",
                        help="Output format")
    parser.add_argument("--ocr", action="store_true", help="Enable OCR")
    parser.add_argument("--parallel", action="store_true",
                        help="Parallel processing")

    args = parser.parse_args()
    process_catalog_cli(args.file, args.format, args.ocr, args.parallel)
