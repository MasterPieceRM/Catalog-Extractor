"""
Image Extraction Module
Handles cropping product images from PDF pages based on bounding boxes
"""
import base64
import io
import logging
from pathlib import Path
from typing import List, Optional, Tuple
from PIL import Image
from dataclasses import dataclass
import fitz  # PyMuPDF

logger = logging.getLogger(__name__)


@dataclass
class EmbeddedImage:
    """An image extracted directly from PDF with exact coordinates"""
    image: Image.Image
    bbox: List[float]  # [x1, y1, x2, y2] in PDF coordinates
    page_num: int
    base64_data: str
    width: int
    height: int


@dataclass
class CroppedImage:
    """Result of cropping an image from a page"""
    image: Image.Image
    bbox: List[float]
    base64_data: str
    width: int
    height: int


class ImageExtractor:
    """Handles extracting and cropping product images from PDF pages"""
    
    def __init__(self, default_padding: int = 5):
        """
        Initialize the image extractor.
        
        Args:
            default_padding: Pixels to add around bounding box
        """
        self.default_padding = default_padding
    
    def crop_image_from_page(
        self, 
        page_image: Image.Image, 
        bbox: List[float],
        padding: int = None,
        normalize_coords: bool = True
    ) -> Optional[CroppedImage]:
        """
        Crop a region from a page image based on bounding box.
        
        Args:
            page_image: PIL Image of the full page
            bbox: Bounding box [x1, y1, x2, y2] - can be normalized (0-1) or pixel coords
            padding: Optional padding around the crop (defaults to default_padding)
            normalize_coords: If True, bbox values are interpreted as percentages (0-1)
        
        Returns:
            CroppedImage with the cropped region and base64 data, or None if invalid
        """
        if not bbox or len(bbox) != 4:
            logger.warning(f"Invalid bbox: {bbox}")
            return None
        
        try:
            padding = padding if padding is not None else self.default_padding
            page_width, page_height = page_image.size
            
            x1, y1, x2, y2 = bbox
            
            # Convert normalized coordinates to pixels if needed
            if normalize_coords and all(0 <= v <= 1 for v in [x1, y1, x2, y2]):
                x1 = int(x1 * page_width)
                y1 = int(y1 * page_height)
                x2 = int(x2 * page_width)
                y2 = int(y2 * page_height)
            else:
                # Assume pixel coordinates
                x1, y1, x2, y2 = int(x1), int(y1), int(x2), int(y2)
            
            # Ensure proper ordering
            if x1 > x2:
                x1, x2 = x2, x1
            if y1 > y2:
                y1, y2 = y2, y1
            
            # Apply padding with bounds checking
            x1 = max(0, x1 - padding)
            y1 = max(0, y1 - padding)
            x2 = min(page_width, x2 + padding)
            y2 = min(page_height, y2 + padding)
            
            # Validate dimensions
            if x2 - x1 < 10 or y2 - y1 < 10:
                logger.warning(f"Cropped region too small: {x2-x1}x{y2-y1}")
                return None
            
            # Crop the image
            cropped = page_image.crop((x1, y1, x2, y2))
            
            # Convert to base64
            base64_data = self.image_to_base64(cropped)
            
            return CroppedImage(
                image=cropped,
                bbox=[x1, y1, x2, y2],
                base64_data=base64_data,
                width=cropped.width,
                height=cropped.height
            )
            
        except Exception as e:
            logger.error(f"Failed to crop image: {e}")
            return None
    
    def extract_embedded_images(
        self,
        pdf_path: Path,
        page_num: int
    ) -> List[EmbeddedImage]:
        """
        Extract all embedded images from a PDF page using fitz.
        
        Args:
            pdf_path: Path to PDF file
            page_num: Page number (0-indexed)
        
        Returns:
            List of EmbeddedImage objects with exact bounding boxes
        """
        embedded_images = []
        
        try:
            with fitz.open(pdf_path) as doc:
                if page_num >= len(doc):
                    logger.warning(f"Page {page_num} does not exist in document")
                    return []
                
                page = doc[page_num]
                page_rect = page.rect
                page_width = page_rect.width
                page_height = page_rect.height
                
                # Get all images on the page
                image_list = page.get_images(full=True)
                
                for img_idx, img_info in enumerate(image_list):
                    try:
                        xref = img_info[0]  # Image reference number
                        
                        # Get image bounding box(es) on the page
                        img_rects = page.get_image_rects(xref)
                        
                        if not img_rects:
                            continue
                        
                        # Use the first rectangle (main placement)
                        rect = img_rects[0]
                        
                        # Extract the actual image data
                        base_image = doc.extract_image(xref)
                        image_bytes = base_image["image"]
                        
                        # Convert to PIL Image
                        pil_image = Image.open(io.BytesIO(image_bytes))
                        
                        # Convert to RGB if needed
                        if pil_image.mode in ('RGBA', 'P', 'LA'):
                            pil_image = pil_image.convert('RGB')
                        
                        # Get bounding box as normalized coordinates (0-1)
                        bbox = [
                            rect.x0 / page_width,
                            rect.y0 / page_height,
                            rect.x1 / page_width,
                            rect.y1 / page_height
                        ]
                        
                        # Convert to base64
                        base64_data = self.image_to_base64(pil_image)
                        
                        embedded_images.append(EmbeddedImage(
                            image=pil_image,
                            bbox=bbox,
                            page_num=page_num,
                            base64_data=base64_data,
                            width=pil_image.width,
                            height=pil_image.height
                        ))
                        
                        logger.debug(f"Extracted embedded image {img_idx} at bbox {bbox}")
                        
                    except Exception as e:
                        logger.warning(f"Failed to extract image {img_idx}: {e}")
                        continue
                
                logger.info(f"Extracted {len(embedded_images)} embedded images from page {page_num}")
                
        except Exception as e:
            logger.error(f"Failed to extract embedded images from PDF: {e}")
        
        return embedded_images
    
    @staticmethod
    def calculate_iou(bbox1: List[float], bbox2: List[float]) -> float:
        """
        Calculate Intersection over Union (IoU) between two bounding boxes.
        
        Args:
            bbox1, bbox2: Bounding boxes as [x1, y1, x2, y2] (normalized 0-1)
        
        Returns:
            IoU value between 0 and 1
        """
        x1 = max(bbox1[0], bbox2[0])
        y1 = max(bbox1[1], bbox2[1])
        x2 = min(bbox1[2], bbox2[2])
        y2 = min(bbox1[3], bbox2[3])
        
        # Check if there's an intersection
        if x2 <= x1 or y2 <= y1:
            return 0.0
        
        intersection = (x2 - x1) * (y2 - y1)
        
        area1 = (bbox1[2] - bbox1[0]) * (bbox1[3] - bbox1[1])
        area2 = (bbox2[2] - bbox2[0]) * (bbox2[3] - bbox2[1])
        
        union = area1 + area2 - intersection
        
        if union <= 0:
            return 0.0
        
        return intersection / union
    
    @staticmethod
    def calculate_centroid_distance(bbox1: List[float], bbox2: List[float]) -> float:
        """
        Calculate the Euclidean distance between centroids of two bounding boxes.
        
        Args:
            bbox1, bbox2: Bounding boxes as [x1, y1, x2, y2] (normalized 0-1)
        
        Returns:
            Distance between centroids (lower is better)
        """
        # Calculate centroids
        cx1 = (bbox1[0] + bbox1[2]) / 2
        cy1 = (bbox1[1] + bbox1[3]) / 2
        cx2 = (bbox2[0] + bbox2[2]) / 2
        cy2 = (bbox2[1] + bbox2[3]) / 2
        
        # Euclidean distance
        return ((cx1 - cx2) ** 2 + (cy1 - cy2) ** 2) ** 0.5
    
    def find_best_matching_image(
        self,
        llm_bbox: List[float],
        embedded_images: List[EmbeddedImage],
        min_iou_threshold: float = 0.3,
        max_centroid_distance: float = 0.25  # Max distance to consider for centroid fallback
    ) -> Optional[EmbeddedImage]:
        """
        Find the embedded image that best matches the LLM's bounding box.
        
        Uses IoU (Intersection over Union) as primary matching method.
        Falls back to centroid distance if no IoU match found.
        
        Args:
            llm_bbox: Bounding box from LLM [x1, y1, x2, y2] (normalized 0-1)
            embedded_images: List of extracted embedded images
            min_iou_threshold: Minimum IoU to consider a match (primary method)
            max_centroid_distance: Maximum centroid distance for fallback matching
        
        Returns:
            Best matching EmbeddedImage or None if no good match
        """
        if not embedded_images or not llm_bbox or len(llm_bbox) != 4:
            return None
        
        best_match = None
        best_iou = min_iou_threshold
        
        # Primary matching: IoU-based
        for emb_img in embedded_images:
            iou = self.calculate_iou(llm_bbox, emb_img.bbox)
            if iou > best_iou:
                best_iou = iou
                best_match = emb_img
        
        if best_match:
            logger.debug(f"Found matching image with IoU={best_iou:.3f}")
            return best_match
        
        # Fallback: Centroid distance-based matching
        # This helps when LLM bbox is slightly off but still in the right area
        best_distance = max_centroid_distance
        closest_match = None
        
        for emb_img in embedded_images:
            distance = self.calculate_centroid_distance(llm_bbox, emb_img.bbox)
            if distance < best_distance:
                best_distance = distance
                closest_match = emb_img
        
        if closest_match:
            logger.info(f"Using centroid fallback match (distance={best_distance:.3f})")
            return closest_match
        
        logger.debug(f"No matching image found for bbox {llm_bbox}")
        return None
    
    def extract_product_images(
        self, 
        page_image: Image.Image, 
        products: List,
        page_num: int,
        pdf_path: Path = None
    ) -> List:
        """
        Extract images for all products on a page.
        
        Uses fitz to extract embedded images from the PDF when pdf_path is provided,
        then matches them to LLM bounding boxes using IoU for accurate cropping.
        Falls back to page-crop method if pdf_path not provided or no match found.
        
        Args:
            page_image: PIL Image of the page (used as fallback)
            products: List of Product objects with image_bbox in raw_attributes
            page_num: Current page number
            pdf_path: Optional path to PDF file for embedded image extraction
        
        Returns:
            The same products with images field populated
        """
        from src.schemas import ProductImage
        
        # Extract embedded images from PDF if path provided
        embedded_images = []
        if pdf_path:
            embedded_images = self.extract_embedded_images(pdf_path, page_num)
            if embedded_images:
                logger.info(f"Using {len(embedded_images)} embedded images for matching")
        
        for product in products:
            # Check for image bbox in raw_attributes
            image_bbox = product.raw_attributes.get('image_bbox')
            
            if not image_bbox:
                # No image_bbox from LLM means this product has no image - skip it
                # Only try to populate if there's an existing img with bbox but missing data
                if product.images:
                    for img in product.images:
                        if img.bbox and not img.image_data:
                            cropped = self.crop_image_from_page(page_image, img.bbox)
                            if cropped:
                                img.image_data = cropped.base64_data
                continue  # Always skip to next product if no image_bbox from LLM
            
            # Try to match with embedded image first (fitz method - perfect cropping)
            matched_image = None
            if embedded_images:
                matched_image = self.find_best_matching_image(image_bbox, embedded_images)
            
            if matched_image:
                # Use the perfectly cropped embedded image
                product_image = ProductImage(
                    image_id=f"img_{product.product_id}_{page_num}",
                    page_num=page_num,
                    bbox=matched_image.bbox,
                    image_data=matched_image.base64_data,
                    description=f"Product image for {product.name[:50]}"
                )
                
                if not product.images:
                    product.images = []
                product.images.append(product_image)
                
                logger.info(f"Matched embedded image for: {product.name[:30]}")
            else:
                # Fallback: Crop from rendered page image using LLM bbox
                cropped = self.crop_image_from_page(page_image, image_bbox)
                
                if cropped:
                    product_image = ProductImage(
                        image_id=f"img_{product.product_id}_{page_num}",
                        page_num=page_num,
                        bbox=cropped.bbox,
                        image_data=cropped.base64_data,
                        description=f"Product image for {product.name[:50]}"
                    )
                    
                    if not product.images:
                        product.images = []
                    product.images.append(product_image)
                    
                    logger.info(f"Cropped fallback image for: {product.name[:30]}")
                else:
                    logger.warning(f"Could not extract image for: {product.name[:30]}")
        
        return products
    
    @staticmethod
    def image_to_base64(image: Image.Image, format: str = "PNG") -> str:
        """Convert PIL Image to base64 string"""
        buffer = io.BytesIO()
        image.save(buffer, format=format)
        return base64.b64encode(buffer.getvalue()).decode('utf-8')
    
    @staticmethod
    def base64_to_image(base64_data: str) -> Optional[Image.Image]:
        """Convert base64 string to PIL Image"""
        try:
            image_data = base64.b64decode(base64_data)
            return Image.open(io.BytesIO(image_data))
        except Exception as e:
            logger.error(f"Failed to decode base64 image: {e}")
            return None
    
    @staticmethod
    def resize_image(
        image: Image.Image, 
        max_width: int = 200, 
        max_height: int = 200
    ) -> Image.Image:
        """Resize image maintaining aspect ratio"""
        ratio = min(max_width / image.width, max_height / image.height)
        if ratio < 1:
            new_size = (int(image.width * ratio), int(image.height * ratio))
            return image.resize(new_size, Image.Resampling.LANCZOS)
        return image


def get_image_extraction_prompt_addition(image_hint: str = "") -> str:
    """
    Get the prompt addition for image bounding box extraction.
    
    Args:
        image_hint: User-provided hint about image locations
    
    Returns:
        Prompt text to add to the main extraction prompt
    """
    hint_text = ""
    if image_hint.strip():
        hint_text = f"\nIMAGE LOCATION HINT: {image_hint.strip()}\n"
    
    return f"""
=== IMAGE EXTRACTION ===
For each product, also identify and return the bounding box of its associated product image.
{hint_text}
Return the image location as "image_bbox": [x1, y1, x2, y2] where:
- All values are PERCENTAGES (0.0 to 1.0) of the page dimensions
- x1, y1 = top-left corner
- x2, y2 = bottom-right corner
- Example: "image_bbox": [0.05, 0.1, 0.3, 0.4] means the image spans from 5%-30% horizontally and 10%-40% vertically

If no image is clearly associated with a product, omit the image_bbox field.
"""
