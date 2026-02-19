"""
Product Schema Definitions
Pydantic models for structured product data extraction
"""
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional, Dict, Any, Union
from enum import Enum
from datetime import datetime


class Currency(str, Enum):
    """Supported currencies"""
    EUR = "EUR"
    USD = "USD"
    GBP = "GBP"
    CHF = "CHF"
    UNKNOWN = "UNKNOWN"


class UnitType(str, Enum):
    """Unit types for quantities"""
    PIECE = "piece"
    KG = "kg"
    G = "g"
    L = "L"
    ML = "ml"
    M = "m"
    CM = "cm"
    MM = "mm"
    M2 = "m²"
    M3 = "m³"
    PACK = "pack"
    BOX = "box"
    PALLET = "pallet"
    SET = "set"
    UNKNOWN = "unknown"


class Price(BaseModel):
    """Price information"""
    amount: Optional[float] = Field(None, description="Numeric price value")
    currency: Currency = Field(
        default=Currency.UNKNOWN, description="Currency code")
    unit: Optional[str] = Field(
        None, description="Price per unit (e.g., 'per kg', 'each')")
    original_text: str = Field(
        "", description="Original price text from document")

    @field_validator('amount', mode='before')
    @classmethod
    def parse_amount(cls, v):
        if isinstance(v, str):
            # Remove currency symbols and parse
            cleaned = v.replace('€', '').replace('$', '').replace('£', '')
            cleaned = cleaned.replace(',', '.').strip()
            try:
                return float(cleaned)
            except ValueError:
                return None
        return v


class Dimensions(BaseModel):
    """Product dimensions"""
    length: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None
    diameter: Optional[float] = None
    unit: UnitType = Field(default=UnitType.CM)
    original_text: str = ""


class Weight(BaseModel):
    """Product weight"""
    value: Optional[float] = None
    unit: UnitType = Field(default=UnitType.KG)
    original_text: str = ""


class ProductVariant(BaseModel):
    """Product variant (size, color, etc.)"""
    variant_id: Optional[str] = None
    name: str = ""
    sku: Optional[str] = None
    price: Optional[Price] = None
    attributes: Dict[str, str] = Field(default_factory=dict)
    in_stock: Optional[bool] = None
    quantity_available: Optional[int] = None


class ProductImage(BaseModel):
    """Reference to product image in the catalog"""
    image_id: str = ""
    page_num: int = 0
    bbox: List[float] = Field(default_factory=list)  # Bounding box [x1, y1, x2, y2]
    description: str = ""
    image_data: Optional[str] = Field(None, description="Base64-encoded image data")


class Product(BaseModel):
    """
    Main product schema for catalog extraction.
    This schema covers the essential fields for most product catalogs.
    """
    # Identification
    product_id: str = Field("", description="Unique product identifier")
    sku: Optional[str] = Field(None, description="Stock Keeping Unit")
    ean: Optional[str] = Field(None, description="EAN/Barcode")
    reference: Optional[str] = Field(
        None, description="Product reference code")

    # Basic Info
    name: str = Field("", description="Product name/title")
    brand: Optional[str] = Field(None, description="Brand name")
    category: Optional[str] = Field(None, description="Product category")
    subcategory: Optional[str] = Field(None, description="Product subcategory")

    # Description
    description: str = Field("", description="Product description")
    short_description: Optional[str] = Field(None, description="Short summary")
    features: List[str] = Field(
        default_factory=list, description="Product features/bullet points")
    specifications: Dict[str, str] = Field(
        default_factory=dict, description="Technical specifications")

    # Pricing
    price: Optional[Price] = Field(None, description="Main price")
    original_price: Optional[Price] = Field(
        None, description="Original price before discount")
    discount_percentage: Optional[float] = Field(
        None, description="Discount percentage")

    # Physical Properties
    dimensions: Optional[Dimensions] = Field(
        None, description="Product dimensions")
    weight: Optional[Weight] = Field(None, description="Product weight")

    # Variants
    variants: List[ProductVariant] = Field(default_factory=list)

    # Media
    images: List[ProductImage] = Field(default_factory=list)

    # Availability
    in_stock: Optional[bool] = Field(None, description="Stock availability")
    stock_quantity: Optional[int] = Field(
        None, description="Available quantity")
    lead_time: Optional[str] = Field(None, description="Delivery lead time")

    # Catalog Info
    page_number: int = Field(0, description="Page number in catalog")
    position_on_page: Optional[str] = Field(
        None, description="Position (e.g., 'top-left')")

    # Source Tracking
    source_blocks: List[str] = Field(
        default_factory=list, description="Block IDs used for extraction")
    source_rows: List[int] = Field(
        default_factory=list, description="Excel row numbers this product was extracted from")
    excel_image: Optional[str] = Field(
        None, description="Base64-encoded image from Excel associated with this product")
    raw_text: str = Field("", description="Original raw text from PDF")
    extraction_confidence: float = Field(
        0.0, description="Confidence score 0-1")
    needs_review: bool = Field(True, description="Flag for QC review")

    # Dynamic attributes from custom schema fields
    raw_attributes: Dict[str, Any] = Field(
        default_factory=dict, description="Additional extracted attributes")

    # Metadata
    extracted_at: Optional[datetime] = Field(default_factory=datetime.now)

    def to_dict(self) -> dict:
        """Convert to dictionary with serializable values"""
        data = self.model_dump()
        if data.get('extracted_at'):
            data['extracted_at'] = data['extracted_at'].isoformat()
        return data

    @classmethod
    def from_raw_text(cls, text: str, page_num: int = 0, block_ids: List[str] = None):
        """Create a minimal product from raw text (for manual review)"""
        return cls(
            raw_text=text,
            page_number=page_num,
            source_blocks=block_ids or [],
            needs_review=True
        )


class ExtractionResult(BaseModel):
    """Result of extracting products from a page or region"""
    page_num: int
    products: List[Product] = Field(default_factory=list)
    # Block IDs not assigned to products
    unmatched_blocks: List[str] = Field(default_factory=list)
    coverage_percentage: float = Field(
        0.0, description="Percentage of blocks assigned to products")
    errors: List[str] = Field(default_factory=list)

    def add_product(self, product: Product):
        self.products.append(product)
        self._update_coverage()

    def _update_coverage(self):
        # Will be calculated based on block assignment
        pass


class CatalogExtractionResult(BaseModel):
    """Complete catalog extraction result"""
    file_name: str
    file_hash: str
    total_pages: int
    page_results: List[ExtractionResult] = Field(default_factory=list)

    @property
    def total_products(self) -> int:
        return sum(len(pr.products) for pr in self.page_results)

    @property
    def products_needing_review(self) -> int:
        count = 0
        for pr in self.page_results:
            count += sum(1 for p in pr.products if p.needs_review)
        return count

    @property
    def average_coverage(self) -> float:
        if not self.page_results:
            return 0.0
        return sum(pr.coverage_percentage for pr in self.page_results) / len(self.page_results)

    def get_all_products(self) -> List[Product]:
        products = []
        for pr in self.page_results:
            products.extend(pr.products)
        return products


# Schema for LLM prompt generation
PRODUCT_EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "products": {
            "type": "array",
            "description": "List of products found in the text",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Product name/title"},
                    "sku": {"type": "string", "description": "SKU or product code"},
                    "reference": {"type": "string", "description": "Reference number"},
                    "brand": {"type": "string", "description": "Brand name"},
                    "category": {"type": "string", "description": "Product category"},
                    "description": {"type": "string", "description": "Product description"},
                    "features": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of product features"
                    },
                    "price": {
                        "type": "object",
                        "properties": {
                            "amount": {"type": "number"},
                            "currency": {"type": "string", "enum": ["EUR", "USD", "GBP", "CHF"]},
                            "original_text": {"type": "string"}
                        }
                    },
                    "specifications": {
                        "type": "object",
                        "description": "Technical specifications as key-value pairs"
                    },
                    "dimensions": {"type": "string", "description": "Product dimensions"},
                    "weight": {"type": "string", "description": "Product weight"},
                    "in_stock": {"type": "boolean"},
                    "variants": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {"type": "string"},
                                "sku": {"type": "string"},
                                "price": {"type": "number"}
                            }
                        }
                    }
                },
                "required": ["name"]
            }
        }
    },
    "required": ["products"]
}


def get_extraction_prompt_template() -> str:
    """Get the prompt template for product extraction"""
    return """You are a product data extraction expert. Extract ALL products from the following catalog text.

RULES:
1. Extract EVERY product mentioned - do not skip any
2. If a field is not found, omit it (don't guess)
3. Preserve exact prices, SKUs, and references as they appear
4. If multiple products share common attributes, still list them separately
5. For product tables, extract each row as a separate product
6. Include any variants (sizes, colors) as separate entries or as variants array

TEXT TO EXTRACT FROM:
{text}

PAGE NUMBER: {page_num}

Respond with a JSON object containing a "products" array. Each product should have at minimum a "name" field.
If no products are found, return {{"products": []}}"""
