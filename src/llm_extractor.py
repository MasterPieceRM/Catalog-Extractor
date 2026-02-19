"""
LLM Structuring Module
Uses LLM to extract structured product data from text blocks or page images
"""
import json
import logging
import base64
import io
from typing import List, Dict, Any, Optional, Union
from dataclasses import dataclass
import re

from PIL import Image
from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential

from src.config import API_KEY, LLM_MODEL, LLM_BASE_URL, MAX_TOKENS_PER_CHUNK, LLM_VISION_ENABLED
from src.schemas import (
    Product, Price, Currency, Dimensions, Weight,
    ProductVariant, ExtractionResult,
    get_extraction_prompt_template, PRODUCT_EXTRACTION_SCHEMA
)

logger = logging.getLogger(__name__)


@dataclass
class ChunkedText:
    """Text chunk with metadata for processing"""
    text: str
    page_num: int
    block_ids: List[str]
    chunk_index: int
    total_chunks: int


class LLMExtractor:
    """Handles LLM-based product extraction and structuring"""

    def __init__(self, api_key: str = None, model: str = None, base_url: str = None,
                 custom_prompt: str = None, vision_enabled: bool = None):
        self.api_key = api_key or API_KEY
        self.model = model or LLM_MODEL
        self.base_url = base_url or LLM_BASE_URL
        self.vision_enabled = vision_enabled if vision_enabled is not None else LLM_VISION_ENABLED

        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url
        )

        self.prompt_template = custom_prompt or get_extraction_prompt_template()

    def set_prompt_template(self, template: str):
        """Set a custom prompt template"""
        self.prompt_template = template

    def _image_to_base64(self, image: Image.Image) -> str:
        """Convert PIL Image to base64 string"""
        buffered = io.BytesIO()
        # Convert to RGB if necessary
        if image.mode in ('RGBA', 'P'):
            image = image.convert('RGB')
        image.save(buffered, format="JPEG", quality=85)
        return base64.b64encode(buffered.getvalue()).decode('utf-8')

    def set_prompt_template(self, template: str):
        """Set a custom prompt template"""
        self.prompt_template = template

    def _estimate_tokens(self, text: str) -> int:
        """Rough token estimation (4 chars per token average)"""
        return len(text) // 4

    def chunk_text(self, blocks: List[dict], page_num: int,
                   max_tokens: int = MAX_TOKENS_PER_CHUNK) -> List[ChunkedText]:
        """Split blocks into chunks that fit within token limits"""
        chunks = []
        current_text = []
        current_block_ids = []
        current_tokens = 0
        chunk_idx = 0

        for block in blocks:
            block_text = block.get("text", "")
            block_tokens = self._estimate_tokens(block_text)

            if current_tokens + block_tokens > max_tokens and current_text:
                # Save current chunk
                chunks.append(ChunkedText(
                    text="\n\n".join(current_text),
                    page_num=page_num,
                    block_ids=current_block_ids.copy(),
                    chunk_index=chunk_idx,
                    total_chunks=0  # Will update later
                ))
                chunk_idx += 1
                current_text = []
                current_block_ids = []
                current_tokens = 0

            current_text.append(block_text)
            current_block_ids.append(block.get("block_id", ""))
            current_tokens += block_tokens

        # Don't forget the last chunk
        if current_text:
            chunks.append(ChunkedText(
                text="\n\n".join(current_text),
                page_num=page_num,
                block_ids=current_block_ids,
                chunk_index=chunk_idx,
                total_chunks=0
            ))

        # Update total chunks count
        for chunk in chunks:
            chunk.total_chunks = len(chunks)

        return chunks

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def _call_llm(self, prompt: str) -> str:
        """Call LLM with retry logic (text only)"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a product data extraction expert. Always respond with valid JSON only. No explanations, no markdown, just pure JSON starting with { and ending with }."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.1,
                max_tokens=4000
            )
            content = response.choices[0].message.content
            return self._clean_llm_response(content)
        except Exception as e:
            logger.error(f"LLM call failed: {e}")
            raise

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def _call_llm_vision(self, prompt: str, image: Image.Image) -> str:
        """Call vision LLM with an image"""
        try:
            # Convert image to base64
            image_base64 = self._image_to_base64(image)

            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a product data extraction expert analyzing catalog pages. Always respond with valid JSON only. No explanations, no markdown, just pure JSON starting with { and ending with }."
                    },
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{image_base64}"
                                }
                            }
                        ]
                    }
                ],
                temperature=0.1,
                max_tokens=4000
            )
            content = response.choices[0].message.content
            logger.info(
                f"Vision LLM Response (first 500 chars): {content[:500] if content else 'None'}")
            return self._clean_llm_response(content)
        except Exception as e:
            logger.error(f"Vision LLM call failed: {e}")
            raise

    def _clean_llm_response(self, content: str) -> str:
        """Clean up LLM response to extract JSON"""
        if not content:
            return content

        content = content.strip()
        # Remove markdown code blocks if present
        if content.startswith("```json"):
            content = content[7:]
        elif content.startswith("```"):
            content = content[3:]
        if content.endswith("```"):
            content = content[:-3]
        content = content.strip()

        # Try to find JSON object in the response
        if not content.startswith("{"):
            # Look for first { and last }
            start = content.find("{")
            end = content.rfind("}") + 1
            if start != -1 and end > start:
                content = content[start:end]

        return content

    def extract_products_from_image(self, image: Image.Image, page_num: int) -> List[Product]:
        """Extract products from a page image using vision LLM"""
        # Build a simpler prompt for vision (no text placeholder needed)
        prompt = self.prompt_template.replace(
            "{text}", "[See attached image]").replace("{page_num}", str(page_num))

        try:
            response = self._call_llm_vision(prompt, image)
            logger.info(
                f"Vision LLM Response (first 500 chars): {response[:500] if response else 'None'}")
            result = self._parse_llm_response(response, page_num, [], "")
            return result
        except Exception as e:
            logger.error(f"Vision product extraction failed: {e}")
            return []

    def extract_products_from_excel(
        self,
        data_text: str,
        batch_num: int,
        schema_fields: List[Dict[str, Any]] = None,
        layout_context: str = "",
        headers: List[str] = None,
    ) -> List[Product]:
        """
        Extract products from Excel tabular data using LLM.

        Args:
            data_text: Formatted text representation of spreadsheet rows
            batch_num: Batch number for tracking
            schema_fields: List of schema field dicts with name, type, description, hint
            layout_context: Optional description of the sheet layout
            headers: Column headers for reference
        """
        # Build field descriptions with hints
        field_lines = []
        if schema_fields:
            for field in schema_fields:
                field_name = field.get('name', 'unknown')
                req = "REQUIRED" if field.get('required') else "optional"
                field_type = field.get('type', 'text')
                desc = field.get('description', '')
                hint = field.get('hint', '')

                field_line = f"  • {field_name} ({req}, {field_type}): {desc}"
                if hint:
                    field_line += f"\n      → EXTRACTION HINT: {hint}"
                field_lines.append(field_line)

        if not field_lines:
            field_lines = ["  • name (REQUIRED, text): Product name"]

        fields_str = "\n".join(field_lines)

        # Build the prompt
        prompt_parts = [
            "You are a product data extraction expert analyzing an Excel spreadsheet.",
            ""
        ]

        if layout_context and layout_context.strip():
            prompt_parts.append("=== SHEET LAYOUT DESCRIPTION ===")
            prompt_parts.append(layout_context.strip())
            prompt_parts.append("")

        prompt_parts.extend([
            "=== FIELDS TO EXTRACT ===",
            fields_str,
            "",
        ])

        if headers:
            prompt_parts.append("=== COLUMN HEADERS ===")
            prompt_parts.append(str(headers))
            prompt_parts.append("")

        prompt_parts.extend([
            "=== SPREADSHEET DATA ===",
            data_text,
            "",
            "=== EXTRACTION RULES ===",
            "1. Extract EVERY product from the rows above",
            "2. Follow the EXTRACTION HINTS carefully — they describe transformations, calculations, and special handling",
            "3. If a hint says to apply a discount, calculate the discounted value",
            "4. If a hint says to combine columns, do so exactly as described",
            "5. If multiple rows belong to the same product, merge them into one product",
            "6. If a field value is not found, omit it (do not guess or invent values)",
            "7. Preserve exact text for SKUs, codes, references as they appear",
            "8. Return ONLY valid JSON, no explanations",
            "",
            "=== RESPONSE FORMAT ===",
            'Respond with a JSON object: {"products": [{...}, {...}]}',
            'Each product MUST include a "_source_rows" field: a list of row numbers (from the data above) that this product was extracted from.',
            'Each product should have at minimum a "name" field.',
        ])

        prompt = "\n".join(prompt_parts)

        try:
            response = self._call_llm(prompt)
            logger.info(
                f"Excel LLM Response batch {batch_num} (first 500 chars): {response[:500] if response else 'None'}")
            result = self._parse_llm_response(
                response, batch_num, [], data_text)
            return result
        except Exception as e:
            logger.error(
                f"Excel product extraction failed for batch {batch_num}: {e}")
            return []

    def learn_excel_pattern(
        self,
        headers: List[str],
        sample_rows: List[List[Any]],
        schema_fields: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        """
        Send a small sample of Excel data to the LLM to learn the column→field mapping.
        Returns a mapping dict that can be applied to all rows without further LLM calls.

        The mapping format:
        {
            "column_mappings": {
                "<schema_field_name>": {
                    "source_columns": ["ColA"],  // list of column names to pull from
                    "transform": "none" | "extract_number" | "concatenate" | "first_non_empty",
                    "fixed_value": null | "some value"  // if the field is always the same
                }
            },
            "row_grouping": "one_per_row" | "multi_row",
            "group_key_column": null | "column name"  // for multi_row grouping
        }
        """
        # Build sample text
        sample_text_lines = [f"COLUMNS: {headers}"]
        for i, row in enumerate(sample_rows):
            row_dict = {h: v for h, v in zip(headers, row)}
            sample_text_lines.append(
                f"Row {i+1}: {json.dumps(row_dict, default=str)}")
        sample_text = "\n".join(sample_text_lines)

        # Build field descriptions
        field_info = []
        for f in schema_fields:
            hint = f.get('hint', '')
            desc = f.get('description', '')
            line = f"  - {f['name']} ({f.get('type','text')}): {desc}"
            if hint:
                line += f" [HINT: {hint}]"
            field_info.append(line)
        fields_str = "\n".join(field_info)

        prompt = f"""You are a data mapping expert. Analyze this spreadsheet sample and create a column-to-field mapping.

=== TARGET SCHEMA FIELDS ===
{fields_str}

=== SPREADSHEET SAMPLE ===
{sample_text}

=== TASK ===
Create a JSON mapping that describes how to convert each spreadsheet row into a product with the target schema fields.

Rules:
1. Map each schema field to one or more source columns
2. If a field can't be found in any column, set source_columns to [] and fixed_value to null
3. Use "extract_number" transform for price fields that contain currency symbols or text
4. Use "concatenate" if multiple columns should be joined (e.g. combining separate dimension columns)
5. Use "first_non_empty" if multiple columns could contain the value and you want the first non-empty one
6. Determine if each row is one product ("one_per_row") or if multiple rows form one product ("multi_row")
7. For multi_row grouping, specify which column is the grouping key

Return ONLY this JSON structure:
{{
    "column_mappings": {{
        "<field_name>": {{
            "source_columns": ["column_name1"],
            "transform": "none",
            "fixed_value": null
        }}
    }},
    "row_grouping": "one_per_row",
    "group_key_column": null,
    "notes": "brief explanation of the mapping logic"
}}"""

        try:
            response = self._call_llm(prompt)
            mapping = json.loads(response)
            logger.info(
                f"Learned Excel pattern: {json.dumps(mapping, indent=2)[:500]}")
            return mapping
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse pattern mapping: {e}")
            # Try repair
            repaired = self._repair_json(response)
            if repaired:
                try:
                    return json.loads(repaired)
                except json.JSONDecodeError:
                    pass
            raise ValueError(f"LLM returned invalid mapping JSON: {e}")
        except Exception as e:
            logger.error(f"Pattern learning failed: {e}")
            raise

    def extract_products_from_text(self, text: str, page_num: int,
                                   block_ids: List[str] = None) -> List[Product]:
        """Extract products from a single text chunk"""
        if not text.strip():
            return []

        prompt = self.prompt_template.format(text=text, page_num=page_num)

        try:
            response = self._call_llm(prompt)
            logger.info(
                f"LLM Response (first 500 chars): {response[:500] if response else 'None'}")
            result = self._parse_llm_response(
                response, page_num, block_ids or [], text)
            return result
        except Exception as e:
            logger.error(f"Product extraction failed: {e}")
            # Return a placeholder product for manual review
            return [Product.from_raw_text(text, page_num, block_ids)]

    def _repair_json(self, text: str) -> str:
        """Try to repair truncated JSON by adding missing closing brackets/braces"""
        if not text:
            return None

        # Count opening and closing brackets
        open_braces = text.count('{')
        close_braces = text.count('}')
        open_brackets = text.count('[')
        close_brackets = text.count(']')

        # Build closing sequence
        closing = ''

        # Add missing close brackets
        for _ in range(open_brackets - close_brackets):
            closing += ']'

        # Add missing close braces
        for _ in range(open_braces - close_braces):
            closing += '}'

        if not closing:
            return None  # Nothing to repair

        repaired = text.rstrip() + closing
        return repaired

    def _parse_llm_response(self, response: str, page_num: int,
                            block_ids: List[str], raw_text: str = "") -> List[Product]:
        """Parse LLM JSON response into Product objects"""
        if not response:
            logger.error("Empty response from LLM")
            return [Product.from_raw_text(raw_text, page_num, block_ids)]

        try:
            data = json.loads(response)
        except json.JSONDecodeError as e:
            # Try to repair truncated JSON
            logger.warning(f"JSON parse failed, attempting repair: {e}")
            repaired = self._repair_json(response)
            if repaired:
                try:
                    data = json.loads(repaired)
                    logger.info("Successfully repaired JSON response")
                except json.JSONDecodeError:
                    logger.error(f"Failed to parse even repaired JSON")
                    logger.error(f"Response was: {response[:500]}")
                    return [Product.from_raw_text(raw_text, page_num, block_ids)]
            else:
                logger.error(f"Failed to parse LLM response as JSON: {e}")
                logger.error(f"Response was: {response[:500]}")
                return [Product.from_raw_text(raw_text, page_num, block_ids)]

        products = []

        # Handle different response formats
        if isinstance(data, list):
            raw_products = data
        elif isinstance(data, dict):
            raw_products = data.get("products", [])
            # If no products key, maybe the dict IS a product
            if not raw_products and "name" in data:
                raw_products = [data]
        else:
            raw_products = []

        for idx, raw_product in enumerate(raw_products):
            try:
                product = self._convert_to_product(
                    raw_product, page_num, block_ids, idx)
                products.append(product)
            except Exception as e:
                logger.warning(f"Failed to convert product {idx}: {e}")
                continue

        return products

    def _convert_to_product(self, raw: dict, page_num: int,
                            block_ids: List[str], idx: int) -> Product:
        """Convert raw LLM output to Product schema"""
        # Parse price
        price = None
        if raw.get("price"):
            price_data = raw["price"] if isinstance(raw["price"], dict) else {
                "amount": raw["price"]}
            price = Price(
                amount=price_data.get("amount"),
                currency=Currency(price_data.get("currency", "UNKNOWN")),
                original_text=str(price_data.get("original_text", ""))
            )

        # Parse dimensions
        dimensions = None
        if raw.get("dimensions"):
            if isinstance(raw["dimensions"], str):
                dimensions = Dimensions(original_text=raw["dimensions"])
            else:
                dimensions = Dimensions(**raw["dimensions"])

        # Parse weight
        weight = None
        if raw.get("weight"):
            if isinstance(raw["weight"], str):
                weight = Weight(original_text=raw["weight"])
            else:
                weight = Weight(**raw["weight"])

        # Parse variants
        variants = []
        for var in raw.get("variants", []):
            var_price = None
            if var.get("price"):
                var_price = Price(amount=var["price"])
            variants.append(ProductVariant(
                name=var.get("name", ""),
                sku=var.get("sku"),
                price=var_price
            ))

        # Parse features
        features = raw.get("features", [])
        if isinstance(features, str):
            features = [f.strip() for f in features.split(",")]

        # Parse specifications
        specs = raw.get("specifications", {})
        if isinstance(specs, str):
            specs = {}

        # Create product with unique ID
        product_id = raw.get("sku") or raw.get(
            "reference") or f"p{page_num}_{idx}"

        # Extract source rows from LLM response (for Excel batch extraction)
        source_rows = raw.get("_source_rows", []) or raw.get("source_rows", [])
        if isinstance(source_rows, (int, float)):
            source_rows = [int(source_rows)]
        elif isinstance(source_rows, list):
            source_rows = [int(r) for r in source_rows if r is not None]
        else:
            source_rows = []

        return Product(
            product_id=product_id,
            sku=raw.get("sku"),
            ean=raw.get("ean"),
            reference=raw.get("reference"),
            name=raw.get("name", "Unknown Product"),
            brand=raw.get("brand"),
            category=raw.get("category"),
            subcategory=raw.get("subcategory"),
            description=raw.get("description", ""),
            short_description=raw.get("short_description"),
            features=features,
            specifications=specs,
            price=price,
            dimensions=dimensions,
            weight=weight,
            variants=variants,
            in_stock=raw.get("in_stock"),
            stock_quantity=raw.get("stock_quantity"),
            page_number=page_num,
            source_blocks=block_ids,
            source_rows=source_rows,
            raw_text=raw.get("_raw_text", ""),
            extraction_confidence=0.8,  # Base confidence
            needs_review=True,
            raw_attributes=self._extract_raw_attributes(raw)
        )

    def _extract_raw_attributes(self, raw: dict) -> dict:
        """Extract all attributes from raw LLM response for custom schema fields"""
        # Known fields that are handled specially (but NOT image_bbox - we want that in raw_attributes)
        known_fields = {
            'sku', 'ean', 'reference', 'name', 'brand', 'category', 'subcategory',
            'description', 'short_description', 'features', 'specifications',
            'price', 'dimensions', 'weight', 'variants', 'in_stock', 'stock_quantity',
            '_raw_text', 'product_id', '_source_rows', 'source_rows',
        }

        # Store all other fields as raw_attributes (including image_bbox for image extraction)
        raw_attrs = {}
        for key, value in raw.items():
            if key not in known_fields and value is not None:
                raw_attrs[key] = value

        return raw_attrs

    def extract_from_page(self, blocks: List[dict], page_num: int) -> ExtractionResult:
        """Extract all products from a page's blocks"""
        result = ExtractionResult(page_num=page_num)

        if not blocks:
            result.coverage_percentage = 100.0  # No blocks = fully covered
            return result

        # Chunk the blocks if needed
        chunks = self.chunk_text(blocks, page_num)

        all_extracted_block_ids = set()

        for chunk in chunks:
            products = self.extract_products_from_text(
                chunk.text,
                chunk.page_num,
                chunk.block_ids
            )

            for product in products:
                result.add_product(product)
                all_extracted_block_ids.update(product.source_blocks)

        # Calculate coverage
        all_block_ids = set(b.get("block_id", "")
                            for b in blocks if b.get("text"))
        unmatched = all_block_ids - all_extracted_block_ids
        result.unmatched_blocks = list(unmatched)

        if all_block_ids:
            result.coverage_percentage = (
                len(all_block_ids) - len(unmatched)) / len(all_block_ids) * 100
        else:
            result.coverage_percentage = 100.0

        return result


class ProductRefiner:
    """Refines and validates extracted products"""

    def __init__(self, llm_extractor: LLMExtractor = None):
        self.llm = llm_extractor or LLMExtractor()

    def validate_product(self, product: Product) -> tuple[bool, List[str]]:
        """Validate a product and return (is_valid, list_of_issues)"""
        issues = []

        # Required field checks
        if not product.name or product.name == "Unknown Product":
            issues.append("Missing product name")

        # Price validation
        if product.price and product.price.amount:
            if product.price.amount < 0:
                issues.append("Negative price")
            if product.price.amount > 1000000:
                issues.append("Suspiciously high price")

        # SKU format validation (basic)
        if product.sku:
            if len(product.sku) < 2:
                issues.append("SKU too short")

        # Confidence check
        if product.extraction_confidence < 0.5:
            issues.append("Low extraction confidence")

        is_valid = len(issues) == 0
        return is_valid, issues

    def merge_duplicate_products(self, products: List[Product]) -> List[Product]:
        """Merge products that appear to be duplicates"""
        if not products:
            return products

        # Group by potential duplicate keys
        groups: Dict[str, List[Product]] = {}

        for product in products:
            # Create a key for grouping
            key_parts = []
            if product.sku:
                key_parts.append(f"sku:{product.sku}")
            if product.ean:
                key_parts.append(f"ean:{product.ean}")
            if not key_parts and product.name:
                # Use normalized name as fallback
                key_parts.append(f"name:{self._normalize_name(product.name)}")

            key = "|".join(
                key_parts) if key_parts else f"id:{product.product_id}"

            if key not in groups:
                groups[key] = []
            groups[key].append(product)

        # Merge each group
        merged = []
        for key, group in groups.items():
            if len(group) == 1:
                merged.append(group[0])
            else:
                merged.append(self._merge_products(group))

        return merged

    def _normalize_name(self, name: str) -> str:
        """Normalize product name for comparison"""
        return re.sub(r'\s+', ' ', name.lower().strip())

    def _merge_products(self, products: List[Product]) -> Product:
        """Merge multiple products into one, taking best data from each"""
        if len(products) == 1:
            return products[0]

        base = products[0].model_copy()

        for other in products[1:]:
            # Take non-empty values from other products
            if not base.description and other.description:
                base.description = other.description
            if not base.price and other.price:
                base.price = other.price
            if not base.brand and other.brand:
                base.brand = other.brand

            # Merge features
            base.features = list(set(base.features + other.features))

            # Merge specifications
            base.specifications.update(other.specifications)

            # Merge source blocks
            base.source_blocks = list(
                set(base.source_blocks + other.source_blocks))

        return base

    def enrich_product(self, product: Product, context: str = "") -> Product:
        """Use LLM to enrich product with additional details"""
        if not context:
            context = product.raw_text

        if not context:
            return product

        prompt = f"""Given this product information, extract any additional details that might be missing:

Current Product Data:
- Name: {product.name}
- SKU: {product.sku}
- Description: {product.description}
- Price: {product.price.amount if product.price else 'N/A'}

Additional Context:
{context}

Return a JSON with any additional fields you can extract:
- brand
- category
- features (as array)
- specifications (as object)
- dimensions
- weight

Only include fields you can confidently extract. Return empty object {{}} if no additional info found."""

        try:
            response = self.llm._call_llm(prompt)
            additional = json.loads(response)

            # Merge additional data
            if additional.get("brand") and not product.brand:
                product.brand = additional["brand"]
            if additional.get("category") and not product.category:
                product.category = additional["category"]
            if additional.get("features"):
                product.features = list(
                    set(product.features + additional["features"]))
            if additional.get("specifications"):
                product.specifications.update(additional["specifications"])

            product.extraction_confidence = min(
                product.extraction_confidence + 0.1, 1.0)

        except Exception as e:
            logger.warning(f"Product enrichment failed: {e}")

        return product
