"""
Web Image Search module for fetching product images from the web.
Uses DuckDuckGo image search (no API key required).
"""

import logging
import requests
from typing import List, Optional
from dataclasses import dataclass
import re
import base64
from io import BytesIO
from PIL import Image

logger = logging.getLogger(__name__)


@dataclass
class WebImage:
    """Represents an image found on the web"""
    url: str
    thumbnail_url: str
    title: str
    source: str
    width: int = 0
    height: int = 0


def search_images(query: str, max_results: int = 20) -> List[WebImage]:
    """
    Search for images using DuckDuckGo.
    
    Args:
        query: Search query string
        max_results: Maximum number of results to return (default 20)
    
    Returns:
        List of WebImage objects
    """
    if not query or not query.strip():
        return []
    
    try:
        # Use the new ddgs package
        from ddgs import DDGS
        
        results = []
        with DDGS() as ddgs:
            images = ddgs.images(
                query,
                max_results=max_results,
                safesearch='moderate'
            )
            
            for img in images:
                results.append(WebImage(
                    url=img.get('image', ''),
                    thumbnail_url=img.get('thumbnail', ''),
                    title=img.get('title', ''),
                    source=img.get('source', ''),
                    width=img.get('width', 0),
                    height=img.get('height', 0)
                ))
        
        logger.info(f"Found {len(results)} images for query: {query}")
        return results
        
    except ImportError:
        logger.error("ddgs not installed. Run: pip install ddgs")
        return []
    except Exception as e:
        logger.error(f"Image search failed: {e}")
        return []


def download_image(url: str, timeout: int = 10) -> Optional[str]:
    """
    Download an image from URL and return as base64 string.
    
    Args:
        url: Image URL to download
        timeout: Request timeout in seconds
    
    Returns:
        Base64 encoded image string or None if failed
    """
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, timeout=timeout, headers=headers)
        response.raise_for_status()
        
        # Verify it's an image
        content_type = response.headers.get('Content-Type', '')
        if not content_type.startswith('image/'):
            logger.warning(f"Not an image: {content_type}")
            return None
        
        # Convert to PIL Image and back to ensure it's valid
        img = Image.open(BytesIO(response.content))
        
        # Convert to RGB if needed
        if img.mode in ('RGBA', 'P', 'LA'):
            img = img.convert('RGB')
        
        # Save to buffer
        buffer = BytesIO()
        img.save(buffer, format='JPEG', quality=85)
        
        return base64.b64encode(buffer.getvalue()).decode('utf-8')
        
    except Exception as e:
        logger.warning(f"Failed to download image from {url}: {e}")
        return None


def build_search_query(product_data: dict, field: str) -> str:
    """
    Build a search query from product data using the specified field.
    
    Args:
        product_data: Dictionary of product attributes
        field: Field name to use for search (e.g., 'name', 'sku', 'reference')
    
    Returns:
        Search query string
    """
    value = product_data.get(field, '')
    if not value:
        # Fallback to name
        value = product_data.get('name', '')
    
    # Clean up the query - remove special characters
    query = re.sub(r'[^\w\s-]', ' ', str(value))
    query = ' '.join(query.split())  # Normalize whitespace
    
    return query
