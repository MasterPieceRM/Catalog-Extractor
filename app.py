"""
Streamlit QC/Review UI for PDF Catalog Product Extractor
Main application entry point
"""
from src.schemas import Product, ProductImage
from src.llm_extractor import LLMExtractor
from src.pdf_processor import PDFProcessor, PDFDocument
from src.image_extractor import ImageExtractor, get_image_extraction_prompt_addition
from src.excel_exporter import ExcelExporter
from src.excel_processor import ExcelProcessor, ExcelDocument
from src.config import UPLOAD_DIR, OUTPUT_DIR, PAGE_TITLE, LLM_VISION_ENABLED
import sys
import streamlit as st
import json
import base64
import io
import math
import re
from pathlib import Path
from datetime import datetime
from PIL import Image

# Page configuration must be first Streamlit command
st.set_page_config(
    page_title="Catalog Extractor",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))


# Default extraction schema fields with extraction hints
DEFAULT_SCHEMA_FIELDS = [
    {"name": "name", "type": "text", "required": True,
        "description": "Product name/title", "hint": ""},
    {"name": "sku", "type": "text", "required": False,
        "description": "SKU or product code", "hint": ""},
    {"name": "price", "type": "number", "required": False,
        "description": "Product price", "hint": ""},
    {"name": "description", "type": "text", "required": False,
        "description": "Product description", "hint": ""},
    {"name": "brand", "type": "text", "required": False,
        "description": "Brand name", "hint": ""},
    {"name": "category", "type": "text", "required": False,
        "description": "Product category", "hint": ""},
]

def evaluate_computed_fields(product, computed_fields: list):
    """Calculate values for computed fields based on product attributes"""
    results = {}
    if not computed_fields:
        return results

    # Helper functions for Excel compatibility
    def ROUNDUP(n, d=0):
        try:
            factor = 10 ** d
            return math.ceil(float(n) * factor) / factor
        except: return n

    def ROUNDDOWN(n, d=0):
        try:
            factor = 10 ** d
            return math.floor(float(n) * factor) / factor
        except: return n
        
    def IF(condition, true_val, false_val):
        return true_val if condition else false_val

    # Prepare context with math functions
    context = {
        "math": math, "round": round, "abs": abs, "min": min, "max": max,
        "ROUNDUP": ROUNDUP, "roundup": ROUNDUP,
        "ROUNDDOWN": ROUNDDOWN, "rounddown": ROUNDDOWN,
        "IF": IF, "if": IF,
        "SUM": sum, "sum": sum,
        "AVERAGE": lambda *args: sum(args)/len(args) if args else 0,
        "average": lambda *args: sum(args)/len(args) if args else 0
    }
    
    # Add product field values to context
    # Need to access schema fields to know what fields exist
    # But product object has attributes directly mapped or in raw_attributes
    
    # We'll just inspect the product object
    try:
        # 1. Add direct attributes
        for attr in dir(product):
            if not attr.startswith('_') and not callable(getattr(product, attr)):
                val = getattr(product, attr)
                # Handle price objects specially if needed
                if attr == 'price' and hasattr(val, 'amount'):
                    val = val.amount
                
                # Convert to numeric if possible (for math operations)
                if val is not None:
                    try:
                        context[attr] = float(str(val).replace(',', '').replace('$', '').strip())
                    except:
                        context[attr] = val
        
        # 2. Add raw attributes (which cover schema fields)
        if hasattr(product, 'raw_attributes') and product.raw_attributes:
            for k, v in product.raw_attributes.items():
                if k not in context: # Don't overwrite direct attrs
                    try:
                        context[k] = float(str(v).replace(',', '').replace('$', '').strip())
                    except:
                        context[k] = v
                        
    except Exception as e:
        pass # Best effort context building

    # Evaluate each formula
    for field in computed_fields:
        name = field['name']
        formula = field.get('formula', '')
        if not formula:
            results[name] = ""
            continue
            
        try:
            # Replace {field} with field variable name for python eval
            # Simple replacement of {name} -> name
            py_formula = formula
            
            # Find all {var} patterns
            vars_needed = re.findall(r'\{(\w+)\}', py_formula)
            for var in vars_needed:
                # Check if var is in context, simpler string replace
                py_formula = py_formula.replace(f"{{{var}}}", var)
                if var not in context:
                    context[var] = 0 # Default to 0 for missing fields to avoid NameError? Or None?
                    # Excel treats empty as 0 in math.
            
            # RESTRICTED Eval
            # Only allow builtins in our safe execution context (math functions)
            # Remove __builtins__ access
            res = eval(py_formula, {"__builtins__": None}, context)
            
            # Format result if float
            if isinstance(res, float):
                results[name] = round(res, 2)
            else:
                results[name] = res
                
        except Exception as e:
            results[name] = "Error" # Simplified error message for UI
            
    return results

# Default page layout context
DEFAULT_PAGE_CONTEXT = ""


def init_session_state():
    """Initialize session state variables"""
    defaults = {
        'pdf_doc': None,
        'current_page': 0,
        'products': [],
        'extraction_complete': False,
        'selected_product': None,
        'selected_product_idx': None,
        'view_mode': 'extraction',  # extraction, review, export, edit, schema
        'file_hash': None,
        'processing_status': None,
        'schema_fields': [f.copy() for f in DEFAULT_SCHEMA_FIELDS],
        'computed_fields': [],  # List of dicts: {'name': 'total_cost', 'formula': '{price} * {qty}'}
        'field_to_delete': None,  # For schema field deletion
        'page_layout_context': DEFAULT_PAGE_CONTEXT,  # Describes overall page layout
        'editing_field_hint': None,  # Track which field hint is being edited
        'use_ocr': False,  # Enable OCR for image-based PDFs
        'image_position_hint': '',  # Hint for where product images are typically located
        # Bulk extraction state
        'bulk_extraction_running': False,
        'bulk_extraction_cancel': False,
        'bulk_extraction_progress': 0,
        'bulk_extraction_total': 0,
        'bulk_extraction_current_page': 0,
        # List of (page_num, product_count) tuples
        'bulk_extraction_results': [],
        # App mode (tab selection)
        'app_mode': 'pdf',  # 'pdf' or 'excel'
        # Excel-specific state
        'excel_doc': None,
        'excel_products': [],  # Separate from PDF products
        'excel_current_sheet': None,
        'excel_file_hash': None,
        'excel_header_row': None,  # None=auto-detect, 0=no header, N=specific row
        'excel_start_row': 1,
        'excel_end_row': None,
        'excel_view_mode': 'extraction',  # extraction, review, export
        'excel_preview_cache': None,  # Cached preview DataFrame
        'excel_preview_cache_key': None,  # Cache key for invalidation
        'excel_include_images': True,  # Whether to extract images with products
        'excel_llm_batch_size': 20,  # Number of rows per LLM batch
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_sidebar():
    """Render the sidebar with file upload and navigation"""
    with st.sidebar:
        # Dynamic title based on mode
        if st.session_state.app_mode == 'pdf':
            st.title("📄 PDF Extractor")
        else:
            st.title("📊 Excel Extractor")
        st.divider()

        # Mode selector (PDF vs Excel) - buttons style
        st.subheader("📁 Mode")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📄 PDF", width='stretch',
                         type="primary" if st.session_state.app_mode == 'pdf' else "secondary",
                         key="mode_pdf_btn"):
                st.session_state.app_mode = 'pdf'
                st.rerun()
        with col2:
            if st.button("📊 Excel", width='stretch',
                         type="primary" if st.session_state.app_mode == 'excel' else "secondary",
                         key="mode_excel_btn"):
                st.session_state.app_mode = 'excel'
                st.rerun()

        st.divider()

        # Show mode-specific sidebar content
        if st.session_state.app_mode == 'pdf':
            render_pdf_sidebar_content()
        else:
            render_excel_sidebar_content()


def render_pdf_sidebar_content():
    """Render PDF-specific sidebar content"""
    with st.sidebar:
        # File upload
        st.subheader("📄 Upload PDF Catalog")
        uploaded_file = st.file_uploader(
            "Choose a PDF file",
            type=['pdf'],
            help="Upload a product catalog PDF to extract products"
        )

        if uploaded_file:
            # Save uploaded file
            file_path = UPLOAD_DIR / uploaded_file.name
            with open(file_path, 'wb') as f:
                f.write(uploaded_file.getvalue())

            # Check if it's a new file
            processor = PDFProcessor()
            new_hash = processor.get_file_hash(file_path)

            if new_hash != st.session_state.file_hash:
                st.session_state.file_hash = new_hash
                st.session_state.pdf_file_path = file_path  # Store path for image extraction
                # Reset everything when a new file is loaded
                st.session_state.pdf_doc = None
                st.session_state.products = []
                st.session_state.extraction_complete = False
                st.session_state.current_page = 0
                st.session_state.selected_product = None
                st.session_state.selected_product_idx = None
                st.session_state.view_mode = 'extraction'
                st.session_state.processing_status = None
                st.session_state.schema_fields = [
                    f.copy() for f in DEFAULT_SCHEMA_FIELDS]
                st.session_state.page_layout_context = DEFAULT_PAGE_CONTEXT
                st.session_state.editing_field_hint = None
                st.session_state.field_to_delete = None
                if 'page_products' in st.session_state:
                    st.session_state.page_products = {}
                if 'extraction_success' in st.session_state:
                    st.session_state.extraction_success = None
                if 'extraction_error' in st.session_state:
                    st.session_state.extraction_error = None
                if 'run_vision_extraction' in st.session_state:
                    st.session_state.run_vision_extraction = False
                st.info(f"New file loaded: {uploaded_file.name}")

            # Show vision mode status
            if LLM_VISION_ENABLED:
                st.success("🖼️ Vision Mode: ON")
                st.caption("AI will analyze page images directly")

            if st.button("🔄 Load/Reload PDF", width='stretch'):
                with st.spinner("Loading PDF..."):
                    st.session_state.pdf_doc = processor.extract_all_pages(
                        file_path)
                    st.session_state.current_page = 0

                st.success(
                    f"Loaded {st.session_state.pdf_doc.total_pages} pages")

        st.divider()

        # Navigation
        if st.session_state.pdf_doc:
            st.subheader("📑 Navigation")

            # Page selector with callback to ensure immediate update
            def on_page_change():
                st.session_state.current_page = st.session_state.page_selector - 1

            st.number_input(
                "Page",
                min_value=1,
                max_value=st.session_state.pdf_doc.total_pages,
                value=st.session_state.current_page + 1,
                step=1,
                key="page_selector",
                on_change=on_page_change
            )

            # View mode - vertical buttons
            st.subheader("👁️ View Mode")

            if st.button("📄 Extraction", width='stretch',
                         type="primary" if st.session_state.view_mode == 'extraction' else "secondary"):
                st.session_state.view_mode = 'extraction'
                st.rerun()

            if st.button("🔍 Review", width='stretch',
                         type="primary" if st.session_state.view_mode == 'review' else "secondary"):
                st.session_state.view_mode = 'review'
                st.rerun()

            if st.button("📤 Export", width='stretch',
                         type="primary" if st.session_state.view_mode == 'export' else "secondary"):
                st.session_state.view_mode = 'export'
                st.rerun()

            if st.button("📋 Schema", width='stretch',
                         type="primary" if st.session_state.view_mode == 'schema' else "secondary"):
                st.session_state.view_mode = 'schema'
                st.rerun()

            st.divider()

            # Product count
            st.subheader("📊 Stats")
            st.metric("Products", len(st.session_state.products))
        else:
            # Show schema config even without PDF loaded
            st.subheader("⚙️ Setup")
            if st.button("📋 Configure Schema", width='stretch'):
                st.session_state.view_mode = 'schema'
                st.rerun()

        st.divider()


def render_extraction_view():
    """Render the extraction view with PDF preview and extracted products"""
    if not st.session_state.pdf_doc:
        st.info("👆 Upload a PDF catalog to get started")
        st.markdown("""
        **Getting Started:**
        1. Upload a PDF catalog using the sidebar
        2. Click "Load/Reload PDF" to process it
        3. Go to **Schema** view to configure what fields to extract
        4. Come back here to extract products from each page
        """)
        return

    # Initialize page_products if not exists
    if 'page_products' not in st.session_state:
        st.session_state.page_products = {}  # {page_num: [products]}

    # Check if extraction was triggered
    page = st.session_state.pdf_doc.pages[st.session_state.current_page]

    if st.session_state.get('run_vision_extraction'):
        st.session_state.run_vision_extraction = False
        with st.spinner("🔄 Extracting products from page image..."):
            do_vision_extraction(page)
        st.rerun()

    # Show any previous extraction results
    if st.session_state.get('extraction_success'):
        st.success(f"✅ {st.session_state.extraction_success}")
        st.session_state.extraction_success = None
    if st.session_state.get('extraction_error'):
        st.error(f"❌ {st.session_state.extraction_error}")
        st.session_state.extraction_error = None

    col_pdf, col_products = st.columns([1, 1])

    page_num = page.page_num

    with col_pdf:
        st.subheader(
            f"📄 Page {page_num + 1} of {st.session_state.pdf_doc.total_pages}")

        # Display page image
        if page.image:
            st.image(page.image, width="stretch")
        else:
            st.warning("No image available for this page")

        # Show raw text toggle
        with st.expander("📝 View Raw Text", expanded=False):
            all_text = "\n\n".join([b.get("text", "")
                                   for b in page.blocks if b.get("text")])
            st.text_area("Page Text", value=all_text, height=200,
                         disabled=True, label_visibility="collapsed")

        st.divider()

        # Extract button - triggers extraction directly
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            def start_extraction():
                st.session_state.run_vision_extraction = True

            st.button("🔍 Extract Products", width="stretch",
                      type="primary", key="extract_page_btn", on_click=start_extraction)

        with col_btn2:
            if st.button("🗑️ Clear Page Results", width="stretch", key="clear_page_btn"):
                # Remove products from this page
                st.session_state.page_products[page_num] = []
                st.session_state.products = [
                    p for p in st.session_state.products if p.page_number != page_num]
                st.rerun()

        # === BULK EXTRACTION SECTION ===
        st.divider()

        # Page range selection for bulk extraction
        total_pages = st.session_state.pdf_doc.total_pages

        with st.expander("⚙️ Configure pages to extract", expanded=False):
            st.markdown("**Page Range**")
            col_start, col_end = st.columns(2)
            with col_start:
                start_page = st.number_input(
                    "From page",
                    min_value=1,
                    max_value=total_pages,
                    value=st.session_state.get('bulk_start_page', 1),
                    key="bulk_start_page_input"
                )
                st.session_state.bulk_start_page = start_page
            with col_end:
                end_page = st.number_input(
                    "To page",
                    min_value=1,
                    max_value=total_pages,
                    value=st.session_state.get('bulk_end_page', total_pages),
                    key="bulk_end_page_input"
                )
                st.session_state.bulk_end_page = end_page

            # Validate range
            if start_page > end_page:
                st.warning("Start page must be less than or equal to end page")

            pages_to_extract_count = max(0, end_page - start_page + 1)
            st.caption(
                f"Will extract: {pages_to_extract_count} pages (pages {start_page} to {end_page})")

        # Get range values
        start_p = st.session_state.get('bulk_start_page', 1)
        end_p = st.session_state.get('bulk_end_page', total_pages)
        pages_count = max(0, end_p - start_p + 1)
        btn_label = f"🚀 Extract Pages {start_p}-{end_p}" if pages_count < total_pages else f"🚀 Extract All Pages"

        # Extract All button - runs extraction inline with progress
        if st.button(btn_label, type="primary", key="extract_all_btn"):
            # Get pages to extract based on range
            start_p = st.session_state.get('bulk_start_page', 1)
            end_p = st.session_state.get('bulk_end_page', total_pages)

            # Validate range
            if start_p > end_p:
                st.warning(
                    "Invalid page range! Start page must be less than or equal to end page.")
            else:
                pages_list = list(range(start_p - 1, end_p)
                                  )  # Convert to 0-indexed

                if not pages_list:
                    st.warning("No pages to extract!")
                else:
                    # Set flag that bulk extraction is running
                    st.session_state.bulk_extraction_running = True

                    # Create a placeholder for the stop button and render it immediately
                    stop_placeholder = st.empty()
                    with stop_placeholder.container():
                        if st.button("⏹️ Stop Extraction", key="stop_bulk_btn", type="secondary"):
                            st.session_state.stop_bulk_extraction = True

                    # Use st.status for real-time progress
                    with st.status(f"Extracting {len(pages_list)} pages...", expanded=True) as status:
                        results = []
                        stopped = False

                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        results_container = st.container()

                        for i, page_idx in enumerate(pages_list):
                            # Check for stop (use session state flag)
                            if st.session_state.get('stop_bulk_extraction'):
                                stopped = True
                                st.session_state.stop_bulk_extraction = False
                                break

                            # Update progress
                            progress = (i) / len(pages_list)
                            progress_bar.progress(progress)
                            status_text.markdown(
                                f"**Processing page {page_idx + 1}** ({i + 1}/{len(pages_list)})")

                            # Count products before
                            products_before = len(
                                [p for p in st.session_state.products if p.page_number == page_idx])

                            # Extract
                            current_page_obj = st.session_state.pdf_doc.pages[page_idx]
                            try:
                                do_vision_extraction(current_page_obj)
                            except Exception as e:
                                with results_container:
                                    st.warning(
                                        f"Page {page_idx + 1}: Error - {str(e)[:50]}")
                                results.append((page_idx, 0))
                                continue

                            # Count products after
                            products_after = len(
                                [p for p in st.session_state.products if p.page_number == page_idx])
                            products_extracted = products_after - products_before
                            results.append((page_idx, products_extracted))

                            # Show result
                            with results_container:
                                st.text(
                                    f"Page {page_idx + 1}: {products_extracted} products")

                            # Rate limit delay (30 req/min = 2 sec between requests)
                            if i < len(pages_list) - 1:  # Don't delay after last page
                                import time
                                time.sleep(2)

                        # Final progress
                        progress_bar.progress(1.0)

                        # Summary
                        total_products = sum(r[1] for r in results)
                        pages_done = len(results)

                        if stopped:
                            status.update(
                                label=f"⏹️ Stopped: {total_products} products from {pages_done} pages", state="complete")
                        else:
                            status.update(
                                label=f"✅ Done: {total_products} products from {pages_done} pages", state="complete")

                    # Clear stop placeholder and running flag
                    stop_placeholder.empty()
                    st.session_state.bulk_extraction_running = False

    with col_products:
        st.subheader("📦 Extracted Products")

        # Get products for this page
        page_products = [
            p for p in st.session_state.products if p.page_number == page_num]

        if not page_products:
            st.info(
                "No products extracted from this page yet.\n\nClick **Extract Products** to start.")

            # Show schema reminder
            with st.expander("💡 Extraction Tips", expanded=True):
                st.markdown("""
                **Before extracting:**
                1. Go to **Schema** view in the sidebar
                2. Define which fields to extract (name, price, color, etc.)
                3. Add **extraction hints** for each field to guide the AI
                4. Describe your **page layout** so the AI understands the structure
                
                **Example hints:**
                - *"Style name is at the top of the page and applies to all products"*
                - *"Color is listed above the price, expand abbreviations like BLK=Black"*
                """)
        else:
            # Action buttons for all products
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                if st.button("✅ Approve All", key="approve_all_page"):
                    for p in page_products:
                        p.needs_review = False
                    st.rerun()
            with col_b:
                if st.button("🔄 Refresh Page", key="refresh_page_btn"):
                    st.rerun()
            with col_c:
                if st.button("➡️ Next Page", key="next_page_btn"):
                    if st.session_state.current_page < st.session_state.pdf_doc.total_pages - 1:
                        st.session_state.current_page += 1
                        st.rerun()

            st.divider()

            # Display each product as an editable preview
            for idx, product in enumerate(page_products):
                render_product_preview(product, idx, page_num)


@st.fragment
def render_product_preview(product, idx: int, page_num: int):
    """Render a single product as an editable preview card"""
    review_icon = "⚠️" if product.needs_review else "✅"
    has_image = bool(product.images and any(
        img.image_data for img in product.images))
    image_icon = "🖼️" if has_image else ""

    with st.expander(f"{review_icon} {image_icon} {product.name[:40]}{'...' if len(product.name) > 40 else ''}", expanded=product.needs_review):
        # Display product image if available
        if has_image:
            col_img, col_fields = st.columns([1, 2])
            with col_img:
                for img in product.images:
                    if img.image_data:
                        try:
                            image_bytes = base64.b64decode(img.image_data)
                            st.image(
                                image_bytes, caption="Product Image", width=150)
                        except Exception as e:
                            st.warning(f"Could not display image: {e}")
                        break  # Only show first image

                # Remove Background Button
                if st.button("✂️ Remove BG", key=f"rembg_{page_num}_{idx}"):
                    if product.images and product.images[0].image_data:
                        with st.spinner("Removing background..."):
                            try:
                                # Suppress ONNX Runtime and Numba warnings
                                import warnings
                                import os
                                os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
                                with warnings.catch_warnings():
                                    warnings.filterwarnings(
                                        "ignore", category=UserWarning)
                                    warnings.filterwarnings(
                                        "ignore", category=DeprecationWarning)
                                    warnings.filterwarnings(
                                        "ignore", message=".*TBB.*")
                                    warnings.filterwarnings(
                                        "ignore", message=".*GPU.*")
                                    warnings.filterwarnings(
                                        "ignore", message=".*onnxruntime.*")
                                    import rembg
                                from PIL import Image
                                import io

                                # Decode
                                img_data = base64.b64decode(
                                    product.images[0].image_data)
                                input_img = Image.open(io.BytesIO(img_data))

                                # Remove BG (suppress warnings during removal too)
                                with warnings.catch_warnings():
                                    warnings.simplefilter("ignore")
                                    output_img = rembg.remove(input_img)

                                # Process to properly handle transparency
                                output_buffer = io.BytesIO()
                                output_img.save(output_buffer, format='PNG')
                                new_b64 = base64.b64encode(
                                    output_buffer.getvalue()).decode('utf-8')

                                # Update
                                product.images[0].image_data = new_b64
                                # st.success("Background removed!") # Optional, might flicker
                                st.rerun()
                            except Exception as e:
                                st.error(f"Failed to remove background: {e}")

                # Image adjustment controls
                with st.popover("🔧 Adjust Image"):
                    st.caption("Upload a new image or search the web")

                    # Option 2: Upload custom image
                    st.markdown("**Upload Custom Image**")
                    upload_key = f"upload_img_{page_num}_{idx}"
                    uploaded_img = st.file_uploader(
                        "Upload image",
                        type=['png', 'jpg', 'jpeg'],
                        key=upload_key,
                        label_visibility="collapsed"
                    )
                    if uploaded_img and not st.session_state.get(f"processed_{upload_key}"):
                        try:
                            from PIL import Image
                            import io
                            img_pil = Image.open(uploaded_img)
                            # Convert to base64
                            buffer = io.BytesIO()
                            img_pil.save(buffer, format='PNG')
                            new_b64 = base64.b64encode(
                                buffer.getvalue()).decode('utf-8')

                            if product.images:
                                product.images[0].image_data = new_b64
                                # Clear bbox since it's custom
                                product.images[0].bbox = None
                            else:
                                from src.schemas import ProductImage
                                product.images = [ProductImage(
                                    image_id=f"custom_{product.product_id}",
                                    page_num=page_num,
                                    image_data=new_b64,
                                    description="Custom uploaded image"
                                )]
                            st.session_state[f"processed_{upload_key}"] = True
                            st.success("Image replaced!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to upload: {e}")
                    elif not uploaded_img:
                        st.session_state[f"processed_{upload_key}"] = False

                    st.divider()

                    # Option 3: Search web for images
                    def web_image_search_fragment(prod, p_num, p_idx):
                        """Web image search - updates product and closes popover on selection"""
                        st.markdown("**🌐 Search Web for Image**")

                        # Get field names from schema
                        schema_fields_local = st.session_state.get(
                            'schema_fields', DEFAULT_SCHEMA_FIELDS)
                        available_fields = [f['name']
                                            for f in schema_fields_local]

                        # Multi-select for fields to combine
                        search_fields = st.multiselect(
                            "Combine fields for search",
                            available_fields,
                            default=[
                                'name'] if 'name' in available_fields else available_fields[:1],
                            key=f"search_fields_{p_num}_{p_idx}"
                        )

                        # Build combined search query from selected fields
                        search_parts = []
                        for field in search_fields:
                            value = getattr(prod, field, None) if hasattr(
                                prod, field) else prod.raw_attributes.get(field, '')
                            if value and str(value).strip():
                                search_parts.append(str(value).strip())

                        search_value = ' '.join(
                            search_parts) if search_parts else prod.name

                        search_clicked = st.button(
                            f"🔍 Search '{search_value[:40]}...'", key=f"web_search_{p_num}_{p_idx}")

                        # Perform search if clicked OR if loading flag is set
                        if search_clicked or st.session_state.get(f"web_search_loading_{prod.product_id}"):
                            if search_clicked:
                                st.session_state[f"web_search_query_{prod.product_id}"] = search_value

                            with st.spinner("Searching for images..."):
                                try:
                                    from src.web_image_search import search_images
                                    query = st.session_state.get(
                                        f"web_search_query_{prod.product_id}", search_value)
                                    results = search_images(
                                        query, max_results=10)
                                    st.session_state[f"web_search_results_{prod.product_id}"] = results
                                    st.session_state[f"web_search_loading_{prod.product_id}"] = False
                                except Exception as e:
                                    st.error(f"Search failed: {e}")

                        # Display search results with prettier gallery
                        web_results = st.session_state.get(
                            f"web_search_results_{prod.product_id}")
                        if web_results:
                            st.markdown(
                                f"##### 🖼️ Found {len(web_results)} images")
                            st.caption("Click 'Use This' to select an image")

                            # Check if we need to download a selected image
                            selected_key = f"selected_web_img_{prod.product_id}"
                            if st.session_state.get(selected_key) is not None:
                                selected_idx = st.session_state[selected_key]
                                result = web_results[selected_idx]
                                with st.spinner("⏳ Downloading..."):
                                    from src.web_image_search import download_image
                                    img_b64 = download_image(result.url)
                                    if img_b64:
                                        from src.schemas import ProductImage
                                        if prod.images:
                                            prod.images[0].image_data = img_b64
                                            prod.images[0].bbox = None
                                        else:
                                            prod.images = [ProductImage(
                                                image_id=f"web_{prod.product_id}",
                                                page_num=p_num,
                                                image_data=img_b64,
                                                description=f"Web image: {result.title[:50]}"
                                            )]
                                    st.session_state[selected_key] = None
                                    st.session_state[f"web_search_results_{prod.product_id}"] = None
                                    st.rerun(scope="app")

                            # Display in rows of 4 for better visibility
                            cols_per_row = 4
                            for row_start in range(0, len(web_results), cols_per_row):
                                cols = st.columns(cols_per_row)
                                for col_idx, result_idx in enumerate(range(row_start, min(row_start + cols_per_row, len(web_results)))):
                                    result = web_results[result_idx]
                                    with cols[col_idx]:
                                        try:
                                            st.image(
                                                result.thumbnail_url, width='stretch')

                                            # Show source domain and dimensions
                                            source_domain = result.source[:20] + "..." if len(
                                                result.source) > 20 else result.source
                                            info_text = f"📍 {source_domain}"
                                            if result.width and result.height:
                                                info_text += f" | 📐 {result.width}×{result.height}"
                                            st.caption(info_text)

                                            # Select button - just set session state, download happens on rerun
                                            if st.button("✅ Use This", key=f"select_web_img_{p_num}_{p_idx}_{result_idx}"):
                                                st.session_state[f"selected_web_img_{prod.product_id}"] = result_idx
                                                st.rerun(scope="app")
                                        except:
                                            pass  # Skip broken thumbnails

                            # Clear search button
                            if st.button("🗑️ Clear Results", key=f"clear_search_{p_num}_{p_idx}"):
                                st.session_state[f"web_search_results_{prod.product_id}"] = None
                                st.rerun(scope="app")

                    # Call the fragment helper
                    web_image_search_fragment(product, page_num, idx)
        else:
            # No image - show option to add one
            with st.expander("📷 Add Image", expanded=False):
                st.caption(
                    "This product has no image. Upload one or search the web.")

                # Upload option
                st.markdown("**Upload Image**")
                upload_no_img_key = f"upload_no_img_{page_num}_{idx}"
                uploaded_img = st.file_uploader(
                    "Upload image",
                    type=['png', 'jpg', 'jpeg'],
                    key=upload_no_img_key,
                    label_visibility="collapsed"
                )
                if uploaded_img and not st.session_state.get(f"processed_{upload_no_img_key}"):
                    try:
                        from PIL import Image
                        import io
                        img_pil = Image.open(uploaded_img)
                        buffer = io.BytesIO()
                        img_pil.save(buffer, format='PNG')
                        new_b64 = base64.b64encode(
                            buffer.getvalue()).decode('utf-8')

                        from src.schemas import ProductImage
                        product.images = [ProductImage(
                            image_id=f"upload_{product.product_id}",
                            page_num=page_num,
                            image_data=new_b64,
                            description="Uploaded image"
                        )]
                        st.session_state[f"processed_{upload_no_img_key}"] = True
                        st.success("Image added!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to upload: {e}")
                elif not uploaded_img:
                    st.session_state[f"processed_{upload_no_img_key}"] = False

                st.divider()

                # Web search option
                def web_search_no_image_fragment(prod, p_num, p_idx):
                    """Web image search for products without images"""
                    st.markdown("**🌐 Search Web for Image**")

                    schema_fields_local = st.session_state.get(
                        'schema_fields', DEFAULT_SCHEMA_FIELDS)
                    available_fields = [f['name'] for f in schema_fields_local]

                    search_fields = st.multiselect(
                        "Combine fields for search",
                        available_fields,
                        default=[
                            'name'] if 'name' in available_fields else available_fields[:1],
                        key=f"search_fields_noimg_{p_num}_{p_idx}"
                    )

                    search_parts = []
                    for field in search_fields:
                        value = getattr(prod, field, None) if hasattr(
                            prod, field) else prod.raw_attributes.get(field, '')
                        if value and str(value).strip():
                            search_parts.append(str(value).strip())

                    search_value = ' '.join(
                        search_parts) if search_parts else prod.name

                    search_clicked = st.button(
                        f"🔍 Search '{search_value[:40]}...'", key=f"web_search_noimg_{p_num}_{p_idx}")

                    if search_clicked or st.session_state.get(f"web_search_loading_noimg_{prod.product_id}"):
                        if search_clicked:
                            st.session_state[f"web_search_query_noimg_{prod.product_id}"] = search_value

                        with st.spinner("Searching for images..."):
                            try:
                                from src.web_image_search import search_images
                                query = st.session_state.get(
                                    f"web_search_query_noimg_{prod.product_id}", search_value)
                                results = search_images(query, max_results=10)
                                st.session_state[f"web_search_results_noimg_{prod.product_id}"] = results
                                st.session_state[f"web_search_loading_noimg_{prod.product_id}"] = False
                            except Exception as e:
                                st.error(f"Search failed: {e}")

                    web_results = st.session_state.get(
                        f"web_search_results_noimg_{prod.product_id}")
                    if web_results:
                        st.markdown(
                            f"##### 🖼️ Found {len(web_results)} images")
                        st.caption("Click 'Use This' to select an image")

                        # Display in rows of 4
                        cols_per_row = 4
                        for row_start in range(0, len(web_results), cols_per_row):
                            cols = st.columns(cols_per_row)
                            for col_idx, result_idx in enumerate(range(row_start, min(row_start + cols_per_row, len(web_results)))):
                                result = web_results[result_idx]
                                with cols[col_idx]:
                                    try:
                                        st.image(result.thumbnail_url,
                                                 width='stretch')

                                        # Show source domain and dimensions
                                        source_domain = result.source[:20] + "..." if len(
                                            result.source) > 20 else result.source
                                        info_text = f"📍 {source_domain}"
                                        if result.width and result.height:
                                            info_text += f" | 📐 {result.width}×{result.height}"
                                        st.caption(info_text)

                                        if st.button("✅ Use This", key=f"select_web_noimg_{p_num}_{p_idx}_{result_idx}"):
                                            from src.web_image_search import download_image
                                            with st.spinner("⏳ Downloading..."):
                                                img_b64 = download_image(
                                                    result.url)
                                                if img_b64:
                                                    from src.schemas import ProductImage
                                                    prod.images = [ProductImage(
                                                        image_id=f"web_{prod.product_id}",
                                                        page_num=p_num,
                                                        image_data=img_b64,
                                                        description=f"Web image: {result.title[:50]}"
                                                    )]
                                                    # Clear results to close the search view
                                                    st.session_state[
                                                        f"web_search_results_noimg_{prod.product_id}"] = None
                                                    st.rerun()
                                                else:
                                                    st.error(
                                                        "❌ Download failed")
                                    except:
                                        pass

                        # Clear search button
                        if st.button("🗑️ Clear Results", key=f"clear_search_noimg_{p_num}_{p_idx}"):
                            st.session_state[f"web_search_results_noimg_{prod.product_id}"] = None
                            st.rerun()

                web_search_no_image_fragment(product, page_num, idx)

            col_img = None
            col_fields = st.container()

        # Get all fields from schema
        with col_fields if has_image else st.container():
            schema_fields = st.session_state.get(
                'schema_fields', DEFAULT_SCHEMA_FIELDS)

            # Display/edit each field
            for field in schema_fields:
                field_name = field['name']
                field_type = field.get('type', 'text')

                # Get current value from product
                current_value = getattr(product, field_name, None) if hasattr(
                    product, field_name) else product.raw_attributes.get(field_name, '')

                # Handle price specially - could be Price object or string
                if field_name == 'price' and product.price:
                    if hasattr(product.price, 'amount'):
                        # It's a Price object - just show the amount
                        current_value = str(product.price.amount)
                    else:
                        # It's a string (from vision extraction)
                        current_value = str(product.price)

                # Display as text input for editing
                new_value = st.text_input(
                    f"{field_name.replace('_', ' ').title()}",
                    value=str(current_value) if current_value else "",
                    key=f"prod_{page_num}_{idx}_{field_name}"
                )

                # Update product if value changed
                if new_value != str(current_value or ""):
                    if hasattr(product, field_name):
                        setattr(product, field_name, new_value)
                    else:
                        product.raw_attributes[field_name] = new_value

        # Show any additional extracted attributes not in schema (excluding image_bbox)
        extra_attrs = {k: v for k, v in product.raw_attributes.items()
                       if k not in [f['name'] for f in schema_fields] and k != 'image_bbox' and v}
        if extra_attrs:
            with st.expander("Additional Extracted Data"):
                for key, value in extra_attrs.items():
                    st.text(f"{key}: {value}")

        # Action buttons
        col1, col2, col3 = st.columns(3)
        with col1:
            if product.needs_review:
                if st.button("✅ Approve", key=f"approve_{page_num}_{idx}"):
                    product.needs_review = False
                    st.rerun()
        with col2:
            if st.button("🗑️ Delete", key=f"delete_{page_num}_{idx}"):
                st.session_state.products = [
                    p for p in st.session_state.products if p.product_id != product.product_id]
                st.rerun()


def render_review_view():
    """Render the product review/QC view"""
    if not st.session_state.products:
        st.info("No products extracted yet. Go to Extraction view to extract products.")
        return

    st.subheader("🔍 Product Review & QC")

    # Filter options
    col1, col2 = st.columns([1, 2])
    with col1:
        filter_review = st.checkbox("Show only needs review", value=True)
    with col2:
        search_query = st.text_input(
            "🔍 Search products",
            placeholder="Search by name, SKU, or any field...",
            key="pdf_search_query",
            label_visibility="collapsed"
        )

    filter_page = st.selectbox(
        "Filter by page",
        ["All"] + [str(i+1) for i in range(
            st.session_state.pdf_doc.total_pages if st.session_state.pdf_doc else 0)],
        key="pdf_filter_page"
    )

    # Filter products
    products = st.session_state.products.copy()

    if filter_review:
        products = [p for p in products if p.needs_review]

    if filter_page != "All":
        products = [p for p in products if p.page_number ==
                    int(filter_page) - 1]

    # Apply search filter
    if search_query:
        query_lower = search_query.lower()
        filtered = []
        for p in products:
            if query_lower in (p.name or '').lower():
                filtered.append(p)
            elif query_lower in (getattr(p, 'sku', '') or '').lower():
                filtered.append(p)
            elif any(query_lower in str(v).lower() for v in p.raw_attributes.values()):
                filtered.append(p)
        products = filtered

    st.caption(f"Showing {len(products)} products")

    # Bulk action buttons
    col_approve, col_delete, col_spacer = st.columns([1, 1, 2])
    with col_approve:
        if st.button("✅ Approve All", key="approve_all_btn", type="primary"):
            for p in products:
                p.needs_review = False
            st.success(f"Approved {len(products)} products!")
            st.rerun()
    with col_delete:
        if st.button("🗑️ Delete All Shown", key="delete_all_shown_btn"):
            # Get product IDs to delete
            ids_to_delete = {p.product_id for p in products}
            st.session_state.products = [
                p for p in st.session_state.products if p.product_id not in ids_to_delete
            ]
            st.success(f"Deleted {len(ids_to_delete)} products!")
            st.rerun()

    # Display products
    for idx, product in enumerate(products):
        render_product_review_card(product, idx)


def render_product_review_card(product: Product, idx: int):
    """Render a product card for review using schema fields"""
    status_icon = "✅" if not product.needs_review else "⏳"

    with st.expander(
        f"{status_icon} {product.name[:50]}{'...' if len(product.name) > 50 else ''} | Page {product.page_number + 1}",
        expanded=product.needs_review
    ):
        # Check if product has image
        has_image = bool(product.images and any(
            img.image_data for img in product.images))

        # Layout with image on left if available
        if has_image:
            col_img, col_info, col_actions = st.columns([1, 2, 1])
            with col_img:
                for img in product.images:
                    if img.image_data:
                        try:
                            image_bytes = base64.b64decode(img.image_data)
                            st.image(image_bytes, width=100)
                        except Exception:
                            st.caption("🖼️ Image error")
                        break
        else:
            col_info, col_actions = st.columns([3, 1])

        with col_info:
            # Display fields from schema
            schema_fields = st.session_state.get(
                'schema_fields', DEFAULT_SCHEMA_FIELDS)

            for field in schema_fields:
                field_name = field['name']

                # Get value from product attribute or raw_attributes
                if hasattr(product, field_name):
                    value = getattr(product, field_name)
                else:
                    value = product.raw_attributes.get(field_name)

                # Handle price specially
                if field_name == 'price' and value:
                    if hasattr(value, 'amount'):
                        value = str(value.amount)
                    else:
                        value = str(value)

                # Display the field
                display_name = field_name.replace('_', ' ').title()
                st.markdown(f"**{display_name}:** {value or 'N/A'}")

            # Computed fields preview
            computed_fields = st.session_state.get('computed_fields', [])
            if computed_fields:
                st.markdown("---")
                st.caption("∑ Computed Fields")
                results = evaluate_computed_fields(product, computed_fields)
                for field in computed_fields:
                    name = field['name']
                    # Use unique key if creating widgets, but here we just use markdown
                    val = results.get(name, "")
                    st.markdown(f"**{name.replace('_', ' ').title()} (calc):** {val}")

            # Show any extra raw_attributes not in schema (excluding image_bbox)
            extra_attrs = {k: v for k, v in product.raw_attributes.items()
                           if k not in [f['name'] for f in schema_fields] and k != 'image_bbox' and v}
            if extra_attrs:
                with st.popover("📋 Additional Data"):
                    for key, value in extra_attrs.items():
                        st.markdown(f"**{key}:** {value}")

        with col_actions:
            # Go to page button
            if st.button("📄 Go to page", key=f"goto_{product.product_id}_{idx}"):
                st.session_state.current_page = product.page_number
                st.session_state.view_mode = 'extraction'
                st.rerun()

            # Approve button
            if st.button("✅ Approve", key=f"approve_{product.product_id}_{idx}"):
                product.needs_review = False
                st.success("Approved!")
                st.rerun()

            # Edit button
            if st.button("✏️ Edit", key=f"edit_{product.product_id}_{idx}"):
                st.session_state.selected_product = product
                st.session_state.view_mode = 'edit'
                st.rerun()

            # Delete button
            if st.button("🗑️ Delete", key=f"delete_{product.product_id}_{idx}"):
                st.session_state.products.remove(product)
                st.success("Deleted!")
                st.rerun()


def render_export_view():
    """Render the export view - CSV only with schema fields"""
    st.subheader("📤 Export Products")

    if not st.session_state.products:
        st.info("No products to export. Extract products first.")
        return

    # Summary
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Products", len(st.session_state.products))
    with col2:
        reviewed = sum(
            1 for p in st.session_state.products if not p.needs_review)
        st.metric("Reviewed", reviewed)
    with col3:
        needs_review = sum(
            1 for p in st.session_state.products if p.needs_review)
        st.metric("Needs Review", needs_review)

    st.divider()

    # Manage products before export
    st.subheader("🗂️ Manage Products")

    col_del1, col_del2 = st.columns(2)
    with col_del1:
        reviewed_count = sum(
            1 for p in st.session_state.products if not p.needs_review)
        if st.button(f"🗑️ Delete All Reviewed ({reviewed_count})", width="stretch"):
            st.session_state.products = [
                p for p in st.session_state.products if p.needs_review]
            st.rerun()
    with col_del2:
        unreviewed_count = sum(
            1 for p in st.session_state.products if p.needs_review)
        if st.button(f"🗑️ Delete All Unreviewed ({unreviewed_count})", width="stretch"):
            st.session_state.products = [
                p for p in st.session_state.products if not p.needs_review]
            st.rerun()

    st.divider()

    # Export options
    # Export options
    st.subheader("⚙️ Export Options")

    # Get schema fields
    schema_fields = st.session_state.get(
        'schema_fields', DEFAULT_SCHEMA_FIELDS)

    col1, col2 = st.columns(2)
    with col1:
        include_reviewed_only = st.checkbox(
            "Export reviewed only", value=False)
    with col2:
        include_page_info = st.checkbox("Include page number", value=True)

    # Filter products
    products_to_export = st.session_state.products
    if include_reviewed_only:
        products_to_export = [
            p for p in products_to_export if not p.needs_review]

    st.divider()

    # ===== EXCEL EXPORT =====
    st.subheader("📊 Excel Export")

    # Count products with images
    products_with_images = sum(
        1 for p in st.session_state.products
        if p.images and any(img.image_data for img in p.images)
    )
    st.caption(
        f"{products_with_images} of {len(st.session_state.products)} products have images")

    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        include_images_excel = st.checkbox(
            "Include images in Excel", value=True)
        if include_images_excel:
            remove_bg_all = st.checkbox("✂️ Remove backgrounds (Slow!)", value=False,
                                        help="Automatically remove backgrounds from all images during export")
    with col_opt2:
        image_size = st.selectbox(
            "Image size",
            ["small", "medium", "large"],
            index=1,
            help="Small: 60px, Medium: 100px, Large: 150px"
        )

    if not include_images_excel:
        remove_bg_all = False

    if st.button("📊 Generate Excel File", width="stretch", type="primary"):
        with st.spinner("Generating Excel file... (this may take a while if removing backgrounds)"):
            try:
                excel_exporter = ExcelExporter()
                excel_bytes = excel_exporter.export_products_to_excel(
                    products=products_to_export,
                    schema_fields=schema_fields,
                    computed_fields=st.session_state.get('computed_fields', []),
                    include_images=include_images_excel,
                    image_size=image_size,
                    include_page_info=include_page_info,
                    remove_bg=remove_bg_all
                )

                st.download_button(
                    "📥 Download Excel",
                    excel_bytes,
                    file_name=f"products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )
                st.success("✅ Excel file generated! Click above to download.")
            except Exception as e:
                st.error(f"Failed to generate Excel: {str(e)}")

    # Preview
    st.divider()
    st.subheader("📋 Preview")

    if st.session_state.products:
        import pandas as pd

        products_to_show = st.session_state.products[:10]  # Show first 10
        preview_rows = []

        for p in products_to_show:
            row = {}
            for field in schema_fields:
                field_name = field['name']
                if hasattr(p, field_name):
                    value = getattr(p, field_name)
                else:
                    value = p.raw_attributes.get(field_name, '')

                if field_name == 'price' and value:
                    if hasattr(value, 'amount'):
                        value = value.amount

                row[field_name] = value

            # Computed fields
            computed_fields = st.session_state.get('computed_fields', [])
            if computed_fields:
                comp_vals = evaluate_computed_fields(p, computed_fields)
                row.update(comp_vals)

            row['page'] = p.page_number + 1
            preview_rows.append(row)

        df = pd.DataFrame(preview_rows)
        st.dataframe(df, width="stretch")

        if len(st.session_state.products) > 10:
            st.caption(
                f"Showing first 10 of {len(st.session_state.products)} products")


def do_vision_extraction(page):
    """Actually perform vision extraction - called via session state trigger"""
    page_num = page.page_num

    try:
        # Get the schema prompt with hints
        prompt_template = get_schema_prompt()

        # Create extractor with vision enabled
        extractor = LLMExtractor(
            custom_prompt=prompt_template, vision_enabled=True)

        # Extract products from the page image
        products = extractor.extract_products_from_image(page.image, page_num)

        if not products:
            st.session_state.extraction_error = "No products found. The AI returned empty results."
            return

        # Extract product images using ImageExtractor with fitz for perfect cropping
        if page.image:
            image_extractor = ImageExtractor()
            pdf_path = st.session_state.get('pdf_file_path')
            products = image_extractor.extract_product_images(
                page.image, products, page_num, pdf_path=pdf_path
            )

        # Add products to session state
        added_count = 0
        images_extracted = 0
        for product in products:
            product.product_id = f"p{page_num}_{len(st.session_state.products)}"
            product.page_number = page_num
            product.needs_review = True
            st.session_state.products.append(product)
            added_count += 1
            if product.images:
                images_extracted += 1

        success_msg = f"Extracted {added_count} products from page {page_num + 1}!"
        if images_extracted > 0:
            success_msg += f" ({images_extracted} with images)"
        st.session_state.extraction_success = success_msg

    except Exception as e:
        import traceback
        st.session_state.extraction_error = f"Vision extraction failed: {str(e)}\n\n{traceback.format_exc()}"


def extract_products_from_page(page):
    """Extract products from a page using LLM (text mode only - vision is handled separately)"""
    page_num = page.page_num

    # Check if vision mode is enabled - vision extraction is handled in render_extraction_view
    use_vision = LLM_VISION_ENABLED and hasattr(
        page, 'image') and page.image is not None

    if use_vision:
        # Vision mode is handled by the button callback in render_extraction_view
        return

    # TEXT MODE - Original text-based extraction
    # Collect all text from the page
    all_text_parts = []
    for block in page.blocks:
        text = block.get("text", "").strip()
        if text:
            all_text_parts.append(text)

    if not all_text_parts:
        st.error("⚠️ No text found on this page!")
        st.warning("""
        **Possible reasons:**
        1. The PDF page is image-based (scanned document)
        2. Text is embedded in images
        3. The PDF uses non-standard fonts
        
        **Solution:** Enable OCR in the sidebar settings, or use a different PDF.
        """)
        return

    # Combine all text with clear separators
    combined_text = "\n---\n".join(all_text_parts)

    # Always show the raw text first so user can see what we're working with
    st.info(
        f"📝 **Text Mode** - Found {len(all_text_parts)} text blocks ({len(combined_text)} characters total)")

    with st.expander("👁️ View Raw Text Being Sent to AI", expanded=True):
        st.text_area(
            "Raw text from page",
            combined_text,
            height=300,
            disabled=True,
            label_visibility="collapsed"
        )
        st.caption(
            "This is ALL the text the AI will analyze. If it looks wrong or incomplete, the PDF may need OCR processing.")

    # Confirm extraction
    if not st.button("✅ Proceed with Extraction", type="primary", key="confirm_extract"):
        st.info("Review the text above, then click 'Proceed with Extraction'")
        return

    with st.spinner(f"🔄 Sending to AI for extraction..."):
        try:
            # Get the schema prompt with hints
            prompt_template = get_schema_prompt()

            # Show the full prompt being sent
            full_prompt = prompt_template.format(
                text=combined_text, page_num=page_num + 1)

            with st.expander("🔍 Debug: Full prompt sent to AI", expanded=False):
                st.code(full_prompt, language="text")

            # Create extractor and call LLM
            extractor = LLMExtractor(custom_prompt=prompt_template)

            # Extract products from the combined text
            products = extractor.extract_products_from_text(
                combined_text,
                page_num,
                [b.get("block_id", f"block_{i}") for i, b in enumerate(
                    page.blocks) if b.get("text")]
            )

            if not products:
                st.error("❌ No products found!")
                st.warning("""
                **The AI returned no products. Possible reasons:**
                1. The text doesn't contain recognizable product data
                2. The schema hints don't match the catalog format
                3. The page layout description is missing or incorrect
                
                **Try:**
                - Go to Schema view and add better extraction hints
                - Describe your page layout (e.g., "Style name at top applies to all products")
                - Check if the raw text above actually contains product information
                """)
                return

            # Add products to session state
            added_count = 0
            for product in products:
                product.product_id = f"p{page_num}_{len(st.session_state.products)}"
                product.page_number = page_num
                product.needs_review = True
                st.session_state.products.append(product)
                added_count += 1

            st.success(
                f"✅ Extracted {added_count} products from page {page_num + 1}!")
            st.rerun()

        except Exception as e:
            import traceback
            st.error(f"❌ Extraction failed: {str(e)}")
            with st.expander("Error details"):
                st.code(traceback.format_exc(), language="python")


def render_edit_view():
    """Render the product edit view using schema fields"""
    product = st.session_state.selected_product

    if not product:
        st.warning("No product selected for editing")
        if st.button("← Back to Review"):
            st.session_state.view_mode = 'review'
            st.rerun()
        return

    st.subheader(f"✏️ Edit Product")

    # Back button
    if st.button("← Back to Review"):
        st.session_state.view_mode = 'review'
        st.session_state.selected_product = None
        st.rerun()

    st.divider()

    # Get schema fields
    schema_fields = st.session_state.get(
        'schema_fields', DEFAULT_SCHEMA_FIELDS)

    # Store new values
    new_values = {}

    # Create two columns for the form
    col1, col2 = st.columns(2)

    # Split fields between columns
    mid = (len(schema_fields) + 1) // 2

    for i, field in enumerate(schema_fields):
        field_name = field['name']
        field_type = field.get('type', 'text')

        # Get current value
        if hasattr(product, field_name):
            current_value = getattr(product, field_name)
        else:
            current_value = product.raw_attributes.get(field_name, '')

        # Handle price specially
        if field_name == 'price' and current_value:
            if hasattr(current_value, 'amount'):
                current_value = current_value.amount
            else:
                # Try to extract number from string
                try:
                    current_value = float(str(current_value).replace(
                        '$', '').replace('€', '').replace(',', '').strip())
                except:
                    current_value = 0.0

        # Choose column
        with col1 if i < mid else col2:
            display_name = field_name.replace('_', ' ').title()

            if field_type == 'number':
                new_values[field_name] = st.number_input(
                    display_name,
                    value=float(current_value) if current_value else 0.0,
                    step=0.01,
                    key=f"edit_{field_name}"
                )
            else:
                new_values[field_name] = st.text_input(
                    display_name,
                    value=str(current_value) if current_value else "",
                    key=f"edit_{field_name}"
                )

    # Show any extra raw_attributes not in schema
    extra_attrs = {k: v for k, v in product.raw_attributes.items()
                   if k not in [f['name'] for f in schema_fields] and v}
    if extra_attrs:
        st.markdown("### Additional Extracted Data")
        for key, value in extra_attrs.items():
            display_name = key.replace('_', ' ').title()
            new_values[f"extra_{key}"] = st.text_input(
                display_name,
                value=str(value) if value else "",
                key=f"edit_extra_{key}"
            )

    # Raw text reference
    if product.raw_text:
        with st.expander("📝 Original Raw Text"):
            st.text(product.raw_text)

    st.divider()

    col_save, col_cancel = st.columns(2)

    with col_save:
        if st.button("💾 Save Changes", type="primary", width="stretch"):
            # Update product with new values
            for field in schema_fields:
                field_name = field['name']
                new_val = new_values.get(field_name)

                # Convert empty strings to None
                if new_val == "" or new_val == 0.0:
                    new_val = None

                if hasattr(product, field_name):
                    setattr(product, field_name, new_val)
                else:
                    product.raw_attributes[field_name] = new_val

            # Update extra attributes
            for key, value in extra_attrs.items():
                new_val = new_values.get(f"extra_{key}")
                product.raw_attributes[key] = new_val if new_val else None

            product.needs_review = False

            st.success("Product updated!")
            st.session_state.view_mode = 'review'
            st.session_state.selected_product = None
            st.rerun()

    with col_cancel:
        if st.button("❌ Cancel", width="stretch"):
            st.session_state.view_mode = 'review'
            st.session_state.selected_product = None
            st.rerun()


def delete_schema_field(field_name: str):
    """Callback to delete a schema field by name"""
    st.session_state.schema_fields = [
        f for f in st.session_state.schema_fields
        if f['name'] != field_name
    ]


def update_schema_field(old_name: str, idx: int):
    """Callback to update a schema field from input widgets"""
    # Get values from session state widgets
    new_name = st.session_state.get(
        f"edit_name_{old_name}", old_name).lower().replace(" ", "_")
    new_type = st.session_state.get(f"edit_type_{old_name}", "text")
    new_desc = st.session_state.get(f"edit_desc_{old_name}", "")
    new_req = st.session_state.get(f"edit_req_{old_name}", False)

    # Preserve existing hint
    current_hint = ""
    for f in st.session_state.schema_fields:
        if f['name'] == old_name:
            current_hint = f.get('hint', '')
            break

    # Check for duplicate names (excluding self and computed fields)
    other_names = [f['name'] for i, f in enumerate(st.session_state.schema_fields) if i != idx]
    computed_names = [f['name'] for f in st.session_state.get('computed_fields', [])]
    
    if new_name in other_names or new_name in computed_names:
        st.session_state.edit_error = f"Field name '{new_name}' already exists!"
        return

    # Update the field (preserve hint)
    st.session_state.schema_fields[idx] = {
        "name": new_name,
        "type": new_type,
        "description": new_desc,
        "required": new_req,
        "hint": current_hint
    }
    st.session_state.edit_error = None


def save_field_hint(field_name: str):
    """Save the extraction hint for a field"""
    hint_value = st.session_state.get(f"hint_text_{field_name}", "")
    for field in st.session_state.schema_fields:
        if field['name'] == field_name:
            field['hint'] = hint_value
            break
    st.session_state.editing_field_hint = None


def render_schema_config():
    """Render schema configuration view with extraction hints"""
    # Determine current mode
    is_excel_mode = st.session_state.app_mode == 'excel'
    mode_label = "Excel" if is_excel_mode else "PDF"

    st.subheader("📋 Configure Extraction Schema")
    st.caption(
        f"Define fields and add extraction hints to guide the AI on how to find data in your {mode_label.lower()} catalog")

    # Show any edit errors
    if st.session_state.get('edit_error'):
        st.error(st.session_state.edit_error)
        st.session_state.edit_error = None

    # ===== PAGE/SHEET LAYOUT CONTEXT =====
    if is_excel_mode:
        st.markdown("### 📊 Sheet Layout Description")
        st.caption(
            "Describe the overall structure of your Excel sheet to help the AI understand the data layout")
        layout_placeholder = """Example:
- Each row represents one product
- Column A contains product images
- Column B has SKU codes
- Prices are in Column E (wholesale) and Column F (retail)
- Some products span multiple rows (grouped by style)
- Header row is row 1, data starts at row 2"""
    else:
        st.markdown("### 📄 Page Layout Description")
        st.caption(
            "Describe the overall structure of your catalog pages to help the AI understand the layout")
        layout_placeholder = """Example:
- Each page shows multiple products in a grid layout
- The STYLE name appears at the top of the page and applies to ALL products on that page
- Products are arranged in rows with image on left, details on right
- Prices are shown as 'Wholesale: $X.XX' and 'Retail: $X.XX'
- Color codes are abbreviated (BLK=Black, WHT=White, etc.)"""

    page_context = st.text_area(
        "Layout Context",
        value=st.session_state.get('page_layout_context', ''),
        height=120,
        placeholder=layout_placeholder,
        key="page_layout_input",
        label_visibility="collapsed"
    )

    if page_context != st.session_state.get('page_layout_context', ''):
        st.session_state.page_layout_context = page_context

    # ===== IMAGE POSITION HINT (PDF only) =====
    if not is_excel_mode:
        st.markdown("### 🖼️ Product Image Position")
        st.caption(
            "Describe where product images are typically located relative to product details")

        image_hint = st.text_area(
            "Image Position Hint",
            value=st.session_state.get('image_position_hint', ''),
            height=80,
            placeholder="""Examples:
- Product image is to the LEFT of the product name and details
- Images are ABOVE the product description
- Each product has a small thumbnail in the top-left corner of its section""",
            key="image_position_input",
            label_visibility="collapsed"
        )

        if image_hint != st.session_state.get('image_position_hint', ''):
            st.session_state.image_position_hint = image_hint

    st.divider()

    # ===== SCHEMA FIELDS WITH HINTS =====
    st.markdown("### 🏷️ Extraction Fields")
    st.caption(
        "For each field, add an extraction hint explaining WHERE and HOW to find the data")

    for idx, field in enumerate(st.session_state.schema_fields):
        field_name = field['name']

        with st.container():
            # Field header row
            col1, col2, col3, col4, col5 = st.columns([2, 1, 2, 0.8, 0.5])

            with col1:
                st.text_input(
                    "Name",
                    value=field_name,
                    key=f"edit_name_{field_name}",
                    on_change=update_schema_field,
                    args=(field_name, idx),
                    label_visibility="collapsed",
                    placeholder="Field name"
                )
            with col2:
                st.selectbox(
                    "Type",
                    ["text", "number", "boolean", "list"],
                    index=["text", "number", "boolean", "list"].index(
                        field.get('type', 'text')),
                    key=f"edit_type_{field_name}",
                    on_change=update_schema_field,
                    args=(field_name, idx),
                    label_visibility="collapsed"
                )
            with col3:
                st.text_input(
                    "Description",
                    value=field.get('description', ''),
                    key=f"edit_desc_{field_name}",
                    on_change=update_schema_field,
                    args=(field_name, idx),
                    label_visibility="collapsed",
                    placeholder="What is this field"
                )
            with col4:
                st.checkbox(
                    "Req",
                    value=field.get('required', False),
                    key=f"edit_req_{field_name}",
                    on_change=update_schema_field,
                    args=(field_name, idx),
                    help="Required field"
                )
            with col5:
                st.button(
                    "🗑️",
                    key=f"remove_field_{field_name}",
                    on_click=delete_schema_field,
                    args=(field_name,)
                )

            # Extraction hint row
            hint_value = field.get('hint', '')
            hint_placeholder = f"How to find '{field_name}' in the catalog..."

            if is_excel_mode:
                # Excel-specific placeholders
                if field_name == 'name':
                    hint_placeholder = "e.g., 'Product name is in column B or C'"
                elif field_name == 'price':
                    hint_placeholder = "e.g., 'Wholesale price is in column E, retail in column F'"
                elif field_name == 'sku':
                    hint_placeholder = "e.g., 'SKU is in column A, format: XXX-YYYY'"
                elif field_name == 'color':
                    hint_placeholder = "e.g., 'Color is in column D, may use abbreviations like BLK=Black'"
            else:
                # PDF-specific placeholders
                if field_name == 'name':
                    hint_placeholder = "e.g., 'Product name is the large bold text above the image'"
                elif field_name == 'price':
                    hint_placeholder = "e.g., 'Look for Wholesale price, ignore Retail price. Format: $XX.XX'"
                elif field_name == 'sku':
                    hint_placeholder = "e.g., 'SKU is the alphanumeric code starting with the brand prefix'"
                elif field_name == 'color':
                    hint_placeholder = "e.g., 'Color appears above wholesale. Use full name, not abbreviation (BLK=Black)'"

            new_hint = st.text_area(
                f"🔍 Extraction Hint for '{field_name}'",
                value=hint_value,
                height=68,
                placeholder=hint_placeholder,
                key=f"hint_text_{field_name}",
                label_visibility="collapsed"
            )

            # Auto-save hint when changed
            if new_hint != hint_value:
                field['hint'] = new_hint

            st.markdown("---")

    # Header legend removed per user request

    st.divider()

    # ===== COMPUTED FIELDS =====
    st.markdown("### ∑ Computed Formula Fields")
    st.info("""
    **Add fields calculated from extracted data using Excel-like formulas.**
    
    **Supported Functions:** `ROUNDUP(n, d)`, `ROUNDDOWN(n, d)`, `IF(cond, true, false)`, `SUM(...)`, `AVERAGE(...)`  
    **Syntax:** Use `{field_name}` to reference other fields.  
    **Example:** `ROUNDUP({price} * 1.2, 0)` adds 20% markup rounded up to integer.
    """)

    computed_fields = st.session_state.get('computed_fields', [])
    
    # Display existing computed fields
    if computed_fields:
        for idx, field in enumerate(computed_fields):
            with st.container():
                col1, col2, col3 = st.columns([2, 4, 0.5])
                with col1:
                    st.text_input(
                        "Name",
                        value=field['name'],
                        key=f"comp_name_{idx}",
                        disabled=True,
                        label_visibility="collapsed"
                    )
                with col2:
                    new_formula = st.text_input(
                        "Formula",
                        value=field['formula'],
                        key=f"comp_formula_{idx}",
                        label_visibility="collapsed",
                        placeholder="e.g. {price} * 1.2"
                    )
                    if new_formula != field['formula']:
                        field['formula'] = new_formula
                with col3:
                    if st.button("🗑️", key=f"del_comp_{idx}", help="Delete computed field"):
                        st.session_state.computed_fields.pop(idx)
                        st.rerun()

    # Add new computed field
    with st.container():
        st.write("Add New Computed Field")
        col1, col2, col3 = st.columns([2, 4, 1])
        with col1:
            new_comp_name = st.text_input(
                "Name", key="new_comp_name", placeholder="total_price", label_visibility="collapsed")
        with col2:
            new_comp_formula = st.text_input(
                "Formula", key="new_comp_formula", placeholder="{price} * {quantity}", label_visibility="collapsed")
        with col3:
            if st.button("➕ Add", key="add_comp_btn"):
                if new_comp_name and new_comp_formula:
                    # simplistic validation
                    clean_name = new_comp_name.lower().replace(" ", "_")
                    
                    # Check duplicates
                    all_names = [f['name'] for f in st.session_state.schema_fields] + \
                                [f['name'] for f in computed_fields]
                    
                    if clean_name in all_names:
                        st.error(f"Field '{clean_name}' already exists")
                    else:
                        st.session_state.computed_fields.append({
                            "name": clean_name,
                            "formula": new_comp_formula
                        })
                        st.rerun()
                else:
                    st.warning("Name and Formula required")

    st.divider()

    # ===== ADD NEW FIELD =====
    st.markdown("### ➕ Add New Field")

    col1, col2, col3, col4 = st.columns([2, 1, 3, 1])

    with col1:
        new_field_name = st.text_input(
            "Field Name", key="new_field_name", placeholder="e.g., size, color, weight")
    with col2:
        new_field_type = st.selectbox(
            "Type", ["text", "number", "boolean", "list"], key="new_field_type")
    with col3:
        new_field_desc = st.text_input(
            "Description", key="new_field_desc", placeholder="What this field contains")
    with col4:
        new_field_req = st.checkbox("Required", key="new_field_req")

    if st.button("➕ Add Field", width="stretch"):
        if new_field_name:
            existing_names = [f['name']
                              for f in st.session_state.schema_fields]
            if new_field_name.lower() in [n.lower() for n in existing_names]:
                st.error("Field already exists!")
            else:
                st.session_state.schema_fields.append({
                    "name": new_field_name.lower().replace(" ", "_"),
                    "type": new_field_type,
                    "required": new_field_req,
                    "description": new_field_desc,
                    "hint": "",
                    "formula": ""
                })
                st.success(f"Added field: {new_field_name}")
                st.rerun()
        else:
            st.warning("Enter a field name")

    st.divider()

    # Quick add common fields
    st.markdown("### Quick Add Common Fields")

    common_fields = [
        {"name": "size", "type": "text", "description": "Product size", "hint": ""},
        {"name": "color", "type": "text", "description": "Product color", "hint": ""},
        {"name": "weight", "type": "text",
            "description": "Product weight", "hint": ""},
        {"name": "dimensions", "type": "text",
            "description": "Product dimensions", "hint": ""},
        {"name": "material", "type": "text",
            "description": "Product material", "hint": ""},
        {"name": "quantity", "type": "number",
            "description": "Available quantity", "hint": ""},
        {"name": "unit_price", "type": "number",
            "description": "Price per unit", "hint": ""},
        {"name": "ean", "type": "text", "description": "EAN/Barcode", "hint": ""},
        {"name": "manufacturer", "type": "text",
            "description": "Manufacturer name", "hint": ""},
        {"name": "model", "type": "text", "description": "Model number", "hint": ""},
        {"name": "style", "type": "text",
            "description": "Style name/number", "hint": ""},
        {"name": "wholesale_price", "type": "number",
            "description": "Wholesale price", "hint": ""},
        {"name": "retail_price", "type": "number",
            "description": "Retail/MSRP price", "hint": ""},
        {"name": "upc", "type": "text", "description": "UPC code", "hint": ""},
    ]

    existing_names = [f['name'].lower()
                      for f in st.session_state.schema_fields]
    available_fields = [
        f for f in common_fields if f['name'].lower() not in existing_names]

    if available_fields:
        cols = st.columns(5)
        for idx, field in enumerate(available_fields[:15]):
            with cols[idx % 5]:
                if st.button(f"+ {field['name']}", key=f"quick_add_{field['name']}"):
                    st.session_state.schema_fields.append({
                        **field,
                        "required": False,
                        "hint": "",
                        "formula": ""
                    })
                    st.rerun()
    else:
        st.info("All common fields already added")

    st.divider()

    # Reset to defaults
    if st.button("🔄 Reset to Defaults"):
        st.session_state.schema_fields = [
            f.copy() for f in DEFAULT_SCHEMA_FIELDS]
        st.session_state.page_layout_context = DEFAULT_PAGE_CONTEXT
        st.rerun()

    # Preview prompt
    st.divider()
    st.markdown("### 👁️ Preview Extraction Prompt")
    if st.checkbox("Show prompt that will be sent to AI", value=False):
        if is_excel_mode:
            prompt = get_excel_schema_prompt()
            st.code(prompt.format(
                headers="[COLUMN HEADERS]", sample_rows="[SAMPLE DATA ROWS]"), language="text")
        else:
            prompt = get_schema_prompt()
            st.code(prompt.format(
                text="[YOUR CATALOG TEXT HERE]", page_num=1), language="text")


def get_schema_prompt():
    """Generate extraction prompt based on current schema with hints"""
    fields = st.session_state.get('schema_fields', None)
    if not fields:
        fields = DEFAULT_SCHEMA_FIELDS.copy()

    page_context = st.session_state.get('page_layout_context', '')
    image_hint = st.session_state.get('image_position_hint', '')

    # Build field descriptions with hints
    field_lines = []
    for field in fields:
        try:
            field_name = field.get('name', 'unknown')
            req = "REQUIRED" if field.get('required') else "optional"
            field_type = field.get('type', 'text')
            desc = field.get('description', '')
            hint = field.get('hint', '')

            field_line = f"  • {field_name} ({req}, {field_type}): {desc}"
            if hint:
                field_line += f"\n      → EXTRACTION HINT: {hint}"
            field_lines.append(field_line)
        except Exception:
            continue

    if not field_lines:
        field_lines = ["  • name (REQUIRED, text): Product name"]

    fields_str = "\n".join(field_lines)

    # Build the prompt
    prompt_parts = [
        "You are a product data extraction expert analyzing a product catalog.",
        ""
    ]

    # Add page layout context if provided
    if page_context.strip():
        prompt_parts.append("=== CATALOG LAYOUT DESCRIPTION ===")
        prompt_parts.append(page_context.strip())
        prompt_parts.append("")

    prompt_parts.extend([
        "=== FIELDS TO EXTRACT ===",
        fields_str,
        "",
        "=== EXTRACTION RULES ===",
        "1. Extract EVERY product mentioned - do not skip any product for any reason",
        "2. If a field value is shared across multiple products (e.g., style at top of page), apply it to ALL relevant products",
        "3. If a field is not found for a product, omit it (do not guess or make up values)",
        "4. Preserve exact text for prices, SKUs, codes as they appear",
        "5. For abbreviations, expand them if the hint specifies to do so",
        "6. Multiple text blocks may describe the SAME product - combine them intelligently",
        "7. Return ONLY valid JSON, no explanations",
        "8. IMPORTANT: Products without images are still valid - extract them anyway",
        ""
    ])

    # Add image extraction instructions - minimal
    prompt_parts.append("=== IMAGE EXTRACTION ===")
    prompt_parts.append(
        "For products with visible images, include: \"image_bbox\": [x1, y1, x2, y2]")
    prompt_parts.append(
        "Values are percentages (0.0-1.0) of page dimensions. x1,y1=top-left, x2,y2=bottom-right.")
    if image_hint.strip():
        prompt_parts.append(f"Hint: {image_hint.strip()}")
    prompt_parts.append("")

    prompt_parts.extend([
        "=== CATALOG TEXT (Page {{page_num}}) ===",
        "{{text}}",
        "",
        "=== RESPONSE FORMAT ===",
        "Respond with a JSON object: {{{{\"products\": [{{{{...}}}}, {{{{...}}}}]}}}}",
        "Each product should have at minimum a \"name\" field."
    ])

    return "\n".join(prompt_parts)


def get_excel_schema_prompt():
    """Generate Excel-specific extraction prompt based on current schema with hints"""
    fields = st.session_state.get('schema_fields', None)
    if not fields:
        fields = DEFAULT_SCHEMA_FIELDS.copy()

    layout_context = st.session_state.get('page_layout_context', '')

    # Build field descriptions with hints
    field_lines = []
    for field in fields:
        try:
            field_name = field.get('name', 'unknown')
            req = "REQUIRED" if field.get('required') else "optional"
            field_type = field.get('type', 'text')
            desc = field.get('description', '')
            hint = field.get('hint', '')

            field_line = f"  • {field_name} ({req}, {field_type}): {desc}"
            if hint:
                field_line += f"\n      → EXTRACTION HINT: {hint}"
            field_lines.append(field_line)
        except Exception:
            continue

    if not field_lines:
        field_lines = ["  • name (REQUIRED, text): Product name"]

    fields_str = "\n".join(field_lines)

    # Build the prompt
    prompt_parts = [
        "You are a product data extraction expert analyzing an Excel spreadsheet.",
        ""
    ]

    # Add sheet layout context if provided
    if layout_context.strip():
        prompt_parts.append("=== SHEET LAYOUT DESCRIPTION ===")
        prompt_parts.append(layout_context.strip())
        prompt_parts.append("")

    prompt_parts.extend([
        "=== FIELDS TO EXTRACT ===",
        fields_str,
        "",
        "=== COLUMN HEADERS ===",
        "{headers}",
        "",
        "=== SAMPLE DATA ROWS ===",
        "{sample_rows}",
        "",
        "=== EXTRACTION RULES ===",
        "1. Analyze the column headers and sample data to understand the mapping",
        "2. Determine which columns map to which extraction fields",
        "3. Identify any row grouping patterns (one product per row vs multiple rows per product)",
        "4. Note any transformations needed (e.g., combining columns, parsing values)",
        "5. Return a column mapping that can be applied to all rows",
        "",
        "=== RESPONSE FORMAT ===",
        "Return a JSON object with the learned column mappings."
    ])

    return "\n".join(prompt_parts)


def main():
    """Main application entry point"""
    init_session_state()

    # Render sidebar (contains mode selector)
    render_sidebar()

    # Main content area based on mode
    if st.session_state.app_mode == 'pdf':
        # PDF Extractor views
        if st.session_state.view_mode == 'extraction':
            render_extraction_view()
        elif st.session_state.view_mode == 'review':
            render_review_view()
        elif st.session_state.view_mode == 'export':
            render_export_view()
        elif st.session_state.view_mode == 'edit':
            render_edit_view()
        elif st.session_state.view_mode == 'schema':
            render_schema_config()
        else:
            render_extraction_view()
    else:
        # Excel Extractor views
        render_excel_main_content()


def render_excel_main_content():
    """Render the Excel extractor main content area"""
    if st.session_state.excel_view_mode == 'extraction':
        render_excel_extraction_view()
    elif st.session_state.excel_view_mode == 'review':
        render_excel_review_view()
    elif st.session_state.excel_view_mode == 'export':
        render_excel_export_view()
    elif st.session_state.excel_view_mode == 'schema':
        render_schema_config()  # Shared schema config
    else:
        render_excel_extraction_view()


def render_excel_sidebar_content():
    """Render Excel-specific sidebar content"""
    with st.sidebar:
        # File upload
        st.subheader("📊 Upload Excel File")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload a product catalog Excel file",
            key="excel_uploader"
        )

        if uploaded_file:
            # Header row selector — always available so user can change and reload
            st.markdown("##### 📋 Header Row")
            header_options = ["Auto-detect", "No header", "Specific row"]
            current_hr = st.session_state.get('excel_header_row')
            if current_hr is None:
                default_idx = 0
            elif current_hr == 0:
                default_idx = 1
            else:
                default_idx = 2

            header_choice = st.radio(
                "Header row mode",
                header_options,
                index=default_idx,
                key="excel_header_mode",
                label_visibility="collapsed",
                help="Auto-detect uses the first non-empty row as headers. 'No header' generates Column_1, Column_2, etc.",
                horizontal=True
            )

            new_header_row = None
            if header_choice == "No header":
                new_header_row = 0
            elif header_choice == "Specific row":
                new_header_row = st.number_input(
                    "Header is on row",
                    min_value=1,
                    max_value=999,
                    value=max(1, current_hr or 1),
                    key="excel_header_row_num"
                )

            # Update session state (will take effect on next Load)
            if new_header_row != st.session_state.get('excel_header_row'):
                st.session_state.excel_header_row = new_header_row

            # Load/Reload button
            if st.button("🔄 Load/Reload Excel", width='stretch'):
                import hashlib
                file_bytes = uploaded_file.getvalue()
                new_hash = hashlib.md5(file_bytes).hexdigest()

                # Reset Excel state (keep header_row — user's choice should persist)
                st.session_state.excel_file_hash = new_hash
                st.session_state.excel_products = []
                st.session_state.excel_current_sheet = None
                st.session_state.excel_preview_cache = None
                st.session_state.excel_preview_cache_key = None
                st.session_state.excel_doc = None  # Force full re-processing
                st.session_state.excel_start_row = None
                st.session_state.excel_end_row = None

                # Load the Excel file
                with st.spinner("Loading Excel..."):
                    try:
                        processor = ExcelProcessor()
                        st.session_state.excel_doc = processor.load_excel_from_bytes(
                            file_bytes, uploaded_file.name,
                            header_row=st.session_state.get('excel_header_row')
                        )
                        # Set default sheet
                        if st.session_state.excel_doc.sheet_names:
                            st.session_state.excel_current_sheet = st.session_state.excel_doc.sheet_names[
                                0]
                            sheet = st.session_state.excel_doc.get_sheet(
                                st.session_state.excel_current_sheet)
                            if sheet:
                                st.session_state.excel_end_row = sheet.data_start_excel_row + sheet.total_rows - 1
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to load Excel: {e}")
                        st.session_state.excel_doc = None

            if st.session_state.excel_doc:
                st.caption(f"✅ {st.session_state.excel_doc.file_name}")

        if st.session_state.excel_doc:
            st.divider()

            # Sheet selector
            st.subheader("📃 Sheet")
            sheet_name = st.selectbox(
                "Select sheet",
                st.session_state.excel_doc.sheet_names,
                index=st.session_state.excel_doc.sheet_names.index(
                    st.session_state.excel_current_sheet
                ) if st.session_state.excel_current_sheet else 0,
                key="excel_sheet_select",
                label_visibility="collapsed"
            )
            if sheet_name != st.session_state.excel_current_sheet:
                st.session_state.excel_current_sheet = sheet_name
                st.session_state.excel_preview_cache = None
                st.session_state.excel_preview_cache_key = None

            # Show sheet info
            sheet = st.session_state.excel_doc.get_sheet(sheet_name)
            if sheet:
                st.caption(
                    f"📈 {sheet.total_rows} rows, {len(sheet.headers)} cols")

            st.divider()

            # Row range (uses Excel-absolute row numbers)
            if sheet:
                st.subheader("⚙️ Row Range")
                excel_first = sheet.data_start_excel_row
                excel_last = sheet.data_start_excel_row + sheet.total_rows - 1
                # Ensure session state values are valid before widget renders
                if not st.session_state.excel_start_row or st.session_state.excel_start_row < excel_first or st.session_state.excel_start_row > excel_last:
                    st.session_state.excel_start_row = excel_first
                if not st.session_state.excel_end_row or st.session_state.excel_end_row < excel_first or st.session_state.excel_end_row > excel_last:
                    st.session_state.excel_end_row = excel_last

                col1, col2 = st.columns(2)
                with col1:
                    st.number_input(
                        "From row",
                        min_value=excel_first,
                        max_value=excel_last,
                        key="excel_start_row"
                    )
                with col2:
                    st.number_input(
                        "To row",
                        min_value=excel_first,
                        max_value=excel_last,
                        key="excel_end_row"
                    )

            st.divider()

            # View mode - vertical buttons (matching PDF style)
            st.subheader("👁️ View Mode")

            if st.button("📄 Extraction", width='stretch',
                         type="primary" if st.session_state.excel_view_mode == 'extraction' else "secondary",
                         key="excel_view_extraction"):
                st.session_state.excel_view_mode = 'extraction'
                st.rerun()

            if st.button("🔍 Review", width='stretch',
                         type="primary" if st.session_state.excel_view_mode == 'review' else "secondary",
                         key="excel_view_review"):
                st.session_state.excel_view_mode = 'review'
                st.rerun()

            if st.button("📤 Export", width='stretch',
                         type="primary" if st.session_state.excel_view_mode == 'export' else "secondary",
                         key="excel_view_export"):
                st.session_state.excel_view_mode = 'export'
                st.rerun()

            if st.button("📋 Schema", width='stretch',
                         type="primary" if st.session_state.excel_view_mode == 'schema' else "secondary",
                         key="excel_view_schema"):
                st.session_state.excel_view_mode = 'schema'
                st.rerun()

            st.divider()

            # Stats
            st.subheader("📊 Stats")
            st.metric("Products", len(st.session_state.excel_products))
        else:
            # Show schema config even without Excel file loaded
            st.divider()
            if st.button("📋 Configure Schema", key="excel_configure_schema", width='stretch'):
                st.session_state.excel_view_mode = 'schema'


def render_excel_extraction_view():
    """Render the Excel extraction view"""
    if not st.session_state.excel_doc:
        st.info("👆 Upload an Excel file and click **Load/Reload Excel** to get started")
        return

    sheet_name = st.session_state.excel_current_sheet
    sheet = st.session_state.excel_doc.get_sheet(sheet_name)

    if not sheet:
        st.warning("No sheet selected")
        return

    st.subheader(f"📊 Sheet: {sheet_name}")
    num_images = getattr(sheet, 'image_count', 0) or len(
        getattr(sheet, 'images', []))
    st.caption(
        f"{sheet.total_rows} rows | {len(sheet.headers)} cols | 🖼️ {num_images} images")

    # Preview data with caching and dark mode
    with st.expander("👁️ Preview Data", expanded=True):
        import pandas as pd

        # Adjustable preview size
        preview_size = st.slider(
            "Preview rows",
            min_value=10,
            max_value=min(200, sheet.total_rows),
            value=min(50, sheet.total_rows),
            step=10,
            key="excel_preview_size"
        )

        # Cache key based on sheet + size
        cache_key = f"{sheet_name}_{preview_size}_{sheet.total_rows}"

        # Use cached DataFrame if available
        if st.session_state.excel_preview_cache_key != cache_key:
            preview_data = sheet.get_preview(preview_size)
            if preview_data:
                df = pd.DataFrame(preview_data, columns=sheet.headers)
                # Index = actual Excel row numbers (data starts after headers)
                df.index = range(sheet.data_start_excel_row, sheet.data_start_excel_row + len(df))
                df = df.astype(str)
                st.session_state.excel_preview_cache = df
                st.session_state.excel_preview_cache_key = cache_key
            else:
                st.session_state.excel_preview_cache = None

        df = st.session_state.excel_preview_cache
        if df is not None:
            # Dark mode styling
            st.markdown("""
            <style>
            .excel-preview .stDataFrame [data-testid="stDataFrameResizable"] {
                background-color: #1e1e1e;
            }
            div[data-testid="stDataFrame"] > div {
                background-color: #1e1e1e;
            }
            </style>
            """, unsafe_allow_html=True)
            st.dataframe(
                df,
                width='stretch',
                height=400,
            )
        else:
            st.info("No data to preview")

    # Show images if available (lazy-loaded, paginated)
    if num_images > 0:
        with st.expander(f"🖼️ Images ({num_images} found)", expanded=False):
            st.info(
                f"Found {num_images} embedded images. Images are loaded on demand.")

            image_store = getattr(
                st.session_state.excel_doc, 'image_store', None)
            image_metas = getattr(sheet, 'image_metas', [])

            # Row-based search - images only load when searching
            row_search = st.text_input(
                "🔍 Search by row to load images",
                placeholder="Enter row number or range (e.g., '5' or '10-20')",
                key="excel_image_row_search"
            )

            # Only load images when user searches for specific rows
            if not row_search:
                st.info(
                    "💡 Enter a row number or range above to load images. This improves app performance.")
            else:
                # Filter metas based on search
                filtered_metas = []
                try:
                    if '-' in row_search:
                        start, end = map(int, row_search.split('-'))
                        filtered_metas = [
                            m for m in image_metas if start <= m.row + 1 <= end]
                    else:
                        target_row = int(row_search)
                        filtered_metas = [
                            m for m in image_metas if m.row + 1 == target_row]

                    if filtered_metas:
                        st.success(
                            f"Found {len(filtered_metas)} images matching row(s) {row_search}")
                    else:
                        st.warning(f"No images found for row(s) {row_search}")
                except ValueError:
                    st.warning(
                        "Invalid format. Use a number or range (e.g., '10-20')")

                # Paginate — show max 20 at a time
                PAGE_SIZE = 20
                total_filtered = len(filtered_metas)

                if total_filtered > PAGE_SIZE:
                    max_page = (total_filtered - 1) // PAGE_SIZE
                    img_page = st.number_input(
                        f"Page (1-{max_page + 1})",
                        min_value=1,
                        max_value=max_page + 1,
                        value=1,
                        key="excel_image_page"
                    ) - 1
                    page_start = img_page * PAGE_SIZE
                    page_metas = filtered_metas[page_start:page_start + PAGE_SIZE]
                    st.caption(
                        f"Showing {page_start + 1}-{min(page_start + PAGE_SIZE, total_filtered)} of {total_filtered}")
                else:
                    page_metas = filtered_metas

                # Render only the visible page of images
                if image_store and page_metas:
                    cols_per_row = 4
                    for i in range(0, len(page_metas), cols_per_row):
                        cols = st.columns(cols_per_row)
                        for j, col in enumerate(cols):
                            if i + j < len(page_metas):
                                meta = page_metas[i + j]
                                with col:
                                    try:
                                        img_bytes = image_store.get_image_bytes(
                                            meta.sheet_name, meta.index)
                                        if img_bytes:
                                            st.image(
                                                img_bytes, caption=f"Row {meta.row + 1}")
                                        else:
                                            st.caption(
                                                f"Row {meta.row + 1}: ⚠️")
                                    except Exception:
                                        st.caption(
                                            f"Row {meta.row + 1}: Error")
                elif not image_store:
                    st.warning(
                        "Image store not available (file may need to be reloaded)")

    st.divider()

    # Extraction controls
    start_row = st.session_state.excel_start_row or sheet.data_start_excel_row
    end_row = st.session_state.excel_end_row or (sheet.data_start_excel_row + sheet.total_rows - 1)

    st.checkbox(
        "🖼️ Include Images",
        value=st.session_state.get('excel_include_images', True),
        help="Include product images (may slow down extraction)",
        key="excel_include_images"
    )

    st.info("🧠 **LLM mode** sends rows in batches to the AI. Slower but fully flexible — "
            "hints like *\"apply discount to price\"*, *\"merge multi-row products\"*, etc. all work. "
            "Configure hints in **Schema** settings.")

    # LLM options
    with st.expander("⚙️ LLM Options", expanded=False):
        st.number_input(
            "Rows per batch",
            min_value=1,
            max_value=50,
            value=st.session_state.get('excel_llm_batch_size', 20),
            step=1,
            help="Number of rows sent to the AI per batch. Smaller = more accurate, larger = faster. For multi-row products, use a multiple of row count per product.",
            key="_llm_batch_size_widget",
            on_change=lambda: setattr(st.session_state, 'excel_llm_batch_size', st.session_state._llm_batch_size_widget)
        )

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button(
            f"🧠 Extract rows {start_row}-{end_row} with LLM",
            type="primary",
            key="llm_extract_btn",
            help=f"Send {end_row - start_row + 1} rows in batches of {st.session_state.excel_llm_batch_size} to the AI"
        ):
            do_llm_excel_extraction(
                sheet, start_row, end_row,
                st.session_state.excel_llm_batch_size,
                st.session_state.excel_include_images
            )

    with col_btn2:
        if st.button("🗑️ Clear Products", key="clear_excel_products_llm"):
            st.session_state.excel_products = []
            st.rerun()

    st.divider()

    # Show extracted products
    st.subheader(
        f"📦 Extracted Products ({len(st.session_state.excel_products)})")

    if not st.session_state.excel_products:
        st.info("No products extracted yet. Click **Extract Products** to start.")
    else:
        # Show first 50
        for idx, product in enumerate(st.session_state.excel_products[:50]):
            render_excel_product_card(product, idx)

        if len(st.session_state.excel_products) > 50:
            st.caption(
                f"Showing first 50 of {len(st.session_state.excel_products)} products.")


def render_excel_product_card(product: Product, idx: int):
    """Render a product card for Excel extraction view"""
    review_icon = "⚠️" if product.needs_review else "✅"

    # Show row info in title if available
    source_rows = getattr(product, 'source_rows', [])
    row_info = f" (Rows: {', '.join(map(str, source_rows))})" if source_rows else ""

    with st.expander(f"{review_icon} {product.name[:50]}{'...' if len(product.name) > 50 else ''}{row_info}", expanded=False):
        # Check if product has an image (but don't load it yet)
        excel_image = getattr(product, 'excel_image', None)
        has_image = excel_image is not None

        if has_image:
            col_img, col_fields = st.columns([1, 3])
            with col_img:
                # Lazy load image only when button clicked
                show_image_key = f"show_img_{idx}"
                if show_image_key not in st.session_state:
                    st.session_state[show_image_key] = False

                if st.session_state[show_image_key]:
                    try:
                        import base64
                        image_data = base64.b64decode(excel_image)
                        st.image(image_data, caption="Product Image", width=150)
                    except Exception as e:
                        st.warning(f"Could not display image: {e}")

                    # Remove Background Button
                    if st.button("✂️ Remove BG", key=f"excel_rembg_{idx}"):
                        with st.spinner("Removing background..."):
                            try:
                                import warnings
                                import os
                                os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
                                with warnings.catch_warnings():
                                    warnings.filterwarnings("ignore")
                                    import rembg
                                from PIL import Image
                                import io

                                img_data = base64.b64decode(
                                    product.excel_image)
                                input_img = Image.open(io.BytesIO(img_data))
                                with warnings.catch_warnings():
                                    warnings.simplefilter("ignore")
                                    output_img = rembg.remove(input_img)
                                output_buffer = io.BytesIO()
                                output_img.save(output_buffer, format='PNG')
                                product.excel_image = base64.b64encode(
                                    output_buffer.getvalue()).decode('utf-8')
                                st.rerun()
                            except Exception as e:
                                st.error(f"Failed to remove background: {e}")

                    # Image adjustment controls
                    with st.popover("🔧 Adjust Image"):
                        st.caption("Upload a new image or search the web")

                        # Upload custom image
                        st.markdown("**Upload Custom Image**")
                        upload_key = f"excel_upload_img_{idx}"
                        uploaded_img = st.file_uploader(
                            "Upload image",
                            type=['png', 'jpg', 'jpeg'],
                            key=upload_key,
                            label_visibility="collapsed"
                        )
                        if uploaded_img and not st.session_state.get(f"processed_{upload_key}"):
                            try:
                                from PIL import Image
                                import io
                                img_pil = Image.open(uploaded_img)
                                buffer = io.BytesIO()
                                img_pil.save(buffer, format='PNG')
                                product.excel_image = base64.b64encode(
                                    buffer.getvalue()).decode('utf-8')
                                st.session_state[f"processed_{upload_key}"] = True
                                st.success("Image replaced!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Failed to upload: {e}")
                        elif not uploaded_img:
                            # Clear processed flag when no file
                            st.session_state[f"processed_{upload_key}"] = False

                        st.divider()

                        # Search web for images
                        st.markdown("**🌐 Search Web for Image**")
                        schema_fields = st.session_state.get(
                            'schema_fields', DEFAULT_SCHEMA_FIELDS)
                        available_fields = [f['name'] for f in schema_fields]

                        search_fields = st.multiselect(
                            "Combine fields for search",
                            available_fields,
                            default=[
                                'name'] if 'name' in available_fields else available_fields[:1],
                            key=f"excel_search_fields_{idx}"
                        )

                        search_parts = []
                        for field in search_fields:
                            value = getattr(product, field, None) if hasattr(
                                product, field) else product.raw_attributes.get(field, '')
                            if value and str(value).strip():
                                search_parts.append(str(value).strip())
                        search_value = ' '.join(
                            search_parts) if search_parts else product.name

                        if st.button(f"🔍 Search '{search_value[:30]}...'", key=f"excel_web_search_{idx}"):
                            with st.spinner("Searching..."):
                                try:
                                    from src.web_image_search import search_images, download_image
                                    results = search_images(
                                        search_value, max_results=6)
                                    st.session_state[f"excel_web_results_{idx}"] = results
                                except Exception as e:
                                    st.error(f"Search failed: {e}")

                        web_results = st.session_state.get(
                            f"excel_web_results_{idx}")
                        if web_results:
                            st.markdown(
                                f"##### Found {len(web_results)} images")

                            # Check if we need to download a selected image
                            selected_key = f"excel_selected_img_{idx}"
                            if st.session_state.get(selected_key) is not None:
                                selected_idx = st.session_state[selected_key]
                                result = web_results[selected_idx]
                                with st.spinner("⏳ Downloading..."):
                                    from src.web_image_search import download_image
                                    img_b64 = download_image(result.url)
                                    if img_b64:
                                        product.excel_image = img_b64
                                    st.session_state[selected_key] = None
                                    st.session_state[f"excel_web_results_{idx}"] = None
                                    st.rerun()

                            cols = st.columns(3)
                            for r_idx, result in enumerate(web_results[:6]):
                                with cols[r_idx % 3]:
                                    try:
                                        st.image(result.thumbnail_url,
                                                 width='stretch')
                                        if st.button("✅ Use", key=f"excel_sel_img_{idx}_{r_idx}"):
                                            st.session_state[f"excel_selected_img_{idx}"] = r_idx
                                            st.rerun()
                                    except Exception:
                                        st.caption("Preview unavailable")
                else:
                    if st.button("🖼️ Load Image", key=f"load_img_btn_{idx}"):
                        st.session_state[show_image_key] = True
                        st.rerun()
            with col_fields:
                render_product_fields(product, idx)
        else:
            col_no_img, col_fields = st.columns([1, 3])
            with col_no_img:
                st.info("No image")

                # Upload or search for image
                with st.popover("➕ Add Image"):
                    st.markdown("**Upload Image**")
                    upload_new_key = f"excel_upload_new_img_{idx}"
                    uploaded_img = st.file_uploader(
                        "Upload image",
                        type=['png', 'jpg', 'jpeg'],
                        key=upload_new_key,
                        label_visibility="collapsed"
                    )
                    if uploaded_img and not st.session_state.get(f"processed_{upload_new_key}"):
                        try:
                            from PIL import Image
                            import io
                            img_pil = Image.open(uploaded_img)
                            buffer = io.BytesIO()
                            img_pil.save(buffer, format='PNG')
                            product.excel_image = base64.b64encode(
                                buffer.getvalue()).decode('utf-8')
                            st.session_state[f"processed_{upload_new_key}"] = True
                            st.success("Image added!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to upload: {e}")
                    elif not uploaded_img:
                        st.session_state[f"processed_{upload_new_key}"] = False

                    st.divider()

                    # Search web
                    st.markdown("**🌐 Search Web**")
                    search_value = product.name
                    if st.button(f"🔍 Search '{search_value[:25]}...'", key=f"excel_web_search_new_{idx}"):
                        with st.spinner("Searching..."):
                            try:
                                from src.web_image_search import search_images
                                results = search_images(
                                    search_value, max_results=6)
                                st.session_state[f"excel_web_results_new_{idx}"] = results
                            except Exception as e:
                                st.error(f"Search failed: {e}")

                    web_results = st.session_state.get(
                        f"excel_web_results_new_{idx}")
                    if web_results:
                        # Check if we need to download a selected image
                        selected_key = f"excel_selected_new_img_{idx}"
                        if st.session_state.get(selected_key) is not None:
                            selected_idx = st.session_state[selected_key]
                            result = web_results[selected_idx]
                            with st.spinner("⏳ Downloading..."):
                                from src.web_image_search import download_image
                                img_b64 = download_image(result.url)
                                if img_b64:
                                    product.excel_image = img_b64
                                st.session_state[selected_key] = None
                                st.session_state[f"excel_web_results_new_{idx}"] = None
                                st.rerun()

                        cols = st.columns(3)
                        for r_idx, result in enumerate(web_results[:6]):
                            with cols[r_idx % 3]:
                                try:
                                    st.image(result.thumbnail_url,
                                             width='stretch')
                                    if st.button("✅ Use", key=f"excel_sel_new_img_{idx}_{r_idx}"):
                                        st.session_state[f"excel_selected_new_img_{idx}"] = r_idx
                                        st.rerun()
                                except Exception:
                                    pass
            with col_fields:
                render_product_fields(product, idx)

        col1, col2 = st.columns(2)
        with col1:
            if product.needs_review:
                if st.button("✅ Approve", key=f"excel_approve_{idx}"):
                    product.needs_review = False
                    st.rerun()
        with col2:
            if st.button("🗑️ Delete", key=f"excel_delete_{idx}"):
                st.session_state.excel_products = [
                    p for p in st.session_state.excel_products if p.product_id != product.product_id
                ]
                st.rerun()


def render_product_fields(product: Product, idx: int):
    """Render editable product fields"""
    schema_fields = st.session_state.get(
        'schema_fields', DEFAULT_SCHEMA_FIELDS)

    for field in schema_fields:
        field_name = field['name']
        value = getattr(product, field_name, None) if hasattr(
            product, field_name) else product.raw_attributes.get(field_name, '')

        if field_name == 'price' and value:
            if hasattr(value, 'amount'):
                value = str(value.amount)
            else:
                value = str(value)

        new_value = st.text_input(
            f"{field_name.replace('_', ' ').title()}",
            value=str(value) if value else "",
            key=f"excel_prod_{idx}_{field_name}"
        )

        if new_value != str(value or ""):
            if hasattr(product, field_name):
                setattr(product, field_name, new_value)
            else:
                product.raw_attributes[field_name] = new_value





def do_llm_excel_extraction(sheet, start_row: int, end_row: int, batch_size: int = 20, include_images: bool = True):
    """Extract products by sending row batches to the LLM with full schema/hints context."""
    import json as _json
    from src.schemas import Product, Price, Currency

    # Clear existing products
    st.session_state.excel_products = []

    extractor = LLMExtractor()
    schema_fields = st.session_state.get('schema_fields', DEFAULT_SCHEMA_FIELDS)
    layout_context = st.session_state.get('page_layout_context', '')

    # Build image index if needed
    image_store = None
    if include_images:
        image_store = getattr(st.session_state.excel_doc, 'image_store', None)
        sheet.build_row_image_index()

    # Chunk rows into batches (start_row is Excel-absolute, convert to 0-based data index for slicing)
    data_start = start_row - sheet.data_start_excel_row
    data_end = end_row - sheet.data_start_excel_row + 1
    all_rows = sheet.rows[data_start:data_end]
    total_rows = len(all_rows)
    num_batches = (total_rows + batch_size - 1) // batch_size

    with st.status(f"🧠 Extracting with LLM ({num_batches} batches)...", expanded=True) as status:
        progress_bar = st.progress(0, text=f"Processing batch 0/{num_batches}...")
        total_products = 0

        for batch_idx in range(num_batches):
            batch_start = batch_idx * batch_size
            batch_end = min(batch_start + batch_size, total_rows)
            batch_rows = all_rows[batch_start:batch_end]

            # Format rows as text with actual Excel row numbers
            row_lines = []
            for i, row in enumerate(batch_rows):
                # start_row is already Excel-absolute
                excel_row = start_row + batch_start + i
                row_dict = {h: v for h, v in zip(sheet.headers, row)}
                row_lines.append(f"Row {excel_row}: {_json.dumps(row_dict, default=str)}")
            data_text = "\n".join(row_lines)

            st.text(f"Batch {batch_idx + 1}/{num_batches}: Excel rows {start_row + batch_start}-{start_row + batch_end - 1}")

            try:
                # Call LLM with full schema context
                products = extractor.extract_products_from_excel(
                    data_text=data_text,
                    batch_num=batch_idx + 1,
                    schema_fields=schema_fields,
                    layout_context=layout_context,
                    headers=sheet.headers,
                )

                # Build a lookup: for each row in this batch, create a set of
                # searchable values so we can match products to their source row
                batch_row_values = []
                for i, row in enumerate(batch_rows):
                    excel_row = start_row + batch_start + i
                    # Collect all non-empty string values from this row for matching
                    vals = set()
                    for v in row:
                        if v is not None:
                            s = str(v).strip()
                            if s:
                                vals.add(s.lower())
                    batch_row_values.append((excel_row, i, vals))

                # Batch-preload images for all rows in this batch (one workbook open)
                if include_images and image_store:
                    batch_row_set = set()
                    for excel_row, _, _ in batch_row_values:
                        batch_row_set.add(excel_row - 1)  # Convert to 0-based for preload
                    image_store.preload_for_rows(
                        sheet.name, sheet.image_metas, batch_row_set)

                # Process each product
                for prod_idx, product in enumerate(products):
                    # Try to match product to its source row by finding which row
                    # contains the product's name or SKU
                    matched_row = None
                    match_key = (product.name or '').strip().lower()
                    sku_key = (product.sku or '').strip().lower()

                    if match_key or sku_key:
                        best_score = 0
                        for excel_row, data_idx, row_vals in batch_row_values:
                            score = 0
                            # Check if product name appears in any cell value
                            if match_key:
                                for v in row_vals:
                                    if match_key in v or v in match_key:
                                        score += 2
                                        break
                            # Check SKU match
                            if sku_key and sku_key in row_vals:
                                score += 3
                            if score > best_score:
                                best_score = score
                                matched_row = excel_row

                    # Validate and fix LLM-provided source_rows
                    # The prompt sends absolute Excel rows (e.g. "Row 20: {...}"),
                    # but the LLM sometimes returns relative indices (1, 2, 3...)
                    batch_excel_start = start_row + batch_start
                    batch_excel_end = start_row + batch_end - 1
                    if product.source_rows and len(product.source_rows) > 0:
                        # Check if LLM rows fall within the batch range
                        in_range = all(batch_excel_start <= r <= batch_excel_end for r in product.source_rows)
                        if not in_range:
                            # Try interpreting as 1-based relative indices within the batch
                            offset_rows = [batch_excel_start + (r - 1) for r in product.source_rows]
                            if all(batch_excel_start <= r <= batch_excel_end for r in offset_rows):
                                product.source_rows = offset_rows
                            else:
                                # Can't salvage — discard and let matching handle it
                                product.source_rows = []

                    # Set source_rows: prefer validated LLM-provided, then matched, then fallback
                    if product.source_rows and len(product.source_rows) > 0:
                        # LLM provided valid source_rows — keep them
                        if matched_row and matched_row not in product.source_rows:
                            product.source_rows.insert(0, matched_row)
                    elif matched_row:
                        product.source_rows = [matched_row]
                    else:
                        # Positional fallback: distribute evenly
                        rows_per_product = max(1, len(batch_rows) // max(1, len(products)))
                        product.source_rows = [start_row + batch_start + prod_idx * rows_per_product]

                    # Associate image — prefer source_rows, fall back to positional
                    if include_images and image_store:
                        image_found = False
                        # Primary: use the product's actual source_rows (1-based Excel rows)
                        if product.source_rows:
                            for src_row in product.source_rows:
                                openpyxl_row = src_row - 1  # 1-based Excel → 0-based openpyxl
                                meta = sheet.get_image_meta_for_row(openpyxl_row)
                                if meta:
                                    b64 = image_store.get_image_base64(meta.sheet_name, meta.index)
                                    if b64:
                                        product.excel_image = b64
                                        image_found = True
                                    break

                        # Fallback: positional ownership when source_rows didn't yield an image
                        if not image_found:
                            rows_per_product = max(1, len(batch_rows) // max(1, len(products)))
                            own_start = start_row + batch_start + prod_idx * rows_per_product
                            own_end = start_row + batch_start + (prod_idx + 1) * rows_per_product
                            if prod_idx == len(products) - 1:
                                own_end = start_row + batch_end
                            for r in range(own_start, own_end + 1):
                                openpyxl_row = r - 1
                                meta = sheet.get_image_meta_for_row(openpyxl_row)
                                if meta:
                                    b64 = image_store.get_image_base64(meta.sheet_name, meta.index)
                                    if b64:
                                        product.excel_image = b64
                                    break

                    st.session_state.excel_products.append(product)
                    total_products += 1

            except Exception as e:
                st.warning(f"Batch {batch_idx + 1} failed: {e}")

            # Update progress
            progress = (batch_idx + 1) / num_batches
            progress_bar.progress(progress, text=f"Processed batch {batch_idx + 1}/{num_batches} ({total_products} products)")

        progress_bar.progress(1.0, text=f"✅ Done: {total_products} products extracted!")
        status.update(label=f"✅ Extracted {total_products} products from {total_rows} rows", state="complete")

    st.rerun()


def render_excel_review_view():
    """Render the Excel product review view"""
    if not st.session_state.excel_products:
        st.info("No products extracted yet. Go to Extraction view to extract products.")
        return

    st.subheader("🔍 Product Review")

    # Filter options
    col1, col2 = st.columns([1, 2])
    with col1:
        filter_review = st.checkbox(
            "Show only needs review", value=True, key="excel_filter_review")
    with col2:
        search_query = st.text_input(
            "🔍 Search products",
            placeholder="Search by name, SKU, or any field...",
            key="excel_search_query",
            label_visibility="collapsed"
        )

    # Filter products
    products = st.session_state.excel_products.copy()

    if filter_review:
        products = [p for p in products if p.needs_review]

    # Apply search filter
    if search_query:
        query_lower = search_query.lower()
        filtered = []
        for p in products:
            # Search in name, sku, and raw_attributes
            if query_lower in (p.name or '').lower():
                filtered.append(p)
            elif query_lower in (getattr(p, 'sku', '') or '').lower():
                filtered.append(p)
            elif any(query_lower in str(v).lower() for v in p.raw_attributes.values()):
                filtered.append(p)
        products = filtered

    st.caption(f"Showing {len(products)} products")

    # Bulk actions
    col_approve, col_delete = st.columns(2)
    with col_approve:
        if st.button("✅ Approve All", key="excel_approve_all", type="primary"):
            for p in products:
                p.needs_review = False
            st.rerun()
    with col_delete:
        if st.button("🗑️ Delete All Shown", key="excel_delete_all"):
            ids_to_delete = {p.product_id for p in products}
            st.session_state.excel_products = [
                p for p in st.session_state.excel_products if p.product_id not in ids_to_delete
            ]
            st.rerun()

    st.divider()

    # Display products
    for idx, product in enumerate(products):
        # Offset to avoid key conflicts
        render_excel_product_card(product, idx + 1000)


def render_excel_export_view():
    """Render the Excel export view"""
    st.subheader("📤 Export Products")

    if not st.session_state.excel_products:
        st.info("No products to export. Extract products first.")
        return

    # Summary
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Products", len(st.session_state.excel_products))
    with col2:
        reviewed = sum(
            1 for p in st.session_state.excel_products if not p.needs_review)
        st.metric("Reviewed", reviewed)
    with col3:
        needs_review = sum(
            1 for p in st.session_state.excel_products if p.needs_review)
        st.metric("Needs Review", needs_review)

    st.divider()

    # Export options
    st.subheader("⚙️ Export Options")

    schema_fields = st.session_state.get(
        'schema_fields', DEFAULT_SCHEMA_FIELDS)

    col1, col2 = st.columns(2)
    with col1:
        include_reviewed_only = st.checkbox(
            "Export reviewed only", value=False, key="excel_export_reviewed")
    with col2:
        include_row_info = st.checkbox(
            "Include row number", value=True, key="excel_export_row_info")

    include_images = st.checkbox("📷 Include product images", value=True, key="excel_export_images",
                                 help="Embed images in the exported Excel file (may increase file size)")

    remove_bg_on_export = st.checkbox("✂️ Remove background from images", value=False, key="excel_export_remove_bg",
                                      help="Automatically remove backgrounds from all images during export (slower)")

    # Filter products
    products_to_export = st.session_state.excel_products
    if include_reviewed_only:
        products_to_export = [
            p for p in products_to_export if not p.needs_review]

    # Count products with images
    products_with_images = sum(
        1 for p in products_to_export if getattr(p, 'excel_image', None))
    if products_with_images > 0:
        st.info(
            f"🖼️ {products_with_images} of {len(products_to_export)} products have associated images")

    st.divider()

    # Excel export
    if st.button("📊 Generate Excel File", type="primary", key="excel_generate_export"):
        with st.spinner("Generating Excel file..."):
            try:
                # Remove backgrounds if requested
                if remove_bg_on_export and include_images:
                    products_with_images = [
                        p for p in products_to_export if getattr(p, 'excel_image', None)]
                    if products_with_images:
                        progress_bar = st.progress(
                            0, text="Removing backgrounds...")
                        import warnings
                        import os
                        os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
                        with warnings.catch_warnings():
                            warnings.filterwarnings("ignore")
                            import rembg
                        from PIL import Image
                        import io

                        for i, p in enumerate(products_with_images):
                            try:
                                img_data = base64.b64decode(p.excel_image)
                                input_img = Image.open(io.BytesIO(img_data))
                                with warnings.catch_warnings():
                                    warnings.simplefilter("ignore")
                                    output_img = rembg.remove(input_img)
                                output_buffer = io.BytesIO()
                                output_img.save(output_buffer, format='PNG')
                                p.excel_image = base64.b64encode(
                                    output_buffer.getvalue()).decode('utf-8')
                            except Exception:
                                pass  # Skip failed images
                            progress_bar.progress(
                                (i + 1) / len(products_with_images), text=f"Removing backgrounds... {i+1}/{len(products_with_images)}")
                        progress_bar.empty()

                excel_exporter = ExcelExporter()
                excel_bytes = excel_exporter.export_products_to_excel(
                    products=products_to_export,
                    schema_fields=schema_fields,
                    computed_fields=st.session_state.get('computed_fields', []),
                    include_images=include_images,
                    include_page_info=False,
                    include_row_info=include_row_info
                )

                st.download_button(
                    "📥 Download Excel",
                    excel_bytes,
                    file_name=f"excel_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("✅ Excel file generated!")
            except Exception as e:
                st.error(f"Failed to generate Excel: {e}")

    # Preview
    st.divider()
    st.subheader("📋 Preview")

    if products_to_export:
        import pandas as pd

        preview_rows = []
        for p in products_to_export[:10]:
            row = {}
            for field in schema_fields:
                field_name = field['name']
                if hasattr(p, field_name):
                    value = getattr(p, field_name)
                else:
                    value = p.raw_attributes.get(field_name, '')

                if field_name == 'price' and value:
                    if hasattr(value, 'amount'):
                        value = value.amount

                row[field_name] = value

            # Computed fields
            computed_fields = st.session_state.get('computed_fields', [])
            if computed_fields:
                comp_vals = evaluate_computed_fields(p, computed_fields)
                row.update(comp_vals)

            row['source_rows'] = ', '.join(
                map(str, getattr(p, 'source_rows', [])))
            preview_rows.append(row)

        df = pd.DataFrame(preview_rows)
        # Convert all columns to strings to avoid Arrow serialization errors
        df = df.astype(str)
        st.dataframe(df, width='stretch')

        if len(products_to_export) > 10:
            st.caption(
                f"Showing first 10 of {len(products_to_export)} products")


if __name__ == "__main__":
    main()
