"""
Version Comparison Agent - Streamlit Application
Compare document versions across PDF, Excel, and CSV formats.
"""
import streamlit as st
import io
import os
import re
import tempfile
from config import Config
from pdf_comparator import PDFComparator
from excel_csv_comparator import ExcelCSVComparator
from llm_client import LLMClient

# Page configuration
st.set_page_config(
    page_title="Version Comparison Agent",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    /* Import custom fonts */
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Outfit:wght@300;400;500;600;700&display=swap');
    
    /* Root variables */
    :root {
        --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        --secondary-gradient: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        --success-gradient: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        --dark-bg: #0f0f1a;
        --card-bg: rgba(255, 255, 255, 0.03);
        --border-color: rgba(255, 255, 255, 0.1);
        --text-primary: #ffffff;
        --text-secondary: rgba(255, 255, 255, 0.7);
    }
    
    /* Main container styling */
    .stApp {
        background: linear-gradient(180deg, #0f0f1a 0%, #1a1a2e 50%, #16213e 100%);
    }
    
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Typography */
    h1, h2, h3, h4 {
        font-family: 'Outfit', sans-serif !important;
    }
    
    p, span, div, label {
        font-family: 'Outfit', sans-serif !important;
    }
    
    /* Header styling */
    .main-header {
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 2rem;
    }
    
    .main-header h1 {
        font-size: 3rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 0.5rem;
        letter-spacing: -0.02em;
    }
    
    .main-header p {
        color: rgba(255, 255, 255, 0.6);
        font-size: 1.1rem;
        font-weight: 300;
    }
    
    /* Card styling */
    .upload-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        backdrop-filter: blur(10px);
        transition: all 0.3s ease;
    }
    
    .upload-card:hover {
        border-color: rgba(102, 126, 234, 0.4);
        box-shadow: 0 8px 32px rgba(102, 126, 234, 0.15);
    }
    
    .upload-card h3 {
        color: #fff;
        font-size: 1.1rem;
        font-weight: 500;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* File uploader styling */
    .stFileUploader {
        background: rgba(255, 255, 255, 0.02);
        border-radius: 12px;
        padding: 0.5rem;
    }
    
    .stFileUploader > div {
        border-radius: 12px !important;
    }
    
    .stFileUploader label {
        color: rgba(255, 255, 255, 0.8) !important;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        width: 100%;
        font-family: 'Outfit', sans-serif !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(56, 239, 125, 0.3);
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: rgba(255, 255, 255, 0.02);
        padding: 0.5rem;
        border-radius: 16px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border-radius: 12px;
        padding: 0.75rem 1.5rem;
        color: rgba(255, 255, 255, 0.6);
        font-weight: 500;
        transition: all 0.3s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(255, 255, 255, 0.05);
        color: white;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
    }
    
    /* Alert/Info boxes */
    .stAlert {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
    }
    
    /* Success message */
    .stSuccess {
        background: rgba(56, 239, 125, 0.1);
        border-color: rgba(56, 239, 125, 0.3);
    }
    
    /* Dataframe styling */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f0f1a 0%, #1a1a2e 100%);
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }
    
    [data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }
    
    /* Legend card */
    .legend-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 1rem;
        margin-top: 1rem;
    }
    
    .legend-item {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
    
    .legend-dot {
        width: 12px;
        height: 12px;
        border-radius: 50%;
    }
    
    .legend-dot.red { background: #ff6b6b; }
    .legend-dot.green { background: #51cf66; }
    .legend-dot.yellow { background: #ffd43b; }
    
    /* Spinner */
    .stSpinner > div {
        border-color: #667eea transparent transparent transparent;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(255, 255, 255, 0.03);
        border-radius: 12px;
        font-weight: 500;
    }
    
    /* Metric styling */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 600;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Selectbox */
    .stSelectbox > div > div {
        background: rgba(255, 255, 255, 0.03);
        border-color: rgba(255, 255, 255, 0.1);
        border-radius: 12px;
    }
    
    /* Radio buttons */
    .stRadio > div {
        background: rgba(255, 255, 255, 0.02);
        padding: 1rem;
        border-radius: 12px;
    }
    
    /* Divider */
    hr {
        border-color: rgba(255, 255, 255, 0.08);
    }
    
    /* Code blocks */
    code {
        font-family: 'JetBrains Mono', monospace !important;
        background: rgba(102, 126, 234, 0.1);
        border-radius: 6px;
        padding: 0.2rem 0.4rem;
    }
</style>
""", unsafe_allow_html=True)


def render_header():
    """Render the main header."""
    st.markdown("""
    <div class="main-header">
        <h1>🔄 Version Comparison Agent</h1>
        <p>Compare document versions across PDF, Excel, and CSV formats with AI-powered analysis</p>
    </div>
    """, unsafe_allow_html=True)


def render_sidebar():
    """Render sidebar with configuration and legend."""
    with st.sidebar:
        st.markdown("### ⚙️ Configuration")
        
        # Validate Azure OpenAI configuration
        is_valid, message = Config.validate()
        
        if is_valid:
            st.success("✓ Azure OpenAI Connected", icon="✅")
        else:
            st.error(f"⚠️ {message}")
            st.markdown("""
            Please configure your `.env` file with:
            - `AZURE_OPENAI_API_KEY`
            - `AZURE_OPENAI_ENDPOINT`
            - `AZURE_OPENAI_DEPLOYMENT`
            - `AZURE_OPENAI_API_VERSION`
            """)
        
        st.markdown("---")
        
        # File size info
        st.markdown("### 📁 File Limits")
        st.markdown(f"Max file size: **{Config.MAX_FILE_SIZE_MB} MB**")
        
        st.markdown("---")
        
        # Legend for PDF
        st.markdown("### 🎨 PDF Color Legend")
        st.markdown("""
        <div class="legend-card">
            <div class="legend-item">
                <div class="legend-dot green"></div>
                <span>Added content</span>
            </div>
            <div class="legend-item">
                <div class="legend-dot yellow"></div>
                <span>Modified content</span>
            </div>
            <div class="legend-item" style="opacity: 0.7;">
                <div class="legend-dot red"></div>
                <span>Removed (listed only)</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.caption("*Removed content is shown in the changes list since it doesn't exist in the new file.*")
        
        st.markdown("---")
        
        st.markdown("### 📊 Excel/CSV Legend")
        st.markdown("""
        - **`-`** = No change
        - **Value** = Change description
        """)


def validate_file_size(file) -> bool:
    """Check if file is within size limit."""
    if file is None:
        return True
    return file.size <= Config.MAX_FILE_SIZE_BYTES


def render_pdf_comparison():
    """Render PDF comparison interface."""
    st.markdown("### 📄 PDF Version Comparison")
    st.markdown("Upload two PDF files to compare. Changes will be highlighted in the output document.")
    
    # Extraction method selector
    extraction_options = ["Standard (PyMuPDF)", "LLM-Powered (PDF→Markdown)"]
    
    extraction_method = st.radio(
        "📑 Extraction Method",
        extraction_options,
        horizontal=True,
        key="pdf_extraction_method",
        help="LLM-Powered converts PDF to Markdown for AI comparison"
    )
    
    use_marker = "LLM-Powered" in extraction_method
    
    if use_marker:
        st.info("📝 **LLM-Powered**: Converts PDFs to Markdown, then uses AI for intelligent comparison. Best accuracy!")
    else:
        st.markdown("*Standard extraction - fast local processing with PyMuPDF.*")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="upload-card">
            <h3>📂 Older Version</h3>
        </div>
        """, unsafe_allow_html=True)
        old_file = st.file_uploader(
            "Upload older PDF",
            type=["pdf"],
            key="pdf_old",
            help=f"Select the older version of the PDF document (Max size: {Config.MAX_FILE_SIZE_MB} MB)"
        )
        st.caption(f"⚠️ Maximum file size: {Config.MAX_FILE_SIZE_MB} MB")
        if old_file and not validate_file_size(old_file):
            st.error(f"File too large. Max size: {Config.MAX_FILE_SIZE_MB} MB")
            old_file = None
    
    with col2:
        st.markdown("""
        <div class="upload-card">
            <h3>📂 Newer Version</h3>
        </div>
        """, unsafe_allow_html=True)
        new_file = st.file_uploader(
            "Upload newer PDF",
            type=["pdf"],
            key="pdf_new",
            help=f"Select the newer version of the PDF document (Max size: {Config.MAX_FILE_SIZE_MB} MB)"
        )
        st.caption(f"⚠️ Maximum file size: {Config.MAX_FILE_SIZE_MB} MB")
        if new_file and not validate_file_size(new_file):
            st.error(f"File too large. Max size: {Config.MAX_FILE_SIZE_MB} MB")
            new_file = None
    
    # Comparison button
    st.markdown("")
    
    if st.button("🔍 Compare PDF Versions", key="compare_pdf", use_container_width=True):
        if not old_file or not new_file:
            st.warning("Please upload both PDF files to compare.")
            return
        
        # Check configuration
        is_valid, message = Config.validate()
        if not is_valid:
            st.error(f"Configuration error: {message}")
            return
        
        if use_marker:
            # LLM-Powered (PDF to Markdown) comparison
            run_marker_pdf_comparison(old_file, new_file)
        else:
            # Standard PyMuPDF comparison
            run_standard_pdf_comparison(old_file, new_file)


def run_standard_pdf_comparison(old_file, new_file):
    """Run standard PDF comparison with PyMuPDF."""
    with st.spinner("🔄 Analyzing documents and identifying changes..."):
            try:
                comparator = PDFComparator()
                
                # Reset file pointers and read file bytes
                old_file.seek(0)
                new_file.seek(0)
                old_bytes = old_file.read()
                new_bytes = new_file.read()
                
            # Perform comparison
                result = comparator.compare_pdfs(old_bytes, new_bytes)
                
                st.success("✅ Comparison complete!")
                
                # Display summary
                summary = comparator.get_change_summary(result["changes"])
                st.info(f"**Summary:** {summary}")
                
                # Download button - single file
                st.markdown("### 📥 Download Highlighted PDF")
                
                st.download_button(
                    label="⬇️ Download Comparison Result",
                    data=result["highlighted_pdf"],
                    file_name=f"comparison_{new_file.name}",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                # Show detailed changes
                render_pdf_changes(result["changes"])
            
            except Exception as e:
                st.error(f"Error during comparison: {str(e)}")


def run_marker_pdf_comparison(old_file, new_file):
    """Run PDF comparison using PyMuPDF4LLM (PDF to Markdown)."""
    import fitz
    
    try:
        from marker_pdf_comparator import compare_pdfs_with_marker, get_extraction_method
    except ImportError as e:
        st.error(f"PDF comparator not available: {e}")
        st.info("Install with: `pip install pymupdf4llm`")
        return
    
    # Show available libraries and let user choose
    from marker_pdf_comparator import get_available_libraries, set_active_library
    
    available_libs = get_available_libraries()
    current_lib = get_extraction_method()
    
    if len(available_libs) > 1:
        selected_lib = st.selectbox(
            "📚 PDF-to-Markdown Library",
            available_libs,
            index=available_libs.index(current_lib) if current_lib in available_libs else 0,
            key="pdf_library_selector"
        )
        if selected_lib != current_lib:
            set_active_library(selected_lib)
        st.info(f"🔍 Using **{selected_lib}** for PDF extraction")
    else:
        st.info(f"🔍 Using **{current_lib}** - Converts PDF to Markdown for AI comparison")
    
    with st.spinner("🔄 Converting PDFs to Markdown with Marker..."):
        try:
            # Reset file pointers and read bytes
            old_file.seek(0)
            new_file.seek(0)
            old_bytes = old_file.read()
            new_bytes = new_file.read()
            
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Step 1: Convert and compare
            status_text.text("📄 Converting OLD PDF to Markdown...")
            progress_bar.progress(20)
            
            status_text.text("📄 Converting NEW PDF to Markdown...")
            progress_bar.progress(40)
            
            status_text.text("🤖 Comparing documents with AI...")
            progress_bar.progress(60)
            
            # Run the comparison
            result = compare_pdfs_with_marker(old_bytes, new_bytes)
            
            progress_bar.progress(80)
            
            # Create highlighted PDF (page-by-page approach)
            status_text.text("🎨 Creating highlighted PDF...")
            if "page_changes" in result and result["page_changes"]:
                # Prepare file info for summary page
                import fitz
                from datetime import datetime
                
                old_doc = fitz.open(stream=old_bytes, filetype="pdf")
                new_doc = fitz.open(stream=new_bytes, filetype="pdf")
                
                # Get file metadata
                old_name = old_file.name if hasattr(old_file, 'name') else "Old File.pdf"
                new_name = new_file.name if hasattr(new_file, 'name') else "New File.pdf"
                
                # Try to get file modification time, fallback to current time if not available
                try:
                    if hasattr(old_file, 'name') and os.path.exists(old_file.name):
                        old_timestamp = datetime.fromtimestamp(os.path.getmtime(old_file.name)).strftime("%d/%m/%Y %H:%M:%S")
                    else:
                        old_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                except:
                    old_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
                try:
                    if hasattr(new_file, 'name') and os.path.exists(new_file.name):
                        new_timestamp = datetime.fromtimestamp(os.path.getmtime(new_file.name)).strftime("%d/%m/%Y %H:%M:%S")
                    else:
                        new_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                except:
                    new_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
                old_file_info = {
                    "name": old_name,
                    "pages": len(old_doc),
                    "size_kb": len(old_bytes) // 1024,
                    "timestamp": old_timestamp
                }
                
                new_file_info = {
                    "name": new_name,
                    "pages": len(new_doc),
                    "size_kb": len(new_bytes) // 1024,
                    "timestamp": new_timestamp
                }
                
                old_doc.close()
                new_doc.close()
                
                # Use new page-by-page highlighting with both PDFs
                highlighted_pdf = create_highlighted_pdf_page_by_page(
                    old_bytes, new_bytes, result["page_changes"],
                    old_file_info=old_file_info,
                    new_file_info=new_file_info,
                    changes=result["changes"]
                )
            else:
                # Fallback to old method
                highlighted_pdf = create_highlighted_pdf_from_llm_changes(
                    new_bytes, result["changes"]
                )
            
            progress_bar.progress(100)
            status_text.empty()
            progress_bar.empty()
            
            st.success("✅ Marker comparison complete!")
            
            # Display summary
            changes = result["changes"]
            added = len(changes.get("added", []))
            removed = len(changes.get("removed", []))
            modified = len(changes.get("modified", []))
            
            summary_parts = []
            if removed > 0:
                summary_parts.append(f"🔴 {removed} removed")
            if added > 0:
                summary_parts.append(f"🟢 {added} added")
            if modified > 0:
                summary_parts.append(f"🟡 {modified} modified")
            
            summary = " | ".join(summary_parts) if summary_parts else "No changes detected"
            st.info(f"**Summary:** {summary}")
            
            if changes.get("summary"):
                st.markdown(f"**AI Summary:** {changes['summary']}")
            
            # Download button
            st.markdown("### 📥 Download Highlighted PDF")
            st.download_button(
                label="⬇️ Download Comparison Result",
                data=highlighted_pdf,
                file_name=f"marker_comparison_{new_file.name}",
                mime="application/pdf",
                use_container_width=True
            )
            
            # Show Markdown preview
            with st.expander("📝 View Extracted Markdown", expanded=False):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**OLD Document:**")
                    st.text_area("Old Markdown", result["old_markdown"][:5000], height=300, disabled=True)
                with col2:
                    st.markdown("**NEW Document:**")
                    st.text_area("New Markdown", result["new_markdown"][:5000], height=300, disabled=True)
            
            # Show detailed changes
            render_marker_changes(changes)
            
        except Exception as e:
            st.error(f"Error during Marker comparison: {str(e)}")
            import traceback
            with st.expander("Error Details"):
                st.code(traceback.format_exc())


def clean_text_for_highlighting(text: str) -> list:
    """
    Clean text and extract multiple search candidates for highlighting.
    Returns list of cleaned text variants to try.
    """
    if not text:
        return []
    
    candidates = []
    text = str(text).strip()
    
    # 1. Original text (as-is)
    if len(text) > 2:
        candidates.append(text)
    
    # 2. Remove watermark artifacts (T, F, A, R, D as standalone letters)
    cleaned = re.sub(r'\b[TFARD]\b', '', text)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    if cleaned and cleaned != text and len(cleaned) > 2:
        candidates.append(cleaned)
    
    # 3. Extract numbers (for financial values)
    numbers = re.findall(r'[\d,()]+', text)
    for num in numbers:
        if len(num) > 1:  # At least 2 chars (like "1" or "-")
            candidates.append(num)
    
    # 4. Extract first significant number/value
    # Pattern: (370,308) or 370,308 or -370,308
    value_match = re.search(r'[\(]?[\d,]+[\)]?', text)
    if value_match:
        value = value_match.group(0)
        if value not in candidates:
            candidates.append(value)
    
    # 5. Extract key words (for text changes)
    words = re.findall(r'\b\w{3,}\b', text)  # Words 3+ chars
    if words and len(' '.join(words[:3])) > 5:
        candidates.append(' '.join(words[:3]))
    
    # Remove duplicates while preserving order
    seen = set()
    unique_candidates = []
    for c in candidates:
        c_lower = c.lower()
        if c_lower not in seen and len(c) >= 2:
            seen.add(c_lower)
            unique_candidates.append(c)
    
    return unique_candidates


def highlight_with_context(doc, value: str, field: str, color: tuple) -> bool:
    """
    Highlight text using value and field context for better matching.
    Tries multiple strategies to find and highlight the value.
    """
    import fitz
    
    if not value:
        return False
    
    value = str(value).strip()
    if len(value) < 1:
        return False
    
    # Strategy 1: Direct exact match
    for page in doc:
        instances = page.search_for(value, quads=True)
        if instances:
            try:
                annot = page.add_highlight_annot(instances[0])
                annot.set_colors(stroke=color)
                annot.update()
                return True
            except Exception:
                pass
    
    # Strategy 2: Clean value (remove commas, normalize)
    clean_value = value.replace(',', '').replace('(', '').replace(')', '').strip()
    if clean_value and clean_value != value:
        for page in doc:
            # Try to find the numeric part
            instances = page.search_for(clean_value, quads=True)
            if instances:
                try:
                    annot = page.add_highlight_annot(instances[0])
                    annot.set_colors(stroke=color)
                    annot.update()
                    return True
                except Exception:
                    pass
    
    # Strategy 3: Extract just the number (for values like "(370,308)")
    numeric_match = re.search(r'[\d,]+', value)
    if numeric_match:
        numeric = numeric_match.group(0)
        for page in doc:
            instances = page.search_for(numeric, quads=True)
            if instances:
                try:
                    annot = page.add_highlight_annot(instances[0])
                    annot.set_colors(stroke=color)
                    annot.update()
                    return True
                except Exception:
                    pass
    
    # Strategy 4: Use field context to locate value
    if field:
        field_keywords = []
        # Extract meaningful words from field
        field_words = re.findall(r'\b[A-Z][a-z]+\b', field)
        field_keywords.extend(field_words)
        # Also try last part of field (after >)
        if '>' in field:
            last_part = field.split('>')[-1].strip()
            if last_part:
                field_keywords.append(last_part)
        
        # Try each field keyword
        for keyword in field_keywords[:3]:  # Try first 3 keywords
            for page in doc:
                field_instances = page.search_for(keyword, quads=True)
                if field_instances:
                    # Found field, now look for value nearby
                    # Try with original value
                    value_instances = page.search_for(value, quads=True)
                    if not value_instances and numeric_match:
                        # Try with just the number
                        value_instances = page.search_for(numeric_match.group(0), quads=True)
                    
                    if value_instances:
                        # Find value instance closest to field
                        field_rect = field_instances[0].rect if hasattr(field_instances[0], 'rect') else fitz.Rect(field_instances[0])
                        
                        best_instance = None
                        best_distance = float('inf')
                        
                        for val_inst in value_instances:
                            val_rect = val_inst.rect if hasattr(val_inst, 'rect') else fitz.Rect(val_inst)
                            
                            # Calculate distance (prefer same line, then nearby)
                            y_diff = abs(val_rect.y0 - field_rect.y0)
                            x_diff = abs(val_rect.x0 - field_rect.x1)
                            distance = y_diff * 10 + x_diff  # Heavy penalty for different lines
                            
                            if distance < best_distance and distance < 1000:  # Within reasonable distance
                                best_distance = distance
                                best_instance = val_inst
                        
                        if best_instance:
                            try:
                                annot = page.add_highlight_annot(best_instance)
                                annot.set_colors(stroke=color)
                                annot.update()
                                return True
                            except Exception:
                                pass
    
    return False


def extract_changed_values(old_text: str, new_text: str) -> list:
    """
    Compare old and new text and extract only the values that actually changed.
    Returns list of changed values from the NEW text.
    """
    if not old_text or not new_text:
        return []
    
    # Normalize: remove watermark artifacts and normalize whitespace
    old_clean = re.sub(r'\b[TFARD]\b', '', old_text)
    old_clean = re.sub(r'\s+', ' ', old_clean).strip()
    new_clean = re.sub(r'\b[TFARD]\b', '', new_text)
    new_clean = re.sub(r'\s+', ' ', new_clean).strip()
    
    # Extract all numeric values with their positions
    # Pattern: matches numbers like 137,260, (111,125), -370,308, etc.
    old_matches = list(re.finditer(r'[\(]?[\d,]+[\)]?', old_clean))
    new_matches = list(re.finditer(r'[\(]?[\d,]+[\)]?', new_clean))
    
    old_values = [m.group(0) for m in old_matches]
    new_values = [m.group(0) for m in new_matches]
    
    changed = []
    
    # Strategy 1: If same number of values, compare positionally
    if len(old_values) == len(new_values):
        for old_v, new_v in zip(old_values, new_values):
            # Normalize for comparison
            old_norm = old_v.replace(',', '').replace('(', '').replace(')', '').replace('-', '')
            new_norm = new_v.replace(',', '').replace('(', '').replace(')', '').replace('-', '')
            if old_norm != new_norm:
                changed.append(new_v)
    else:
        # Strategy 2: Different counts - find values in new that aren't in old
        old_set = set(v.replace(',', '') for v in old_values)
        for new_v in new_values:
            new_norm = new_v.replace(',', '').replace('(', '').replace(')', '').replace('-', '')
            if new_norm not in old_set:
                changed.append(new_v)
    
    # Strategy 3: For text changes (non-numeric), find new words
    # Extract meaningful words (3+ chars, not common words)
    common_words = {'the', 'and', 'for', 'account', 'current', 'at', 'to', 'of', 'in', 'on'}
    old_words = set(w.lower() for w in re.findall(r'\b\w{3,}\b', old_clean.lower()) if w not in common_words)
    new_words = set(w.lower() for w in re.findall(r'\b\w{3,}\b', new_clean.lower()) if w not in common_words)
    
    added_words = new_words - old_words
    if added_words:
        # Get actual words from new_text (with original case)
        for word in re.findall(r'\b\w{3,}\b', new_clean):
            if word.lower() in added_words:
                changed.append(word)
    
    # Remove duplicates while preserving order
    seen = set()
    unique_changed = []
    for c in changed:
        c_lower = str(c).lower()
        if c_lower not in seen:
            seen.add(c_lower)
            unique_changed.append(c)
    
    return unique_changed


def _verify_consecutive_instances(instances, page, max_line_gap=50):
    """
    Verify that multiple instances form a consecutive block of text.
    Returns the instances if they're consecutive, None otherwise.
    """
    if not instances or len(instances) <= 1:
        return instances
    
    import fitz
    
    # Sort instances by position (top to bottom, left to right)
    sorted_instances = []
    for inst in instances:
        if hasattr(inst, 'rect'):
            rect = inst.rect
        else:
            rect = fitz.Rect(inst)
        sorted_instances.append((rect.y0, rect.x0, inst))
    
    sorted_instances.sort(key=lambda x: (x[0], x[1]))
    
    # Check if instances are consecutive (each one is close to the next)
    verified = [sorted_instances[0][2]]  # Start with first instance
    
    for i in range(1, len(sorted_instances)):
        # Get actual rects
        prev_inst = sorted_instances[i-1][2]
        curr_inst = sorted_instances[i][2]
        
        prev_rect = prev_inst.rect if hasattr(prev_inst, 'rect') else fitz.Rect(prev_inst)
        curr_rect = curr_inst.rect if hasattr(curr_inst, 'rect') else fitz.Rect(curr_inst)
        
        # Check vertical gap (should be small for consecutive lines)
        y_gap = curr_rect.y0 - prev_rect.y1
        
        # Check horizontal overlap or proximity (text should be in similar x position)
        x_overlap = not (curr_rect.x1 < prev_rect.x0 or curr_rect.x0 > prev_rect.x1)
        x_proximity = abs(curr_rect.x0 - prev_rect.x0) < 200  # Within 200 pixels horizontally
        
        # If gap is reasonable and text is in similar horizontal position, consider consecutive
        if y_gap >= 0 and y_gap < max_line_gap and (x_overlap or x_proximity):
            verified.append(curr_inst)
        else:
            # Gap too large or not aligned - might not be the same block
            # If we have at least 2 consecutive instances, return them
            if len(verified) >= 2:
                return verified
            # Otherwise, this might not be a consecutive block
            return None
    
    return verified if len(verified) >= 2 else instances


def _find_consecutive_text_block(line_instances_list, page, max_line_gap=50):
    """
    Given a list of (line_text, instances) tuples, find a set of instances that form a consecutive block.
    Returns a list of instances that are consecutive, or None if no consecutive block found.
    """
    if not line_instances_list or len(line_instances_list) <= 1:
        return None
    
    import fitz
    
    # Try to find instances that are consecutive
    # We'll look for instances where each line's instance is close to the next line's instance
    
    best_block = None
    best_score = 0
    
    # For each line, try to find an instance that connects to the previous line
    # This is a simplified approach - we'll look for instances that are vertically aligned
    
    # Sort lines by their expected order (first line should be highest, last line lowest)
    # We'll try to find instances that form a vertical sequence
    
    # Strategy: For each line, pick the instance that's closest to where the previous line ended
    selected_instances = []
    
    for line_idx, (line_text, line_insts) in enumerate(line_instances_list):
        if not line_insts:
            continue
        
        if line_idx == 0:
            # First line: pick the topmost instance
            sorted_insts = sorted(line_insts, key=lambda x: (
                x.rect.y0 if hasattr(x, 'rect') else fitz.Rect(x).y0,
                x.rect.x0 if hasattr(x, 'rect') else fitz.Rect(x).x0
            ))
            selected_instances.append(sorted_insts[0])
        else:
            # Subsequent lines: find instance closest to previous line
            prev_inst = selected_instances[-1]
            prev_rect = prev_inst.rect if hasattr(prev_inst, 'rect') else fitz.Rect(prev_inst)
            
            best_inst = None
            best_distance = float('inf')
            
            for inst in line_insts:
                curr_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                
                # Calculate distance (vertical gap + horizontal distance)
                y_gap = curr_rect.y0 - prev_rect.y1
                x_distance = abs(curr_rect.x0 - prev_rect.x0)
                
                # Prefer instances that are directly below and aligned
                if y_gap >= 0 and y_gap < max_line_gap:
                    distance = y_gap * 2 + x_distance
                    if distance < best_distance:
                        best_distance = distance
                        best_inst = inst
            
            if best_inst and best_distance < max_line_gap * 3:
                selected_instances.append(best_inst)
            else:
                # Can't find a consecutive instance for this line
                # Return what we have so far if it's at least 2 lines
                if len(selected_instances) >= 2:
                    return selected_instances
                return None
    
    return selected_instances if len(selected_instances) >= 2 else None


def _find_text_block_by_extraction(page, search_text, context_before="", context_after="", row_label=""):
    """
    Find complete text block by searching for it in chunks and combining results.
    This handles multi-line text better than searching for the entire string.
    Returns list of quads that make up the complete text block.
    """
    import fitz
    
    if not search_text:
        return None
    
    # Normalize search text (remove newlines, normalize spaces)
    search_normalized = re.sub(r'\s+', ' ', search_text.replace('\n', ' ').replace('\r', ' ')).strip()
    search_words = search_normalized.split()
    
    if len(search_words) < 5:
        return None  # Too short, use regular search
    
    # Strategy: Search for text in overlapping chunks and combine consecutive matches
    all_quads = []
    chunk_size = 7  # Search for 7 words at a time
    
    # Find context first to narrow search area
    context_rect = None
    if context_before:
        context_instances = page.search_for(context_before, quads=True)
        if context_instances:
            context_rect = context_instances[0].rect if hasattr(context_instances[0], 'rect') else fitz.Rect(context_instances[0])
    
    # Search for overlapping chunks
    for i in range(0, len(search_words) - chunk_size + 1, 3):  # Step by 3 for overlap
        chunk = " ".join(search_words[i:i+chunk_size])
        chunk_instances = page.search_for(chunk, quads=True)
        
        if chunk_instances:
            # Filter by context if available
            if context_rect:
                for inst in chunk_instances:
                    inst_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                    # Check if instance is near context
                    y_diff = abs(inst_rect.y0 - context_rect.y0)
                    if y_diff < 200:  # Within reasonable distance
                        all_quads.append(inst)
            else:
                all_quads.extend(chunk_instances)
    
    if not all_quads:
        return None
    
    # Remove duplicates and sort by position
    unique_quads = []
    seen_positions = set()
    
    for quad in all_quads:
        rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
        pos_key = (round(rect.y0, 1), round(rect.x0, 1))
        if pos_key not in seen_positions:
            seen_positions.add(pos_key)
            unique_quads.append((rect.y0, rect.x0, quad))
    
    # Sort by position
    unique_quads.sort(key=lambda x: (x[0], x[1]))
    
    # Verify they form a consecutive block
    if len(unique_quads) > 1:
        quads_only = [q[2] for q in unique_quads]
        verified = _verify_consecutive_instances(quads_only, page)
        if verified:
            return verified
    
    # Return all quads even if not perfectly consecutive
    return [q[2] for q in unique_quads]


def _find_complete_text_block(page, full_text, context_before="", context_after="", row_label=""):
    """
    Find and return ALL instances that make up a complete text block.
    For text changes, we need to highlight the ENTIRE text, not just a snippet.
    Returns list of instances (quads) that form the complete text block.
    """
    import fitz
    
    if not full_text:
        return None
    
    # CRITICAL: Replace newlines with spaces for PDF search (PDF text doesn't have exact newlines)
    # This is the key fix - page.search_for() doesn't handle \n well
    text_for_search = full_text.replace('\n', ' ').replace('\r', ' ')
    # Normalize multiple spaces
    text_for_search = re.sub(r'\s+', ' ', text_for_search).strip()
    
    # Build comprehensive list of search variants (longer snippets first for better matching)
    search_variants = []
    
    # Strategy 1: Try the full text (with newlines replaced) if not too long
    if len(text_for_search) <= 800:
        search_variants.append(text_for_search)
    
    # Strategy 2: Try first sentence (most reliable for finding start)
    # Split by periods, not newlines
    sentences = re.split(r'[.!?]\s+', text_for_search)
    if sentences and len(sentences[0]) > 10:
        first_sentence = sentences[0].strip()
        if not first_sentence.endswith('.'):
            first_sentence += '.'
        search_variants.append(first_sentence)
        # Also try first sentence with next sentence if it's short
        if len(first_sentence) < 50 and len(sentences) > 1:
            second_sent = sentences[1].strip()
            if not second_sent.endswith('.'):
                second_sent += '.'
            search_variants.append(first_sentence + " " + second_sent[:50])
    
    # Strategy 3: Try word-based snippets (longer first)
    words = text_for_search.split()
    if len(words) >= 10:
        search_variants.append(" ".join(words[:10]))
    if len(words) >= 7:
        search_variants.append(" ".join(words[:7]))
    if len(words) >= 5:
        search_variants.append(" ".join(words[:5]))
    
    # Strategy 4: Try character-based snippets
    if len(text_for_search) > 200:
        search_variants.append(text_for_search[:200].strip())
    if len(text_for_search) > 150:
        search_variants.append(text_for_search[:150].strip())
    if len(text_for_search) > 100:
        search_variants.append(text_for_search[:100].strip())
    
    # First, find the starting point using context (if available)
    context_rect = None
    search_region = None
    
    if context_before or row_label:
        # Try row_label first (most specific)
        if row_label:
            row_instances = page.search_for(row_label, quads=True)
            if row_instances:
                context_rect = row_instances[0].rect if hasattr(row_instances[0], 'rect') else fitz.Rect(row_instances[0])
        
        # Try context_before
        if not context_rect and context_before:
            context_words = context_before.split()
            context_snippets = []
            if len(context_words) > 5:
                context_snippets.append(" ".join(context_words[-5:]))
            if len(context_words) > 3:
                context_snippets.append(" ".join(context_words[-3:]))
            context_snippets.append(context_before)
            
            for snippet in context_snippets:
                snippet_instances = page.search_for(snippet, quads=True)
                if snippet_instances:
                    context_rect = snippet_instances[0].rect if hasattr(snippet_instances[0], 'rect') else fitz.Rect(snippet_instances[0])
                    break
        
        if context_rect:
            # Define search region around context
            search_region = fitz.Rect(
                max(0, context_rect.x0 - 200),
                max(0, context_rect.y0 - 50),
                min(page.rect.width, context_rect.x1 + 1200),
                min(page.rect.height, context_rect.y1 + 1000)
            )
    
    # Try each search variant
    best_instances = []
    best_score = 0
    
    for variant in search_variants:
        if not variant or len(variant.strip()) < 5:
            continue
        
        try:
            variant_instances = page.search_for(variant.strip(), quads=True)
            if variant_instances:
                # If we have context, filter instances to those near context
                if context_rect and search_region:
                    filtered_instances = []
                    for inst in variant_instances:
                        inst_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                        if search_region.intersects(inst_rect):
                            # Calculate score (prefer longer matches and closer to context)
                            y_diff = abs(inst_rect.y0 - context_rect.y0)
                            x_diff = abs(inst_rect.x0 - context_rect.x1) if inst_rect.x0 > context_rect.x1 else abs(context_rect.x0 - inst_rect.x1)
                            score = len(variant) * 10 - (y_diff * 5 + x_diff)
                            
                            if score > best_score:
                                best_score = score
                                best_instances = [inst]
                            elif abs(score - best_score) < 100:
                                best_instances.append(inst)
                else:
                    # No context, use all instances (prefer longer variants)
                    if len(variant) > best_score:
                        best_score = len(variant)
                        best_instances = variant_instances
        except Exception:
            continue
    
    # If we found instances, verify they form a consecutive block
    if best_instances:
        if len(best_instances) > 1:
            verified = _verify_consecutive_instances(best_instances, page)
            if verified:
                return verified
        return best_instances
    
    return None


def _find_text_with_context(page, search_text, context_before="", context_after="", row_label="", full_text=""):
    """
    Intelligently find text in PDF using context information.
    Uses context-based search first, then falls back to direct search with word matching.
    Returns list of instances (quads) or None.
    """
    import fitz
    
    instances = None
    
    # Strategy 1: Context-based search (MOST RELIABLE for text changes)
    # Find context first, then search for text nearby
    if context_before or context_after or row_label:
        context_rect = None
        context_instances_list = []
        
        # Try row_label first (most specific)
        if row_label:
            row_instances = page.search_for(row_label, quads=True)
            if row_instances:
                context_rect = row_instances[0].rect if hasattr(row_instances[0], 'rect') else fitz.Rect(row_instances[0])
                context_instances_list = row_instances
        
        # Try context_before (try multiple variations for better matching)
        if not context_rect and context_before:
            context_words = context_before.split()
            
            # Try different snippets of context_before
            context_snippets = []
            if len(context_words) > 5:
                # Use last 5 words
                context_snippets.append(" ".join(context_words[-5:]))
            if len(context_words) > 3:
                # Use last 3 words
                context_snippets.append(" ".join(context_words[-3:]))
            # Use full context
            context_snippets.append(context_before)
            
            for snippet in context_snippets:
                snippet_instances = page.search_for(snippet, quads=True)
                if snippet_instances:
                    context_rect = snippet_instances[0].rect if hasattr(snippet_instances[0], 'rect') else fitz.Rect(snippet_instances[0])
                    context_instances_list = snippet_instances
                    break
        
        # If we found context, search for text nearby
        if context_rect:
            # Define search region around context (expanded area)
            search_region = fitz.Rect(
                max(0, context_rect.x0 - 200),
                max(0, context_rect.y0 - 50),
                min(page.rect.width, context_rect.x1 + 1000),
                min(page.rect.height, context_rect.y1 + 600)
            )
            
            # Build comprehensive list of search variants
            search_variants = []
            
            # Use search_text if available
            if search_text:
                search_variants.append(search_text)
            
            # Use full_text if different from search_text
            if full_text and full_text != search_text:
                search_variants.append(full_text)
            
            # Try first sentence of search_text
            if search_text:
                sentences = search_text.split('. ')
                if sentences and len(sentences[0]) > 10:
                    search_variants.append(sentences[0].strip())
            
            # Try first 150 chars (longer for better matching)
            if search_text and len(search_text) > 150:
                search_variants.append(search_text[:150].strip())
            elif search_text and len(search_text) > 100:
                search_variants.append(search_text[:100].strip())
            
            # Try word-by-word matching (first 3-7 words for better coverage)
            if search_text:
                words = search_text.split()
                if len(words) >= 7:
                    search_variants.append(" ".join(words[:7]))
                if len(words) >= 5:
                    search_variants.append(" ".join(words[:5]))
                if len(words) >= 3:
                    search_variants.append(" ".join(words[:3]))
            
            # Try same for full_text
            if full_text and full_text != search_text:
                words = full_text.split()
                if len(words) >= 5:
                    search_variants.append(" ".join(words[:5]))
                if len(words) >= 3:
                    search_variants.append(" ".join(words[:3]))
            
            # Try each variant and find the closest match to context
            best_instances = []
            best_distance = float('inf')
            
            for variant in search_variants:
                if not variant or len(variant.strip()) < 3:
                    continue
                
                try:
                    variant_instances = page.search_for(variant.strip(), quads=True)
                    if variant_instances:
                        # Find instance closest to context
                        for inst in variant_instances:
                            inst_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                            
                            # Check if instance is in search region
                            if search_region.intersects(inst_rect):
                                # Calculate distance to context
                                # Prefer instances on same line (y_diff should be small)
                                y_diff = abs(inst_rect.y0 - context_rect.y0)
                                # Prefer instances to the right of context (x_diff should be positive and reasonable)
                                if inst_rect.x0 > context_rect.x1:
                                    x_diff = inst_rect.x0 - context_rect.x1
                                elif inst_rect.x1 < context_rect.x0:
                                    x_diff = context_rect.x0 - inst_rect.x1
                                else:
                                    # Overlapping horizontally
                                    x_diff = 0
                                
                                # Weight y-difference more heavily (same line is most important)
                                # Allow some flexibility: same line or within 2 lines
                                if y_diff < 60:  # Within 2 lines
                                    distance = y_diff * 5 + x_diff
                                    
                                    if distance < best_distance:
                                        best_distance = distance
                                        best_instances = [inst]
                                    elif abs(distance - best_distance) < 100:  # Similar distance
                                        best_instances.append(inst)
                except Exception:
                    continue  # Skip if search fails
            
            if best_instances:
                # If multiple instances found, prefer the one closest to context
                if len(best_instances) > 1:
                    # Sort by distance and take the closest
                    best_instances.sort(key=lambda inst: (
                        abs((inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)).y0 - context_rect.y0),
                        abs((inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)).x0 - context_rect.x1)
                    ))
                instances = best_instances
    
    # Strategy 2: Direct search with word-by-word matching (if context search failed)
    if not instances and search_text:
        # Try full search_text
        try:
            instances = page.search_for(search_text, quads=True)
        except Exception:
            instances = None
        
        # If not found, try word-by-word (first few words)
        if not instances:
            words = search_text.split()
            if len(words) >= 5:
                # Try first 5 words
                snippet = " ".join(words[:5])
                try:
                    instances = page.search_for(snippet, quads=True)
                except Exception:
                    pass
            
            # If still not found, try first 3 words
            if not instances and len(words) >= 3:
                snippet = " ".join(words[:3])
                try:
                    instances = page.search_for(snippet, quads=True)
                except Exception:
                    pass
        
        # If still not found, try first sentence
        if not instances:
            sentences = search_text.split('. ')
            if sentences and len(sentences[0]) > 10:
                try:
                    instances = page.search_for(sentences[0].strip(), quads=True)
                except Exception:
                    pass
        
        # If still not found, try first 150 chars
        if not instances and len(search_text) > 150:
            try:
                instances = page.search_for(search_text[:150].strip(), quads=True)
            except Exception:
                pass
    
    # Strategy 3: Try full_text if different from search_text
    if not instances and full_text and full_text != search_text:
        try:
            instances = page.search_for(full_text, quads=True)
        except Exception:
            pass
        
        # If not found, try word-by-word
        if not instances:
            words = full_text.split()
            if len(words) >= 5:
                snippet = " ".join(words[:5])
                try:
                    instances = page.search_for(snippet, quads=True)
                except Exception:
                    pass
            if not instances and len(words) >= 3:
                snippet = " ".join(words[:3])
                try:
                    instances = page.search_for(snippet, quads=True)
                except Exception:
                    pass
    
    return instances


def create_summary_page(doc, old_file_info: dict, new_file_info: dict, changes: dict, page_changes: dict) -> None:
    """
    Create a summary page at the beginning of the PDF with file info and change metrics.
    
    Args:
        doc: PyMuPDF document object
        old_file_info: Dict with keys: name, pages, size_kb, timestamp
        new_file_info: Dict with keys: name, pages, size_kb, timestamp
        changes: Dict with keys: modified, added, removed
        page_changes: Dict mapping page_num to list of changes
    """
    import fitz
    from datetime import datetime
    
    # Create new page at the beginning
    summary_page = doc.new_page(0)  # Insert at position 0
    
    # Page dimensions
    page_rect = summary_page.rect
    width = page_rect.width
    height = page_rect.height
    margin = 50
    content_width = width - 2 * margin
    content_height = height - 2 * margin
    
    # Colors
    bg_color = (1, 1, 1)  # White background
    text_color = (0, 0, 0)  # Black text
    header_color = (0.2, 0.2, 0.2)  # Dark gray for headers
    accent_color = (0.4, 0.4, 0.4)  # Gray for accents
    
    # Fill background using shape
    shape = summary_page.new_shape()
    shape.draw_rect(fitz.Rect(0, 0, width, height))
    shape.finish(fill=bg_color, color=bg_color)
    shape.commit()
    
    # Title
    title_font_size = 24
    title_y = margin + 30
    title_rect = fitz.Rect(0, title_y - 15, width, title_y + 15)
    summary_page.insert_textbox(
        title_rect,
        "Compare Results",
        fontsize=title_font_size,
        color=header_color,
        align=1  # Center
    )
    
    # Date and time
    now = datetime.now()
    date_str = now.strftime("%d/%m/%Y %H:%M:%S")
    date_font_size = 10
    date_rect = fitz.Rect(0, margin, width, margin + 20)
    summary_page.insert_textbox(
        date_rect,
        date_str,
        fontsize=date_font_size,
        color=accent_color,
        align=1  # Center
    )
    
    # File comparison section
    file_section_y = title_y + 60
    file_section_height = 120
    
    # Old File
    old_file_x = margin + 20
    old_file_width = (content_width - 60) / 2
    
    summary_page.insert_text(
        (old_file_x, file_section_y),
        "Old File:",
        fontsize=12,
        color=header_color
    )
    
    # Old file name (bold simulation with larger font)
    old_name_y = file_section_y + 20
    old_name = old_file_info.get("name", "Unknown")
    summary_page.insert_text(
        (old_file_x, old_name_y),
        old_name,
        fontsize=11,
        color=text_color
    )
    
    # Old file details
    old_details_y = old_name_y + 18
    old_pages = old_file_info.get("pages", 0)
    old_size = old_file_info.get("size_kb", 0)
    old_details = f"{old_pages} pages ({old_size} KB)"
    summary_page.insert_text(
        (old_file_x, old_details_y),
        old_details,
        fontsize=10,
        color=accent_color
    )
    
    # Old file timestamp
    old_timestamp_y = old_details_y + 15
    old_timestamp = old_file_info.get("timestamp", "")
    if old_timestamp:
        summary_page.insert_text(
            (old_file_x, old_timestamp_y),
            old_timestamp,
            fontsize=9,
            color=accent_color
        )
    
    # "versus" text in center
    versus_y = file_section_y + 40
    versus_rect = fitz.Rect(0, versus_y - 5, width, versus_y + 15)
    summary_page.insert_textbox(
        versus_rect,
        "versus",
        fontsize=10,
        color=accent_color,
        align=1  # Center
    )
    
    # New File
    new_file_x = old_file_x + old_file_width + 40
    new_file_width = (content_width - 60) / 2
    
    summary_page.insert_text(
        (new_file_x, file_section_y),
        "New File:",
        fontsize=12,
        color=header_color
    )
    
    # New file name
    new_name_y = file_section_y + 20
    new_name = new_file_info.get("name", "Unknown")
    summary_page.insert_text(
        (new_file_x, new_name_y),
        new_name,
        fontsize=11,
        color=text_color
    )
    
    # New file details
    new_details_y = new_name_y + 18
    new_pages = new_file_info.get("pages", 0)
    new_size = new_file_info.get("size_kb", 0)
    new_details = f"{new_pages} pages ({new_size} KB)"
    summary_page.insert_text(
        (new_file_x, new_details_y),
        new_details,
        fontsize=10,
        color=accent_color
    )
    
    # New file timestamp
    new_timestamp_y = new_details_y + 15
    new_timestamp = new_file_info.get("timestamp", "")
    if new_timestamp:
        summary_page.insert_text(
            (new_file_x, new_timestamp_y),
            new_timestamp,
            fontsize=9,
            color=accent_color
        )
    
    # Change summary section
    summary_section_y = file_section_y + file_section_height + 40
    summary_section_height = 150
    
    # Total Changes
    total_changes_x = margin + 20
    total_changes_width = (content_width - 60) / 3
    
    summary_page.insert_text(
        (total_changes_x, summary_section_y),
        "Total Changes",
        fontsize=12,
        color=header_color
    )
    
    # Count total changes
    total_count = len(changes.get("modified", [])) + len(changes.get("added", [])) + len(changes.get("removed", []))
    total_count_y = summary_section_y + 30
    summary_page.insert_text(
        (total_changes_x, total_count_y),
        str(total_count),
        fontsize=32,
        color=text_color
    )
    
    # Comparison method
    method_y = total_count_y + 35
    summary_page.insert_text(
        (total_changes_x, method_y),
        "Text only comparison",
        fontsize=10,
        color=accent_color
    )
    
    # Content Changes
    content_x = total_changes_x + total_changes_width + 20
    content_width_section = (content_width - 60) / 3
    
    summary_page.insert_text(
        (content_x, summary_section_y),
        "Content",
        fontsize=12,
        color=header_color
    )
    
    # Counts
    modified_count = len(changes.get("modified", []))
    added_count = len(changes.get("added", []))
    removed_count = len(changes.get("removed", []))
    
    content_y = summary_section_y + 25
    summary_page.insert_text(
        (content_x, content_y),
        f"{modified_count} Replacements",
        fontsize=10,
        color=text_color
    )
    
    content_y += 18
    summary_page.insert_text(
        (content_x, content_y),
        f"{added_count} Insertions",
        fontsize=10,
        color=text_color
    )
    
    content_y += 18
    summary_page.insert_text(
        (content_x, content_y),
        f"{removed_count} Deletions",
        fontsize=10,
        color=text_color
    )
    
    # Styling and Annotations
    styling_x = content_x + content_width_section + 20
    styling_width_section = (content_width - 60) / 3
    
    summary_page.insert_text(
        (styling_x, summary_section_y),
        "Styling and Annotations",
        fontsize=12,
        color=header_color
    )
    
    styling_y = summary_section_y + 25
    summary_page.insert_text(
        (styling_x, styling_y),
        "0 Styling",
        fontsize=10,
        color=text_color
    )
    
    styling_y += 18
    summary_page.insert_text(
        (styling_x, styling_y),
        "0 Annotations",
        fontsize=10,
        color=text_color
    )
    
    # Find first page with changes for "Go to First Change" button
    first_change_page = None
    if page_changes:
        first_change_page = min(page_changes.keys())
    
    # Button area (if first change page exists)
    if first_change_page:
        button_y = summary_section_y + summary_section_height + 30
        button_text = f"Go to First Change (page {first_change_page})"
        # Draw a simple rectangle as button background
        button_rect = fitz.Rect(
            width / 2 - 150,
            button_y - 10,
            width / 2 + 150,
            button_y + 20
        )
        # Draw button background using shape
        button_shape = summary_page.new_shape()
        button_shape.draw_rect(button_rect)
        button_shape.finish(fill=(0.2, 0.4, 0.8), color=(0.2, 0.4, 0.8))
        button_shape.commit()
        button_text_rect = fitz.Rect(
            width / 2 - 150,
            button_y - 10,
            width / 2 + 150,
            button_y + 20
        )
        summary_page.insert_textbox(
            button_text_rect,
            button_text,
            fontsize=10,
            color=(1, 1, 1),  # White text
            align=1  # Center
        )


def create_highlighted_pdf_page_by_page(old_pdf_bytes: bytes, new_pdf_bytes: bytes, page_changes: dict, 
                                         old_file_info: dict = None, new_file_info: dict = None, 
                                         changes: dict = None) -> bytes:
    """
    Process PDF page by page: find old values, replace with new, and highlight with colors.
    Uses the larger PDF as the base. Handles pages that exist in one but not the other.
    
    Color coding:
    - RED: Deletions (text_deleted, removed content, pages only in old PDF)
    - GREEN: Additions (text_added, added content, pages only in new PDF)
    - YELLOW: Modifications (text_modified, numerical changes)
    
    Args:
        old_pdf_bytes: Bytes of the old PDF
        new_pdf_bytes: Bytes of the new PDF
        page_changes: Dict mapping page_num to list of changes
        old_file_info: Dict with file metadata (name, pages, size_kb, timestamp)
        new_file_info: Dict with file metadata (name, pages, size_kb, timestamp)
        changes: Dict with keys: modified, added, removed
    
    Returns:
        Bytes of the highlighted PDF
    """
    import fitz
    
    COLOR_MODIFIED = (1, 0.9, 0.4)  # Yellow for all changes
    
    # Open both PDFs
    old_doc = fitz.open(stream=old_pdf_bytes, filetype="pdf")
    new_doc = fitz.open(stream=new_pdf_bytes, filetype="pdf")
    
    old_page_count = len(old_doc)
    new_page_count = len(new_doc)
    
    # ALWAYS use NEW PDF as base so we can highlight all changes properly
    # This ensures we can see additions and modifications in the new version
    base_doc = new_doc
    base_page_count = new_page_count
    other_page_count = old_page_count
    
    # Create output document from NEW PDF
    output_doc = fitz.open()  # New empty document
    
    # Copy all pages from NEW document (base)
    output_doc.insert_pdf(base_doc)
    
    # Get all page numbers from both PDFs
    all_page_nums = set(range(1, base_page_count + 1)) | set(range(1, other_page_count + 1))
    
    # Process each page
    for page_num in sorted(all_page_nums):
        # PDF pages are 0-indexed
        page_idx = page_num - 1
        
        # Check if page exists in both PDFs
        page_in_new = page_idx < new_page_count
        page_in_old = page_idx < old_page_count
        
        # If page doesn't exist in new PDF, skip (shouldn't happen since we use new as base)
        if not page_in_new:
            continue
        
        # Get the page from output (it's from new PDF)
        if page_idx >= len(output_doc):
            continue
        
        page = output_doc[page_idx]
        
        # Page exists in both - process changes
        if page_num not in page_changes:
            continue
        
        changes_for_page = page_changes[page_num]
        
        # Track which instances have been used (to avoid duplicate highlighting)
        used_instances = set()
        
        # Process changes sequentially (in order from top to bottom)
        for change in changes_for_page:
            change_type = change.get("change_type", "numerical")
            
            # Determine color based on change type
            if change_type == "text_deleted":
                # For deleted text, we can't highlight it in the new PDF since it doesn't exist there
                # Skip highlighting deleted text - it's already been removed
                continue
            
            # Use yellow for all changes (text_added, text_modified, numerical)
            color = COLOR_MODIFIED
            
            if change_type in ["text_added", "text_modified"]:
                # CRITICAL: Use search_text from 2nd LLM (complete, correct text) instead of full_text (might be truncated)
                search_text = change.get("search_text", "").strip()
                full_text = str(change.get("new", "")).strip()
                
                # If no search_text from 2nd LLM, fall back to full_text
                if not search_text:
                    search_text = full_text
                
                context_before = change.get("context_before", "").strip()
                context_after = change.get("context_after", "").strip()
                row_label = change.get("row_label", "").strip()
                
                # Use the same text search logic as before
                instances = None
                # Try multiple search strategies (same as original function)
                if not instances:
                    instances = _find_text_block_by_extraction(
                        page,
                        search_text,
                        context_before=context_before,
                        context_after=context_after,
                        row_label=row_label
                    )
                if not instances:
                    # Use full_text for complete text block search (not just search_text snippet)
                    instances = _find_complete_text_block(
                        page,
                        full_text,  # Use full text, not just search snippet
                        context_before=context_before,
                        context_after=context_after,
                        row_label=row_label
                    )
                if not instances:
                    instances = _find_text_with_context(
                        page, 
                        search_text, 
                        context_before=context_before,
                        context_after=context_after,
                        row_label=row_label,
                        full_text=search_text
                    )
                if not instances:
                    # Try direct search
                    text_no_newlines = search_text.replace('\n', ' ').replace('\r', ' ')
                    text_no_newlines = re.sub(r'\s+', ' ', text_no_newlines).strip()
                    if len(text_no_newlines) <= 1000:
                        instances = page.search_for(text_no_newlines, quads=True)
            else:
                # For numerical changes, use STRICT context-based search
                # CRITICAL: For numerical changes, we need to find the EXACT instance
                # Strategy: Try to find OLD value first (if it still exists), then highlight NEW value at that location
                # OR: Use context to find where NEW value should be
                
                old_value = str(change.get("old", "")).strip()
                new_value = str(change.get("new", "")).strip()
                search_text = change.get("search_text", "").strip()
                if not search_text:
                    search_text = new_value  # Use new value as fallback
                
                if not search_text:
                    continue
                
                context_before = change.get("context_before", "").strip()
                context_after = change.get("context_after", "").strip()
                row_label = change.get("row_label", "").strip()
                
                # CRITICAL: For numerical values, ALWAYS use context if available
                # Don't fall back to simple search - it will pick wrong instances
                instances = None
                
                if row_label or context_before or context_after:
                    # Strategy 1: Try to find OLD value using context (most accurate)
                    # This finds the exact location where the change occurred
                    if old_value and old_value != "-" and old_value.lower() != "empty":
                        old_instances = _find_text_with_context(
                            page,
                            old_value,
                            context_before=context_before,
                            context_after=context_after,
                            row_label=row_label,
                            full_text=old_value
                        )
                        
                        if old_instances:
                            # Found old value - highlight NEW value at same location
                            # Use the old value's location but search for new value nearby
                            old_rect = old_instances[0].rect if hasattr(old_instances[0], 'rect') else fitz.Rect(old_instances[0])
                            
                            # Search for new value in a small region around old value location
                            search_region = fitz.Rect(
                                max(0, old_rect.x0 - 50),
                                max(0, old_rect.y0 - 10),
                                min(page.rect.width, old_rect.x1 + 50),
                                min(page.rect.height, old_rect.y1 + 10)
                            )
                            
                            # Try to find new value in this region
                            new_instances = page.search_for(new_value, quads=True)
                            if new_instances:
                                # Filter to instances in the search region
                                filtered = []
                                for inst in new_instances:
                                    inst_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                                    if search_region.intersects(inst_rect):
                                        filtered.append(inst)
                                
                                if filtered:
                                    instances = filtered
                    
                    # Strategy 2: If old value not found, use context to find new value
                    if not instances:
                        instances = _find_text_with_context(
                            page,
                            search_text,
                            context_before=context_before,
                            context_after=context_after,
                            row_label=row_label,
                            full_text=search_text
                        )
                
                # Strategy 3: If no context available, try to use surrounding text from change
                # This is a last resort - prefer to skip rather than highlight wrong instance
                if not instances:
                    # Try to build context from the change itself
                    section = change.get("section", "").strip()
                    if section:
                        # Try searching for section + new value
                        section_instances = page.search_for(section, quads=True)
                        if section_instances:
                            section_rect = section_instances[0].rect if hasattr(section_instances[0], 'rect') else fitz.Rect(section_instances[0])
                            search_region = fitz.Rect(
                                max(0, section_rect.x0 - 200),
                                max(0, section_rect.y0 - 50),
                                min(page.rect.width, section_rect.x1 + 500),
                                min(page.rect.height, section_rect.y1 + 200)
                            )
                            
                            new_instances = page.search_for(search_text, quads=True)
                            if new_instances:
                                filtered = []
                                for inst in new_instances:
                                    inst_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                                    if search_region.intersects(inst_rect):
                                        filtered.append(inst)
                                
                                if filtered:
                                    instances = filtered
                
                # LAST RESORT: Only if absolutely no context available
                # But this is risky - might highlight wrong instance
                # Better to skip than highlight incorrectly
                section = change.get("section", "").strip()
                if not instances and not (row_label or context_before or context_after or section):
                    # Only do simple search if we have NO context at all
                    # And prefer longer/more unique search text
                    if len(search_text) >= 3:  # At least 3 characters
                        all_instances = page.search_for(search_text, quads=True)
                        if all_instances and len(all_instances) == 1:
                            # Only use if there's exactly one instance (unique)
                            instances = all_instances
                        # If multiple instances, skip - too risky to guess
            
            # Highlight instances with appropriate color
            if instances:
                # Filter out already-used instances
                available_instances = []
                for quad in instances:
                    rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
                    pos_key = (round(rect.y0, 2), round(rect.x0, 2))
                    if pos_key not in used_instances:
                        available_instances.append((quad, pos_key))
                
                if available_instances:
                    # Sort by position
                    available_instances.sort(key=lambda x: (x[1][0], x[1][1]))
                    
                    # For text changes, try to highlight all consecutive instances
                    if change_type in ["text_added", "text_modified"]:
                        quads_only = [quad for quad, _ in available_instances]
                        if len(quads_only) > 1:
                            verified_quads = _verify_consecutive_instances(quads_only, page)
                            if verified_quads:
                                quads_only = verified_quads
                        
                        # If we have multiple instances, try to create a combined rectangle
                        # This ensures no gaps in highlighting
                        if len(quads_only) > 1:
                            try:
                                # Get all rectangles
                                rects = []
                                for quad in quads_only:
                                    rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
                                    rects.append(rect)
                                
                                # Create a combined rectangle that covers all instances
                                min_x = min(r.x0 for r in rects)
                                min_y = min(r.y0 for r in rects)
                                max_x = max(r.x1 for r in rects)
                                max_y = max(r.y1 for r in rects)
                                
                                # Create a combined quad/rectangle
                                combined_rect = fitz.Rect(min_x, min_y, max_x, max_y)
                                
                                # Check if this combined area hasn't been used
                                pos_key = (round(combined_rect.y0, 2), round(combined_rect.x0, 2))
                                if pos_key not in used_instances:
                                    # Create highlight annotation for the combined area
                                    annot = page.add_highlight_annot(combined_rect)
                                    annot.set_colors(stroke=color)
                                    annot.update()
                                    used_instances.add(pos_key)
                                    
                                    # Also mark individual instances as used to avoid duplicates
                                    for quad in quads_only:
                                        rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
                                        individual_key = (round(rect.y0, 2), round(rect.x0, 2))
                                        used_instances.add(individual_key)
                            except Exception as e:
                                print(f"Error creating combined highlight on page {page_num}: {e}")
                                # Fall back to individual highlighting
                                for quad in quads_only:
                                    try:
                                        rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
                                        pos_key = (round(rect.y0, 2), round(rect.x0, 2))
                                        if pos_key not in used_instances:
                                            annot = page.add_highlight_annot(quad)
                                            annot.set_colors(stroke=color)
                                            annot.update()
                                            used_instances.add(pos_key)
                                    except Exception as e2:
                                        print(f"Error highlighting on page {page_num}: {e2}")
                        else:
                            # Single instance - highlight normally
                            for quad in quads_only:
                                try:
                                    rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
                                    pos_key = (round(rect.y0, 2), round(rect.x0, 2))
                                    if pos_key not in used_instances:
                                        annot = page.add_highlight_annot(quad)
                                        annot.set_colors(stroke=color)
                                        annot.update()
                                        used_instances.add(pos_key)
                                except Exception as e:
                                    print(f"Error highlighting on page {page_num}: {e}")
                    else:
                        # For numerical changes, use the instance that best matches context
                        # If we have multiple instances, prefer the one that's closest to context
                        if len(available_instances) > 1 and (context_before or context_after or row_label):
                            # Find context location to help choose the right instance
                            context_rect = None
                            if row_label:
                                row_instances = page.search_for(row_label, quads=True)
                                if row_instances:
                                    context_rect = row_instances[0].rect if hasattr(row_instances[0], 'rect') else fitz.Rect(row_instances[0])
                            
                            if not context_rect and context_before:
                                context_words = context_before.split()
                                if context_words:
                                    context_snippet = " ".join(context_words[-3:]) if len(context_words) >= 3 else context_before
                                    before_instances = page.search_for(context_snippet, quads=True)
                                    if before_instances:
                                        context_rect = before_instances[0].rect if hasattr(before_instances[0], 'rect') else fitz.Rect(before_instances[0])
                            
                            if context_rect:
                                # Sort by distance to context (prefer same line, then closest horizontally)
                                available_instances.sort(key=lambda x: (
                                    abs((x[0].rect if hasattr(x[0], 'rect') else fitz.Rect(x[0])).y0 - context_rect.y0),
                                    abs((x[0].rect if hasattr(x[0], 'rect') else fitz.Rect(x[0])).x0 - context_rect.x1)
                                ))
                        
                        # Use the best instance (first after sorting)
                        quad, pos_key = available_instances[0]
                        try:
                            annot = page.add_highlight_annot(quad)
                            annot.set_colors(stroke=color)
                            annot.update()
                            used_instances.add(pos_key)
                        except Exception as e:
                            print(f"Error highlighting on page {page_num}: {e}")
            else:
                # If no instances found, try a simpler fallback search
                search_text = change.get("search_text", "").strip()
                if not search_text:
                    search_text = str(change.get("new", "")).strip()
                
                if search_text and len(search_text) > 2:
                    # Try simple direct search with cleaned text
                    try:
                        # Clean and limit search text
                        clean_search = search_text.replace('\n', ' ').replace('\r', ' ')
                        clean_search = re.sub(r'\s+', ' ', clean_search).strip()
                        if len(clean_search) > 100:
                            clean_search = clean_search[:100]
                        
                        fallback_instances = page.search_for(clean_search, quads=True)
                        if fallback_instances:
                            # Use first instance found
                            quad = fallback_instances[0]
                            rect = quad.rect if hasattr(quad, 'rect') else fitz.Rect(quad)
                            pos_key = (round(rect.y0, 2), round(rect.x0, 2))
                            if pos_key not in used_instances:
                                annot = page.add_highlight_annot(quad)
                                annot.set_colors(stroke=color)
                                annot.update()
                                used_instances.add(pos_key)
                    except Exception as e:
                        print(f"Error in fallback search on page {page_num}: {e}")
    
    # Close documents
    old_doc.close()
    new_doc.close()
    
    # Save to bytes
    output = io.BytesIO()
    output_doc.save(output)
    output_doc.close()
    
    return output.getvalue()


def create_highlighted_pdf_from_llm_changes(pdf_bytes: bytes, changes: dict) -> bytes:
    """Create a highlighted PDF based on LLM-detected changes."""
    import fitz
    
    COLOR_MODIFIED = (1, 0.9, 0.4)     # Yellow for all changes
    
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Highlight added content (yellow)
    for item in changes.get("added", []):
        text = item.get("text", "")
        if text:
            highlight_with_context(doc, text, "", COLOR_MODIFIED)
    
    # Highlight modified content (yellow) - highlight ONLY the changed values
    for item in changes.get("modified", []):
        old_val = item.get("old", "")
        new_val = item.get("new", "")
        field = item.get("field", "")
        
        if not old_val or not new_val:
            continue
        
        # Extract only the values that actually changed
        changed_values = extract_changed_values(old_val, new_val)
        
        if changed_values:
            # Highlight each changed value individually
            for value in changed_values:
                # Clean the value (remove watermark artifacts)
                clean_value = re.sub(r'\b[TFARD]\b', '', str(value)).strip()
                clean_value = re.sub(r'\s+', ' ', clean_value)
                
                if clean_value and len(clean_value) >= 2:
                    # Try highlighting with context
                    highlighted = highlight_with_context(doc, clean_value, field, COLOR_MODIFIED)
                    
                    # If that failed, try just the numeric part
                    if not highlighted:
                        numeric = re.search(r'[\d,()]+', clean_value)
                        if numeric:
                            highlight_with_context(doc, numeric.group(0), field, COLOR_MODIFIED)
        else:
            # Fallback: try to highlight the new value as-is (cleaned)
            highlight_with_context(doc, new_val, field, COLOR_MODIFIED)
    
    # Save to bytes
    output = io.BytesIO()
    doc.save(output)
    doc.close()
    
    return output.getvalue()


def render_marker_changes(changes: dict):
    """Render marker PDF comparison changes in the UI."""
    if not changes:
        st.info("No changes detected.")
        return
    
    # Display changes by type
    if changes.get("modified"):
        st.markdown("### 🔄 Modified Content")
        for item in changes["modified"]:
            field = item.get("field", "Unknown")
            old_val = item.get("old", "")
            new_val = item.get("new", "")
            st.markdown(f"**{field}:**")
            st.markdown(f"- Old: `{old_val[:200]}{'...' if len(old_val) > 200 else ''}`")
            st.markdown(f"- New: `{new_val[:200]}{'...' if len(new_val) > 200 else ''}`")
            st.markdown("---")
    
    if changes.get("added"):
        st.markdown("### ➕ Added Content")
        for item in changes["added"]:
            text = item.get("text", "")
            st.markdown(f"`{text[:200]}{'...' if len(text) > 200 else ''}`")
            st.markdown("---")
    
    if changes.get("removed"):
        st.markdown("### ➖ Removed Content")
        for item in changes["removed"]:
            text = item.get("text", "")
            st.markdown(f"`{text[:200]}{'...' if len(text) > 200 else ''}`")
            st.markdown("---")
    
    if not (changes.get("modified") or changes.get("added") or 
            changes.get("removed")):
        st.info("No significant changes detected between the documents.")


def create_highlighted_pdf_from_llm_changes(pdf_bytes: bytes, changes: dict) -> bytes:
    """Create a highlighted PDF based on LLM-detected changes."""
    import fitz
    
    COLOR_MODIFIED = (1, 0.9, 0.4)     # Yellow for all changes
    
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Highlight added content (yellow)
    for item in changes.get("added", []):
        text = item.get("text", "")
        if text:
            highlight_with_context(doc, text, "", COLOR_MODIFIED)
    
    # Highlight modified content (yellow) - highlight ONLY the changed values
    for item in changes.get("modified", []):
        old_val = item.get("old", "")
        new_val = item.get("new", "")
        field = item.get("field", "")
        
        if not old_val or not new_val:
            continue
        
        # Extract only the values that actually changed
        changed_values = extract_changed_values(old_val, new_val)
        
        if changed_values:
            # Highlight each changed value individually
            for value in changed_values:
                # Clean the value (remove watermark artifacts)
                clean_value = re.sub(r'\b[TFARD]\b', '', str(value)).strip()
                clean_value = re.sub(r'\s+', ' ', clean_value)
                
                if clean_value and len(clean_value) >= 2:
                    # Try highlighting with context
                    highlighted = highlight_with_context(doc, clean_value, field, COLOR_MODIFIED)
                    
                    # If that failed, try just the numeric part
                    if not highlighted:
                        numeric = re.search(r'[\d,()]+', clean_value)
                        if numeric:
                            highlight_with_context(doc, numeric.group(0), field, COLOR_MODIFIED)
        else:
            # Fallback: try to highlight the new value as-is (cleaned)
            highlight_with_context(doc, new_val, field, COLOR_MODIFIED)
    
    # Save to bytes
    output = io.BytesIO()
    doc.save(output)
    doc.close()
    
    return output.getvalue()


def render_marker_changes(changes: dict):
    """Render changes from Marker comparison."""
    with st.expander("📋 View Detailed Changes", expanded=True):
        
        if changes.get("modified"):
            st.markdown("### 🟡 Modified Values")
            for item in changes["modified"]:
                field = item.get("field", "Item")
                old_val = item.get("old", "")
                new_val = item.get("new", "")
                st.markdown(f"- **{field}**: `{old_val}` → `{new_val}`")
        
        if changes.get("added"):
            st.markdown("### 🟢 Added Content")
            for item in changes["added"]:
                text = item.get("text", "")
                section = item.get("section", "")
                if section:
                    st.markdown(f"- **[{section}]** {text[:200]}{'...' if len(str(text)) > 200 else ''}")
                else:
                    st.markdown(f"- {text[:200]}{'...' if len(str(text)) > 200 else ''}")
        
        if changes.get("removed"):
            st.markdown("### 🔴 Removed Content")
            for item in changes["removed"]:
                text = item.get("text", "")
                section = item.get("section", "")
                if section:
                    st.markdown(f"- **[{section}]** {text[:200]}{'...' if len(str(text)) > 200 else ''}")
                else:
                    st.markdown(f"- {text[:200]}{'...' if len(str(text)) > 200 else ''}")
        
        if not any([changes.get("modified"), changes.get("added"), changes.get("removed")]):
            st.info("No significant changes detected between the documents.")


def create_highlighted_pdf_from_changes(pdf_bytes: bytes, aligned_units: list, llm_payloads: list) -> bytes:
    """Create a highlighted PDF based on detected changes."""
    import fitz
    
    # Colors for highlighting (RGB normalized to 0-1)
    COLOR_MODIFIED = (1, 0.9, 0.4)     # Yellow for all changes
    
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Collect texts to highlight
    added_texts = []
    modified_texts = []
    
    for payload in llm_payloads:
        change_type = payload.get("change_type", "")
        new_content = payload.get("new", "")
        
        if change_type == "added" and new_content and new_content != "[Not present in old version]":
            added_texts.append(new_content)
        elif change_type in ("matched", "modified") and new_content and new_content != "[Not present in new version]":
            modified_texts.append(new_content)
    
    # Also check aligned_units for direct text
    for unit in aligned_units:
        status = unit.get("status", "")
        new_content = unit.get("new")
        
        if isinstance(new_content, str) and new_content:
            if status == "added":
                added_texts.append(new_content)
            elif status == "matched":
                modified_texts.append(new_content)
    
    # Highlight added content (yellow)
    for text in added_texts:
        highlight_text_in_doc(doc, text, COLOR_MODIFIED)
    
    # Highlight modified content (yellow)
    for text in modified_texts:
        highlight_text_in_doc(doc, text, COLOR_MODIFIED)
    
    # Save to bytes
    output = io.BytesIO()
    doc.save(output)
    doc.close()
    
    return output.getvalue()


def highlight_text_in_doc(doc, text: str, color: tuple) -> bool:
    """Highlight text in a PDF document."""
    import fitz
    
    if not text or len(text.strip()) < 3:
        return False
    
    text = text.strip()
    
    # Skip very long texts (they won't match exactly)
    if len(text) > 500:
        # Try to highlight first sentence or significant portion
        sentences = text.split('. ')
        if sentences:
            text = sentences[0][:200]
    
    # Try to find and highlight in each page
    highlighted = False
    for page in doc:
        # Search for the text
        instances = page.search_for(text[:100], quads=True)  # Limit search length
        
        if instances:
            try:
                annot = page.add_highlight_annot(instances[0])
                annot.set_colors(stroke=color)
                annot.update()
                highlighted = True
                break  # Only highlight first occurrence
            except Exception:
                pass
        
        # If exact match fails, try shorter snippets
        if not highlighted and len(text) > 50:
            words = text.split()
            if len(words) > 5:
                # Try first few words
                snippet = ' '.join(words[:5])
                instances = page.search_for(snippet, quads=True)
                if instances:
                    try:
                        annot = page.add_highlight_annot(instances[0])
                        annot.set_colors(stroke=color)
                        annot.update()
                        highlighted = True
                        break
                    except Exception:
                        pass
    
    return highlighted


def render_pdf_changes(changes: dict):
    """Render PDF changes in expandable format."""
    with st.expander("📋 View Detailed Changes", expanded=True):
                    if changes.get("removed"):
                        st.markdown("**🔴 Removed:**")
                        for item in changes["removed"]:
                            if isinstance(item, dict):
                                field = item.get("field", "")
                                value = item.get("value", item.get("text", ""))
                                context = item.get("context", "")
                                if field:
                                    st.markdown(f"- **{field}**: {value}" + (f" *(near: {context})*" if context else ""))
                                else:
                                    display_text = str(value) if value else str(item)
                                    st.markdown(f"- {display_text[:200]}..." if len(display_text) > 200 else f"- {display_text}")
                            else:
                                display_text = str(item)
                                st.markdown(f"- {display_text[:200]}..." if len(display_text) > 200 else f"- {display_text}")
                    
                    if changes.get("added"):
                        st.markdown("**🟢 Added:**")
                        for item in changes["added"]:
                            if isinstance(item, dict):
                                field = item.get("field", "")
                                value = item.get("value", item.get("text", ""))
                                context = item.get("context", "")
                                if field:
                                    st.markdown(f"- **{field}**: {value}" + (f" *(near: {context})*" if context else ""))
                                else:
                                    display_text = str(value) if value else str(item)
                                    st.markdown(f"- {display_text[:200]}..." if len(display_text) > 200 else f"- {display_text}")
                            else:
                                display_text = str(item)
                                st.markdown(f"- {display_text[:200]}..." if len(display_text) > 200 else f"- {display_text}")
                    
                    if changes.get("modified"):
                        st.markdown("**🟡 Modified:**")
                        for item in changes["modified"]:
                            if isinstance(item, dict):
                                field = item.get("field", "Item")
                                old_val = str(item.get("old", item.get("old_value", "")))
                                new_val = str(item.get("new", item.get("new_value", "")))
                                diff = item.get("difference", "")
                                
                                st.markdown(f"- **{field}**: `{old_val}` → `{new_val}`" + (f" *({diff})*" if diff else ""))
                            else:
                                st.markdown(f"- {item}")


def render_excel_csv_comparison():
    """Render Excel/CSV comparison interface."""
    st.markdown("### 📊 Excel / CSV Version Comparison")
    st.markdown("Upload two files to compare. A single output file will be generated showing all changes.")
    
    # File type selector
    file_type = st.radio(
        "Select file type",
        ["Excel (.xlsx)", "CSV (.csv)"],
        horizontal=True,
        key="tabular_file_type"
    )
    
    accepted_type = ["xlsx", "xls"] if "Excel" in file_type else ["csv"]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="upload-card">
            <h3>📂 Older Version</h3>
        </div>
        """, unsafe_allow_html=True)
        old_file = st.file_uploader(
            f"Upload older {file_type.split()[0]}",
            type=accepted_type,
            key="tabular_old",
            help=f"Select the older version of the {file_type.split()[0]} file (Max size: {Config.MAX_FILE_SIZE_MB} MB)"
        )
        st.caption(f"⚠️ Maximum file size: {Config.MAX_FILE_SIZE_MB} MB")
        if old_file and not validate_file_size(old_file):
            st.error(f"File too large. Max size: {Config.MAX_FILE_SIZE_MB} MB")
            old_file = None
    
    with col2:
        st.markdown("""
        <div class="upload-card">
            <h3>📂 Newer Version</h3>
        </div>
        """, unsafe_allow_html=True)
        new_file = st.file_uploader(
            f"Upload newer {file_type.split()[0]}",
            type=accepted_type,
            key="tabular_new",
            help=f"Select the newer version of the {file_type.split()[0]} file (Max size: {Config.MAX_FILE_SIZE_MB} MB)"
        )
        st.caption(f"⚠️ Maximum file size: {Config.MAX_FILE_SIZE_MB} MB")
        if new_file and not validate_file_size(new_file):
            st.error(f"File too large. Max size: {Config.MAX_FILE_SIZE_MB} MB")
            new_file = None
    
    # Comparison button
    st.markdown("")
    
    if st.button("🔍 Compare Files", key="compare_tabular", use_container_width=True):
        if not old_file or not new_file:
            st.warning("Please upload both files to compare.")
            return
        
        with st.spinner("🔄 Comparing files row by row..."):
            try:
                comparator = ExcelCSVComparator()
                
                # Read file bytes
                old_bytes = old_file.read()
                new_bytes = new_file.read()
                
                # Perform comparison based on file type
                if "Excel" in file_type:
                    result = comparator.compare_excel_files(old_bytes, new_bytes)
                    output_filename = f"comparison_{new_file.name}"
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                else:
                    result = comparator.compare_csv_files(old_bytes, new_bytes)
                    # CSV comparison now outputs Excel with 3 sheets
                    output_filename = f"comparison_{new_file.name.rsplit('.', 1)[0]}.xlsx"
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                
                st.success("✅ Comparison complete!")
                
                # Display summary
                st.info(f"**Summary:** {result['summary']}")
                
                # Download button
                st.markdown("### 📥 Download Comparison Result")
                st.download_button(
                    label="⬇️ Download Comparison File",
                    data=result["result_bytes"],
                    file_name=output_filename,
                    mime=mime_type,
                    use_container_width=True
                )
                
                # Preview - show result with Change column
                st.markdown("### 👀 Preview")
                
                if result["file_type"] == "excel":
                    # Show tabs for each sheet
                    sheets = result["result_sheets"]
                    if sheets:
                        sheet_tabs = st.tabs(list(sheets.keys()))
                        for tab, (sheet_name, result_df) in zip(sheet_tabs, sheets.items()):
                            with tab:
                                preview_df = comparator.get_preview_df(result_df)
                                st.dataframe(
                                    preview_df,
                                    use_container_width=True,
                                    height=400
                                )
                                if len(result_df) > 20:
                                    st.caption(f"Showing first 20 of {len(result_df)} rows")
                else:
                    # CSV - show result with Change column
                    result_df = result["result_df"]
                    preview_df = comparator.get_preview_df(result_df)
                    st.dataframe(
                        preview_df,
                        use_container_width=True,
                        height=400
                    )
                    if len(result_df) > 20:
                        st.caption(f"Showing first 20 of {len(result_df)} rows")
                
            except Exception as e:
                st.error(f"Error during comparison: {str(e)}")


def main():
    """Main application entry point."""
    render_header()
    render_sidebar()
    
    # Main content tabs
    tab1, tab2 = st.tabs(["📄 PDF Comparison", "📊 Excel / CSV Comparison"])
    
    with tab1:
        render_pdf_comparison()
    
    with tab2:
        render_excel_csv_comparison()
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<p style='text-align: center; color: rgba(255,255,255,0.4); font-size: 0.85rem;'>"
        "Version Comparison Agent • Powered by Azure OpenAI GPT-4.1"
        "</p>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()


