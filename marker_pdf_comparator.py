"""
PDF Comparison Pipeline with multiple PDF-to-Markdown options.
Tries libraries in order: pdfplumber > PyMuPDF4LLM
"""
import os
import re
import tempfile
from config import Config
from llm_client import LLMClient

# Track available libraries
AVAILABLE_LIBRARIES = []
ACTIVE_LIBRARY = None

# Try pdfplumber
try:
    import pdfplumber
    AVAILABLE_LIBRARIES.append("pdfplumber")
    print("[PDF2MD] pdfplumber loaded successfully")
except ImportError as e:
    print(f"[PDF2MD] pdfplumber not available: {e}")

# Try PyMuPDF4LLM (fallback)
try:
    import pymupdf4llm
    AVAILABLE_LIBRARIES.append("PyMuPDF4LLM")
    print("[PDF2MD] PyMuPDF4LLM loaded successfully")
except ImportError as e:
    print(f"[PDF2MD] PyMuPDF4LLM not available: {e}")

# Set active library (first available)
if AVAILABLE_LIBRARIES:
    ACTIVE_LIBRARY = AVAILABLE_LIBRARIES[0]
    print(f"[PDF2MD] Using: {ACTIVE_LIBRARY}")


def convert_pdf_with_pdfplumber(pdf_path: str) -> str:
    """Convert PDF using pdfplumber with table extraction."""
    import pdfplumber
    
    markdown_parts = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            markdown_parts.append(f"\n## Page {page_num}\n")
            
            # Extract tables first
            tables = page.extract_tables()
            table_bboxes = []
            
            if tables:
                for table in tables:
                    if table and len(table) > 0:
                        # Convert table to markdown
                        md_table = []
                        for i, row in enumerate(table):
                            # Clean cells
                            clean_row = [str(cell).strip() if cell else "" for cell in row]
                            md_table.append("| " + " | ".join(clean_row) + " |")
                            if i == 0:
                                # Add header separator
                                md_table.append("|" + "|".join(["---"] * len(clean_row)) + "|")
                        markdown_parts.append("\n".join(md_table))
                        markdown_parts.append("")
            
            # Extract text
            text = page.extract_text()
            if text:
                markdown_parts.append(text)
    
    return "\n\n".join(markdown_parts)


def extract_pages_from_markdown(markdown: str) -> dict:
    """Extract markdown content page by page. Returns dict: {page_num: markdown_content}"""
    pages = {}
    
    if not markdown:
        return pages
    
    # Split by page markers (## Page N)
    page_sections = re.split(r'## Page (\d+)', markdown)
    
    # Handle content before first page marker (if any)
    if page_sections and page_sections[0].strip():
        pages[1] = page_sections[0].strip()
    
    # Process page markers and their content
    for i in range(1, len(page_sections), 2):
        if i + 1 < len(page_sections):
            try:
                page_num = int(page_sections[i])
                page_content = page_sections[i + 1].strip()
                if page_content:
                    pages[page_num] = page_content
            except (ValueError, IndexError):
                continue
    
    return pages


def convert_pdf_with_pymupdf4llm(pdf_path: str) -> str:
    """Convert PDF using PyMuPDF4LLM."""
    return pymupdf4llm.to_markdown(pdf_path)


def convert_pdf_to_markdown(pdf_path: str, library: str = None) -> str:
    """
    Convert PDF to Markdown using specified or default library.
    """
    lib = library or ACTIVE_LIBRARY
    
    if not lib:
        raise ImportError("No PDF-to-Markdown library available. Install: pip install pdfplumber pymupdf4llm")
    
    print(f"[PDF2MD] Converting with {lib}: {pdf_path}")
    
    if lib == "pdfplumber":
        markdown = convert_pdf_with_pdfplumber(pdf_path)
    elif lib == "PyMuPDF4LLM":
        markdown = convert_pdf_with_pymupdf4llm(pdf_path)
    else:
        raise ValueError(f"Unknown library: {lib}")
    
    print(f"[PDF2MD] Conversion complete: {len(markdown)} characters")
    return markdown


def convert_pdf_bytes_to_markdown(pdf_bytes: bytes, library: str = None) -> str:
    """Convert PDF bytes to Markdown."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    
    try:
        markdown = convert_pdf_to_markdown(tmp_path, library)
    finally:
        os.unlink(tmp_path)
    
    return markdown


def normalize_markdown(markdown: str) -> str:
    """Clean up markdown while preserving structure."""
    if not markdown:
        return ""
    
    # Remove DRAFT watermarks
    markdown = re.sub(r'\bDRAFT\b', '', markdown, flags=re.IGNORECASE)
    
    # Remove standalone page numbers
    markdown = re.sub(r'^\s*\d+\s*$', '', markdown, flags=re.MULTILINE)
    
    # Remove excessive separators
    markdown = re.sub(r'-{5,}', '---', markdown)
    
    # Normalize excessive whitespace
    markdown = re.sub(r'\n{4,}', '\n\n\n', markdown)
    
    # Clean up table formatting
    markdown = re.sub(r'\|\s+', '| ', markdown)
    markdown = re.sub(r'\s+\|', ' |', markdown)
    
    return markdown.strip()


def compare_markdown_page_with_llm(old_page_markdown: str, new_page_markdown: str, page_num: int) -> dict:
    """Compare a single page's markdown content and get old/new value pairs."""
    llm_client = LLMClient()
    
    system_prompt = """You are a document comparison expert. Compare two versions of a SINGLE PAGE and identify ALL changes.

## CHANGE TYPES

### TYPE 1: NUMERICAL CHANGES (tables, financial data)
- Individual numbers, amounts, percentages, dates
- Extract ONE value per change
- Example: "370,308", "(111,125)", "31.59%", "31 March 2025"

### TYPE 2: TEXTUAL CHANGES (paragraphs, sentences, phrases)
- Added text: Entire paragraph/sentence that appears ONLY in NEW
- Deleted text: Entire paragraph/sentence that appears ONLY in OLD
- Modified text: Changed words/phrases within existing text
- Extract COMPLETE text blocks (entire sentences/paragraphs) for highlighting

## OUTPUT FORMAT (strict JSON)
{
  "changes": [
    {
      "old": "exact old value or empty string if added",
      "new": "exact new value or empty string if deleted",
      "change_type": "numerical|text_added|text_deleted|text_modified",
      "context": "brief description",
      "surrounding_text_before": "2-5 words before (for numerical) or paragraph before (for text)",
      "surrounding_text_after": "2-5 words after (for numerical) or paragraph after (for text)",
      "section": "section/table name",
      "row_label": "row label if in table, else empty",
      "position_hint": "unique identifier for repeated values"
    }
  ]
}

## RULES - NUMERICAL CHANGES
- Extract ONE numerical value per change (numbers, amounts, percentages, dates)
- **CRITICAL: Numerical changes contain ONLY numbers - NO text, NO words, NO labels**
- If a line has multiple values, create separate entries for EACH changed value
- Return EXACT text as it appears (commas, parentheses, formatting)
- If same value appears multiple times, return ONE entry per instance with unique position_hint
- Provide context: surrounding_text_before/after (2-5 words), section, row_label
- **Example**: If "Debtors 5 2,840,839" changes to "Debtors 6 2,841,188":
  - ❌ WRONG: {"old": "Debtors 5 2,840,839", "new": "Debtors 6 2,841,188"} (mixes text and numbers)
  - ✅ CORRECT: Three separate entries (text and numerical must NEVER be mixed):
    - Entry 1: {"old": "Debtors", "new": "Debtors", "change_type": "text_modified", ...} (text label)
    - Entry 2: {"old": "5", "new": "6", "change_type": "numerical", ...} (note number)
    - Entry 3: {"old": "2,840,839", "new": "2,841,188", "change_type": "numerical", ...} (value)

## RULES - TEXTUAL CHANGES

### TEXT ADDED (appears only in NEW):
- `change_type`: "text_added"
- `old`: "" (empty string)
- `new`: COMPLETE added text (entire sentence/paragraph)
- **CRITICAL: Text changes contain ONLY words/text - NO numbers, NO amounts**
- Capture the FULL text block for highlighting
- Example: If a new paragraph was added, include the ENTIRE paragraph

### TEXT DELETED (appears only in OLD):
- `change_type`: "text_deleted"
- `old`: COMPLETE deleted text (entire sentence/paragraph)
- `new`: "" (empty string)
- **CRITICAL: Text changes contain ONLY words/text - NO numbers, NO amounts**
- Capture the FULL text block that was removed
- Example: If a paragraph was deleted, include the ENTIRE paragraph

### TEXT MODIFIED (changed within existing text):
- `change_type`: "text_modified"
- `old`: Original text (can be phrase, sentence, or paragraph)
- `new`: New text (can be phrase, sentence, or paragraph)
- **CRITICAL: Text changes contain ONLY words/text - NO numbers, NO amounts**
- If text contains numbers that changed, extract ONLY the text part, create separate numerical entry for numbers
- Capture COMPLETE changed text blocks, not just individual words
- Example: If a sentence was rewritten, include the FULL old and new sentences
- **Example**: If "Tangible assets 4" changes to "Tangible assets 5":
  - ❌ WRONG: {"old": "Tangible assets 4", "new": "Tangible assets 5", "change_type": "text_modified"} (contains number)
  - ✅ CORRECT: {"old": "4", "new": "5", "change_type": "numerical", ...} (just the number)

## GENERAL RULES
- Return changes in ORDER from TOP to BOTTOM
- Only include actual changes (old ≠ new)
- Provide MAXIMUM context for unique identification
- **CRITICAL SEPARATION RULE:**
  - **Text changes (text_added, text_deleted, text_modified) must contain ONLY words/text - NO numbers**
  - **Numerical changes must contain ONLY numbers/amounts - NO text, NO words, NO labels**
  - **If a line has both text and numerical changes, create SEPARATE entries for each**
  - **Example**: "Debtors 5 2,840,839" → "Debtors 6 2,841,188" should be:
    - Entry 1: {"old": "Debtors", "new": "Debtors", "change_type": "text_modified", ...} (text label)
    - Entry 2: {"old": "5", "new": "6", "change_type": "numerical", ...} (note number)
    - Entry 3: {"old": "2,840,839", "new": "2,841,188", "change_type": "numerical", ...} (value)
    - NOT: {"old": "Debtors 5 2,840,839", "new": "Debtors 6 2,841,188", ...} (mixed)

## CRITICAL: IGNORE THESE ELEMENTS (DO NOT REPORT AS CHANGES)
- **Watermarks**: Standalone letters (T, F, A, R, D) scattered in text - remove from values
- **Page numbers**: "Page 1", "Page 2", "## Page 1", "-1-", "-2-", etc.
- **Headers/Footers**: Text that repeats on every page (company name, document title, etc.)
- **DRAFT stamps**: "DRAFT", "CONFIDENTIAL", "FOR REVIEW", etc.
- **Page markers**: "## Page N", section dividers that are just formatting
- **Table separators**: "|---|---|", decorative lines
- **Formatting artifacts**: Extra spaces, line breaks that don't affect meaning
- **Repeated boilerplate**: Standard disclaimers, legal text that appears in both versions unchanged

**If any of these elements appear to change, IGNORE them - they are not meaningful content changes.**

## EXAMPLES

### NUMERICAL CHANGE (pure number):
OLD: "D (995,244) (786,436)" → NEW: "D (624,936) (786,436)"
{
  "old": "(995,244)",
  "new": "(624,936)",
  "change_type": "numerical",
  "context": "Cost of sales 2025",
  "surrounding_text_before": "Closing valuation D",
  "surrounding_text_after": "(786,436)",
  "section": "Profit and Loss Account",
  "row_label": "Cost of sales",
  "position_hint": ""
}

### MIXED CHANGE - MUST SPLIT (text label + note number + value):
OLD: "Debtors 5 2,840,839 3,037,014" → NEW: "Debtors 6 2,841,188 3,037,014"
Return THREE separate entries (text and numerical must NEVER be mixed):

Entry 1 (text label "Debtors"):
{
  "old": "Debtors",
  "new": "Debtors",
  "change_type": "text_modified",
  "context": "Debtors label",
  "surrounding_text_before": "",
  "surrounding_text_after": "6",
  "section": "STATEMENT OF FINANCIAL POSITION",
  "row_label": "Debtors",
  "position_hint": ""
}
(Note: Even if text label didn't change, create a separate entry to show it's part of the line - text and numerical must be separated)

Entry 2 (note number - numerical):
{
  "old": "5",
  "new": "6",
  "change_type": "numerical",
  "context": "Note number for Debtors",
  "surrounding_text_before": "Debtors",
  "surrounding_text_after": "2,841,188",
  "section": "STATEMENT OF FINANCIAL POSITION",
  "row_label": "Debtors",
  "position_hint": ""
}

Entry 3 (value - numerical):
{
  "old": "2,840,839",
  "new": "2,841,188",
  "change_type": "numerical",
  "context": "Debtors 2024 value",
  "surrounding_text_before": "Debtors 6",
  "surrounding_text_after": "3,037,014",
  "section": "STATEMENT OF FINANCIAL POSITION",
  "row_label": "Debtors",
  "position_hint": ""
}

❌ WRONG (mixed - text and numerical combined):
{"old": "Debtors 5 2,840,839", "new": "Debtors 6 2,841,188", "change_type": "text_modified"}
{"old": "5 2,840,839", "new": "6 2,841,188", "change_type": "numerical"}

### TEXT ADDED (new paragraph):
NEW has: "...previous text. This is a new paragraph that was added. It contains important information about the changes. Next text..."
{
  "old": "",
  "new": "This is a new paragraph that was added. It contains important information about the changes.",
  "change_type": "text_added",
  "context": "New paragraph added",
  "surrounding_text_before": "previous text",
  "surrounding_text_after": "Next text",
  "section": "Notes to Accounts",
  "row_label": "",
  "position_hint": ""
}

### TEXT DELETED (removed paragraph):
OLD had: "...previous text. This paragraph was removed entirely. It contained old information. Next text..."
NEW has: "...previous text. Next text..."
{
  "old": "This paragraph was removed entirely. It contained old information.",
  "new": "",
  "change_type": "text_deleted",
  "context": "Paragraph deleted",
  "surrounding_text_before": "previous text",
  "surrounding_text_after": "Next text",
  "section": "Notes to Accounts",
  "row_label": "",
  "position_hint": ""
}

### TEXT MODIFIED (sentence changed):
OLD: "The company reported a loss for the year."
NEW: "The company reported a profit for the year."
{
  "old": "The company reported a loss for the year.",
  "new": "The company reported a profit for the year.",
  "change_type": "text_modified",
  "context": "Result description changed",
  "surrounding_text_before": "previous sentence",
  "surrounding_text_after": "next sentence",
  "section": "Notes to Accounts",
  "row_label": "",
  "position_hint": ""
}

### REPEATED NUMERICAL VALUE:
"195" → "193" appears TWICE:
First: {"old": "195", "new": "193", "change_type": "numerical", ..., "position_hint": "first instance"}
Second: {"old": "195", "new": "193", "change_type": "numerical", ..., "position_hint": "second instance"}

## CRITICAL FOR TEXT CHANGES
- ✅ Capture COMPLETE sentences/paragraphs, not fragments
- ✅ For added text: Include the ENTIRE new text block
- ✅ For deleted text: Include the ENTIRE removed text block
- ✅ For modified text: Include FULL old and new versions
- ❌ Do NOT capture just individual words - capture complete text blocks
- **CRITICAL: Text changes must contain ONLY words/text - NO numbers. If text contains numbers, extract numbers separately as numerical changes.**"""

    # Use string concatenation to avoid f-string format errors with curly braces in markdown
    user_prompt = """Compare these two versions of PAGE """ + str(page_num) + """ and identify ALL changes.

=== OLD VERSION (PAGE """ + str(page_num) + """) ===
""" + old_page_markdown + """

=== NEW VERSION (PAGE """ + str(page_num) + """) ===
""" + new_page_markdown + """

## PROCESSING INSTRUCTIONS

### FOR NUMERICAL CHANGES (tables, financial data):
1. Extract ONE numerical value per change (numbers, amounts, percentages, dates)
2. **CRITICAL: Numerical changes contain ONLY numbers - NO text, NO words, NO labels**
3. If a line has multiple values, create separate entries for EACH changed value
4. Provide context: surrounding_text_before/after (2-5 words), section, row_label
5. **If a line has both text and numbers that changed, create SEPARATE entries (text and numerical must NEVER be mixed):**
   - Example: "Debtors 5 2,840,839" → "Debtors 6 2,841,188"
   - Create: Entry 1 for "Debtors" (text label)
   - Create: Entry 2 for "5" → "6" (note number - numerical)
   - Create: Entry 3 for "2,840,839" → "2,841,188" (value - numerical)
   - NOT: One entry mixing text and numbers

### FOR TEXTUAL CHANGES (paragraphs, sentences):
1. **TEXT ADDED**: If text appears ONLY in NEW:
   - Set change_type: "text_added"
   - old: "" (empty)
   - new: COMPLETE added text (entire sentence/paragraph)
   - Capture the FULL text block for highlighting

2. **TEXT DELETED**: If text appears ONLY in OLD:
   - Set change_type: "text_deleted"
   - old: COMPLETE deleted text (entire sentence/paragraph)
   - new: "" (empty)
   - Capture the FULL removed text block

3. **TEXT MODIFIED**: If text was changed:
   - Set change_type: "text_modified"
   - old: COMPLETE original text (sentence/paragraph)
   - new: COMPLETE new text (sentence/paragraph)
   - **CRITICAL: Text changes contain ONLY words/text - NO numbers, NO amounts**
   - If text contains numbers that changed, extract ONLY the text part, create separate numerical entry for numbers
   - Capture FULL text blocks, not fragments

### CRITICAL SEPARATION RULE:
- **Text changes (text_added, text_deleted, text_modified) = ONLY words/text, NO numbers**
- **Numerical changes = ONLY numbers/amounts, NO text, NO words, NO labels**
- **If a line has both text and numerical changes, create SEPARATE entries for each**
- **Example**: "Tangible assets 4" → "Tangible assets 5" should be:
  - {"old": "4", "new": "5", "change_type": "numerical", ...} (just the number)
  - NOT: {"old": "Tangible assets 4", "new": "Tangible assets 5", "change_type": "text_modified"} (mixed)

### GENERAL:
- Return changes in ORDER from TOP to BOTTOM
- Only include actual changes (old ≠ new)
- Provide maximum context for unique identification
- For repeated values, use position_hint to distinguish

### CRITICAL: IGNORE THESE (DO NOT REPORT):
- **Watermarks**: Standalone letters T, F, A, R, D - remove from values, ignore as changes
- **Page numbers**: "Page 1", "## Page 1", "-1-", etc. - completely ignore
- **Headers/Footers**: Repeating text on every page - ignore if unchanged
- **DRAFT stamps**: "DRAFT", "CONFIDENTIAL" - ignore
- **Formatting**: Table separators, decorative lines, extra spaces - ignore
- **Boilerplate**: Standard legal text, disclaimers that are identical in both - ignore

**Focus ONLY on meaningful content changes (numbers, text, data), not formatting or document structure elements.**

Return JSON with all changes."""

    try:
        response = llm_client.client.chat.completions.create(
            model=llm_client.deployment,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,
            response_format={"type": "json_object"}
        )
        
        import json
        result = json.loads(response.choices[0].message.content)
        changes = result.get("changes", [])
        # import pdb; pdb.set_trace()

        
        # Post-processing: Filter out entries where old == new (no actual change)
        # This catches any LLM mistakes where it returns unchanged values
        filtered_changes = []
        for change in changes:
            old_val = str(change.get("old", "")).strip()
            new_val = str(change.get("new", "")).strip()
            change_type = change.get("change_type", "numerical")
            
            # Normalize for comparison
            # For text changes: remove ALL whitespace (spaces, newlines, tabs) and ignore case
            # For numerical: remove spaces and ignore case (but preserve number format)
            if change_type in ["text_added", "text_deleted", "text_modified"]:
                # Text normalization: remove all whitespace (spaces, newlines, tabs) and lowercase
                old_normalized = re.sub(r'\s+', '', old_val).lower()
                new_normalized = re.sub(r'\s+', '', new_val).lower()
            else:
                # Numerical normalization: remove spaces and lowercase (preserves number structure)
                old_normalized = old_val.replace(' ', '').lower()
                new_normalized = new_val.replace(' ', '').lower()
            
            # Only include if values are actually different
            if old_normalized != new_normalized:
                filtered_changes.append(change)
            else:
                # Log filtered entries for debugging
                print(f"[FILTERED] Removed unchanged value (type: {change_type}): old='{old_val[:50]}...', new='{new_val[:50]}...', context='{change.get('context', '')}'")
        
        return filtered_changes
        
    except Exception as e:
        raise Exception(f"LLM comparison failed: {str(e)}")


def locate_changes_in_markdown(new_page_markdown: str, changes: list, page_num: int) -> list:
    """
    Use LLM to intelligently locate where each change appears in the new page markdown.
    Returns list with location information for each change.
    """
    llm_client = LLMClient()
    
    if not changes:
        return []
    
    # Format changes for the prompt with all context information
    # Escape curly braces in values to avoid f-string format errors
    def escape_braces(text):
        if not text:
            return ""
        return str(text).replace('{', '{{').replace('}', '}}')
    
    changes_text = "\n".join([
        f"{i+1}. Type: {c.get('change_type', 'numerical')}\n"
        f"   old='{escape_braces(c.get('old', ''))[:100]}{'...' if len(c.get('old', '')) > 100 else ''}' → new='{escape_braces(c.get('new', ''))[:100]}{'...' if len(c.get('new', '')) > 100 else ''}'\n"
        f"   Context: {escape_braces(c.get('context', ''))}\n"
        f"   Section: {escape_braces(c.get('section', ''))}\n"
        f"   Row Label: {escape_braces(c.get('row_label', ''))}\n"
        f"   Before: '{escape_braces(c.get('surrounding_text_before', ''))}'\n"
        f"   After: '{escape_braces(c.get('surrounding_text_after', ''))}'\n"
        f"   Position: {escape_braces(c.get('position_hint', 'N/A'))}"
        for i, c in enumerate(changes)
    ])
    
    system_prompt = """You are an intelligent document analysis tool. Your task is to evaluate changes and locate meaningful ones within a document's markdown content.

## YOUR TASK
Given a list of changes (old → new values) and the NEW version's markdown content:
1. **EVALUATE** each change to determine if it's meaningful
2. **SKIP** meaningless changes (spacing, formatting, special characters, minor differences)
3. **LOCATE** only meaningful changes in the markdown

## MEANINGFUL vs MEANINGLESS CHANGES

### MEANINGFUL CHANGES (DO LOCATE):
- **Numerical changes**: Different numbers, amounts, percentages, dates
  - Example: "195" → "193", "(111,125)" → "259,183", "31.59%" → "17.37%"
- **Text content changes**: Different words, phrases, sentences, paragraphs
  - Example: "Net loss" → "Net profit", "Creditors" → "Debtors"
  - Example: Added/deleted entire sentences or paragraphs
- **Structural changes**: New sections, deleted rows, moved content
- **Semantic changes**: Changes that affect meaning or understanding

### MEANINGLESS CHANGES (SKIP - DO NOT LOCATE):
- **Spacing differences**: Only whitespace changed (spaces, tabs, newlines)
  - Example: "Net profit" → "Net  profit" (extra space)
  - Example: "Value: 100" → "Value:100" (missing space)
- **Character encoding**: Different representations of same character
  - Example: "–" (en dash) vs "-" (hyphen) if meaning is identical
  - Example: Smart quotes vs straight quotes if meaning is identical
- **Special character variations**: Different punctuation if meaning unchanged
  - Example: "Item A" vs "Item A." (trailing period)
  - Example: "(100)" vs "( 100 )" (spacing in parentheses)
- **Formatting artifacts**: Markdown formatting differences that don't change content
  - Example: "**Bold**" vs "Bold" (formatting only)
  - Example: Different line breaks in same paragraph
- **Case-only changes**: Only capitalization differs (unless it's a proper noun change)
  - Example: "profit" → "Profit" (unless it's a title/heading change)
- **Trivial differences**: Changes so minor they don't affect understanding
  - Example: "and" → "&" (unless it's a significant abbreviation)
  - Example: "1st" → "first" (unless it's a formatting preference)
- **PDF Extraction/OCR Errors**: Single character insertions, deletions, or substitutions that are clearly extraction artifacts
  - Example: "EngDland" → "England" (OCR inserted 'D' in "England")
  - Example: "undear" → "under" (OCR misread 'r' as 'ea')
  - Example: "teaar" → "tear" (OCR duplicated 'a')
  - Example: "comapny" → "company" (OCR swapped 'a' and 'p')
  - **Rule**: If the change is a single character difference in a word and the corrected version is a common English word, treat as extraction error
  - **Exception**: If the change affects meaning (e.g., "form" → "from"), it's meaningful

## EVALUATION CRITERIA
Before locating a change, ask yourself:
1. **Does this change affect the meaning or understanding of the document?**
2. **Is this a real content change or just formatting/spacing?**
3. **Would a human reader consider this a significant change?**
4. **Is the difference only in presentation, not in substance?**

**If the answer is NO to questions 1-3 or YES to question 4, SKIP the change (do not include it in locations).**

## OUTPUT FORMAT (strict JSON)
{
  "locations": [
    {
      "change_index": 0,
      "search_text": "exact text to search for in PDF",
      "context_before": "text that appears before the value",
      "context_after": "text that appears after the value"
    }
  ]
}

**IMPORTANT**: Only include locations for MEANINGFUL changes. If a change is meaningless, simply omit it from the locations array (but keep the change_index sequential for meaningful ones).

## RULES FOR LOCATING MEANINGFUL CHANGES

### FOR NUMERICAL CHANGES:
- Extract the EXACT numerical value as it appears
- Use context to find the unique instance
- search_text should be the exact number/amount
- **Only locate if the numbers are actually different (not just formatting)**

### FOR TEXT CHANGES (text_added, text_modified):
- For text_added: search_text should be the COMPLETE added text (or first sentence if very long)
- For text_modified: search_text should be the COMPLETE new text (or first sentence if very long)
- For long text blocks, use the first sentence or first 150 characters for search
- The full text will be used for highlighting, but search_text should be a unique snippet
- **Only locate if the text content is actually different (not just spacing/formatting)**

### GENERAL:
- Use ALL provided context to find the EXACT instance
- Use `surrounding_text_before` and `surrounding_text_after` to locate unique position
- Use `section` to narrow down search area
- Use `row_label` for table rows
- Use `position_hint` to distinguish multiple instances
- **Remove watermark artifacts (T, F, A, R, D standalone letters) from search_text**
- **Ignore page numbers, headers, footers when searching - focus on actual content**
- Process changes in order (change_index 0, 1, 2, etc.)
- **CRITICAL: Find the EXACT instance using context, not just any occurrence**
- **CRITICAL: Do NOT search in watermarks, footers, headers, or page numbers - only search in actual document content**

## EXAMPLES

### NUMERICAL CHANGE:
Change: old='(111,125)' → new='259,183'
Markdown contains: "...Net result for the year 259,183 (57,131)..."
Return:
{
  "change_index": 0,
  "search_text": "259,183",
  "context_before": "Net result for the year",
  "context_after": "(57,131)"
}

### NUMERICAL CHANGE (added):
Change: old='-' → new='370,308'
Markdown contains: "...Stocks 370,308 380,026..."
Return:
{
  "change_index": 0,
  "search_text": "370,308",
  "context_before": "Stocks",
  "context_after": "380,026"
}

### TEXT CHANGE (with corrupted/incomplete "new" value):
Change: old='Key performance indicators\nTBC' → new='Key perforbe Turnover and Earnings Before...' (NOTE: "new" value is corrupted/incomplete)
Markdown contains: "...Development and performance\n...\nKey performance indicators\nThe Directors consider the key performance indicators of the Company to be Turnover and Earnings Before\nInterest, Tax, Depreciation and Amortisation (EBITDA).\nTurnover for the year was £25.6m...\nOn behalf of the board..."
Context: surrounding_text_before='Development and performance', surrounding_text_after='On behalf of the board', section='Key performance indicators'

**IMPORTANT**: The "new" value is corrupted ("Key perforbe" instead of "Key performance indicators\nThe Directors consider..."). Use the context to find the correct text.

Return:
{
  "change_index": 0,
  "search_text": "The Directors consider the key performance indicators of the Company to be Turnover and Earnings Before\nInterest, Tax, Depreciation and Amortisation (EBITDA).\nTurnover for the year was £25.6m (2023: £22.4m) an increase of £3.2m (14.3%) as the business expanded its\ncapacity to meet increasing demand.\nEBITDA for the year was £15.0m (2023: £13.4m) an increase of £1.6m (11.9%).",
  "context_before": "Key performance indicators",
  "context_after": "On behalf of the board"
}

**Note**: Extract the COMPLETE text from the markdown that appears between the context markers, not the corrupted "new" value."""

    # Use string concatenation to avoid f-string format errors with curly braces in markdown
    user_prompt = """Evaluate and locate MEANINGFUL changes in the NEW page markdown content.

=== CHANGES TO EVALUATE (in order, top to bottom) ===
""" + changes_text + """

=== NEW PAGE MARKDOWN (PAGE """ + str(page_num) + """) ===
""" + new_page_markdown + """

## STEP 1: EVALUATE EACH CHANGE
For each change, determine if it's MEANINGFUL or MEANINGLESS:

**MEANINGFUL**: Different numbers, different words, added/deleted content, semantic changes
**MEANINGLESS**: Only spacing changed, only formatting changed, only special characters differ, trivial differences

**SKIP meaningless changes** - do not include them in the locations array.

## STEP 2: LOCATE MEANINGFUL CHANGES
For each MEANINGFUL change:
1. **FIRST**: Use the context information to locate the text in the markdown:
   - Use the surrounding_text_before and surrounding_text_after to find the unique position
   - Use the section name to narrow down the search area
   - Use the row_label to find the correct table row
   - Use the position_hint to distinguish between multiple instances

2. **THEN**: Find the actual text in the markdown that matches this context:
   - The "new" value provided might be incomplete, corrupted, or slightly different from what's in the markdown
   - **DO NOT rely solely on matching the exact "new" value** - use context to find the correct location
   - Look for text that appears between the surrounding_text_before and surrounding_text_after
   - Extract the COMPLETE text as it actually appears in the markdown (not the potentially corrupted "new" value)

3. **EXTRACT**: Get the exact text from the markdown:
   - Extract the EXACT text as it appears in the markdown (this will be used to search the PDF)
   - For text changes, extract the COMPLETE sentence/paragraph, not just a snippet
   - If the "new" value is corrupted or incomplete, use the context to find and extract the correct complete text
   - Remove any watermark artifacts (T, F, A, R, D standalone letters) from the search_text

4. **RETURN**: Provide the location information:
   - search_text: The COMPLETE text as it appears in the markdown (use this, not the potentially corrupted "new" value)
   - context_before: Text that appears before this text in the markdown
   - context_after: Text that appears after this text in the markdown

**CRITICAL: The "new" value from the first LLM call might be incomplete or corrupted. Use the context (surrounding_text_before, surrounding_text_after, section) to locate the ACTUAL text in the markdown, then extract the COMPLETE text as it appears. Do not just try to match the "new" value if it's clearly wrong or incomplete.**

**CRITICAL: Use the context information to find the EXACT instance specified, not just any occurrence of the value.**
For example, if position_hint says "first instance" or "main row", find that specific instance, not a different one with the same value.

**CRITICAL: IGNORE when searching:**
- Page numbers ("Page 1", "## Page 1", "-1-")
- Headers/Footers (repeating text on every page)
- Watermarks (standalone T, F, A, R, D letters)
- DRAFT stamps
- Table separators and formatting

**Only search in actual document content (tables, paragraphs, sections), NOT in headers, footers, or watermarks.**

**IMPORTANT**: Only return locations for MEANINGFUL changes. Skip meaningless ones entirely.

Process changes sequentially from top to bottom. Each change_index in the output should correspond to a meaningful change (you may skip some indices if those changes were meaningless).

Return JSON with locations for all MEANINGFUL changes only."""

    try:
        response = llm_client.client.chat.completions.create(
            model=llm_client.deployment,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,
            response_format={"type": "json_object"}
        )
        
        import json
        result = json.loads(response.choices[0].message.content)
        # import pdb; pdb.set_trace()
        return result.get("locations", [])
        
    except Exception as e:
        raise Exception(f"LLM location matching failed: {str(e)}")


def compare_markdown_with_llm(old_markdown: str, new_markdown: str) -> dict:
    """Use LLM to compare two markdown documents."""
    llm_client = LLMClient()
    
    # Truncate if too long
    max_chars = 500000
    if len(old_markdown) > max_chars:
        old_markdown = old_markdown[:max_chars] + "\n\n... [Document truncated] ..."
    if len(new_markdown) > max_chars:
        new_markdown = new_markdown[:max_chars] + "\n\n... [Document truncated] ..."
    
    system_prompt = """You are a document comparison tool. Compare two document versions WORD-BY-WORD and identify ALL differences.

## YOUR TASK - WORD-BY-WORD COMPARISON
1. Go through each section, table row, and paragraph systematically
2. Compare each WORD positionally between OLD and NEW versions
3. For each corresponding position, check if the word is identical
4. If ANY word differs, record the change
5. Track words that appear in one version but not the other

## COMPARISON METHODOLOGY

### For Text Paragraphs:
- Align sentences word-by-word
- Compare: word[1] OLD vs word[1] NEW, word[2] OLD vs word[2] NEW, etc.
- If a word changed, record the old word and new word
- If words were added, record them as added
- If words were removed, record them as removed

### For Tables:
- Compare each cell positionally (row by row, column by column)
- Within each cell, compare word-by-word
- If any word in a cell changed, record the change

### For Numbers/Values:
- Compare digit-by-digit or character-by-character
- "111,125" vs "259,183" → every digit changed
- "-" vs "370,308" → entire value changed

## OUTPUT FORMAT (strict JSON)
{
  "modified": [
    {"field": "descriptive label for context", "old": "exact old value/text", "new": "exact new value/text"}
  ],
  "added": [
    {"text": "exact new text that appears only in NEW version"}
  ],
  "removed": [
    {"text": "exact text that appears only in OLD version"}
  ]
}

## RULES

### MODIFIED
When the same field/position exists in both but ANY WORD changed:
- Extract the EXACT old value and EXACT new value
- Include field name for context (e.g., "Turnover > Sheep", "Balance Sheet > Stocks")
- For tables: use row label as field name
- Even a single word change should be recorded

Examples:
- Stock value: old "-" → new "370,308"
  → {"field": "Current Assets > Stocks", "old": "-", "new": "370,308"}
  
- Net result: old "(111,125)" → new "259,183"
  → {"field": "Net result for the year", "old": "(111,125)", "new": "259,183"}
  
- Text change: old "Net loss for the year" → new "Net profit/(loss) for the year"
  → {"field": "Net result caption", "old": "Net loss for the year", "new": "Net profit/(loss) for the year"}

### ADDED
Words or phrases that exist in NEW but have NO equivalent in OLD at that position:
- New words inserted into existing text
- New table rows
- New paragraphs or sections
- Return the exact text

### REMOVED
Words or phrases that exist in OLD but were completely deleted in NEW:
- Words removed from existing text
- Deleted table rows
- Removed paragraphs or sections
- Return the exact text

## WHAT TO IGNORE
- Watermark artifacts (standalone T, F, A, R, D letters scattered in text)
- Page numbers (## Page 1, -1-, etc.)
- Repeated headers/footers
- "DRAFT" watermarks
- Identical words/content
- Formatting differences (spacing, alignment) - but DO report if spacing changes affect meaning
- Table separators (|---|---|)

## CRITICAL
- Compare WORD-BY-WORD, not just line-by-line
- Return EXACT text values as they appear in the documents
- For highlighting to work, the "new" value must match text in the PDF exactly
- Include enough context in "field" to understand what changed
- Even single word changes should be caught and reported"""

    user_prompt = f"""Compare these two document versions WORD-BY-WORD and identify ALL differences.

=== OLD VERSION ===
{old_markdown}

=== NEW VERSION ===
{new_markdown}

Instructions:
1. Go through each section, table row, and paragraph systematically
2. Compare each WORD positionally between OLD and NEW
3. For each corresponding position, check if words match
4. If ANY word differs, record the exact old value and exact new value
5. Track words added or removed
6. Include descriptive field names for context

Be thorough - compare word-by-word, not just line-by-line. Catch every single change.

Return JSON with all changes."""

    try:
        response = llm_client.client.chat.completions.create(
            model=llm_client.deployment,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,
            response_format={"type": "json_object"}
        )
        
        import json
        result = json.loads(response.choices[0].message.content)
        return {
            "removed": result.get("removed", []),
            "added": result.get("added", []),
            "modified": result.get("modified", [])
        }
        
    except Exception as e:
        raise Exception(f"LLM comparison failed: {str(e)}")


def compare_pdfs_with_marker(old_pdf_bytes: bytes, new_pdf_bytes: bytes, library: str = None) -> dict:
    """
    Main comparison pipeline - PAGE BY PAGE:
    1. Convert both PDFs to Markdown
    2. Extract pages from markdown
    3. Compare each page with LLM
    4. Return page-by-page changes
    """
    lib = library or ACTIVE_LIBRARY
    
    print("=" * 60)
    print(f"PDF COMPARISON ({lib} + LLM) - PAGE BY PAGE")
    print("=" * 60)
    
    # Step 1: Convert to Markdown
    print("\n[Step 1] Converting OLD PDF to Markdown...")
    old_markdown = convert_pdf_bytes_to_markdown(old_pdf_bytes, lib)
    
    print("\n[Step 1] Converting NEW PDF to Markdown...")
    new_markdown = convert_pdf_bytes_to_markdown(new_pdf_bytes, lib)
    
    # Step 2: Normalize
    print("\n[Step 2] Normalizing Markdown...")
    old_markdown = normalize_markdown(old_markdown)
    new_markdown = normalize_markdown(new_markdown)
    
    # Step 3: Extract pages
    print("\n[Step 3] Extracting pages...")
    old_pages = extract_pages_from_markdown(old_markdown)
    new_pages = extract_pages_from_markdown(new_markdown)
    
    print(f"  OLD: {len(old_pages)} pages")
    print(f"  NEW: {len(new_pages)} pages")
    
    # Step 4: Compare each page with LLM
    print("\n[Step 4] Comparing pages with LLM...")
    page_changes = {}  # {page_num: [{"old": "...", "new": "...", "context": "..."}, ...]}
    all_changes = []
    
    # Get all page numbers (union of old and new)
    all_page_nums = set(old_pages.keys()) | set(new_pages.keys())
    
    for page_num in sorted(all_page_nums):
        old_page_content = old_pages.get(page_num, "")
        new_page_content = new_pages.get(page_num, "")
        
        if not old_page_content and not new_page_content:
            continue
        
        print(f"  Comparing page {page_num}...")
        
        try:
            # Step 4a: Compare and get changes
            changes = compare_markdown_page_with_llm(
                old_page_content, 
                new_page_content, 
                page_num
            )
            
            if changes:
                # Step 4b: Use LLM to locate each change in the markdown
                # The second LLM will filter out meaningless changes (spacing, formatting, etc.)
                print(f"    Evaluating and locating {len(changes)} changes in markdown (filtering meaningless ones)...")
                locations = locate_changes_in_markdown(
                    new_page_content,
                    changes,
                    page_num
                )
                
                # Merge location info with changes
                # Only include changes that have locations (meaningful changes)
                # The second LLM filters out meaningless changes (spacing, formatting, etc.)
                # The second LLM also corrects corrupted/incomplete "new" values by extracting the actual text from markdown
                changes_with_locations = []
                for i, change in enumerate(changes):
                    # Find location for this change
                    location = next((loc for loc in locations if loc.get("change_index") == i), None)
                    if location:
                        # This is a meaningful change - include it
                        search_text = location.get("search_text", "")
                        
                        # If the 2nd LLM provided a search_text, use it (it's the corrected text from markdown)
                        # For text changes, prefer the search_text from location as it's the actual text from markdown
                        if search_text and change.get("change_type") in ["text_added", "text_modified"]:
                            # Use the corrected search_text from the 2nd LLM (actual text from markdown)
                            change["search_text"] = search_text
                            # Optionally update "new" with the corrected text if it's more complete
                            if len(search_text) > len(str(change.get("new", ""))):
                                change["new"] = search_text
                        else:
                            # For numerical or if no search_text provided, use original
                            change["search_text"] = search_text or change.get("new", "")
                        
                        change["context_before"] = location.get("context_before", change.get("context_before", ""))
                        change["context_after"] = location.get("context_after", change.get("context_after", ""))
                        changes_with_locations.append(change)
                    # If no location found, the second LLM determined this change is meaningless
                    # (spacing, formatting, special characters, etc.) - skip it
                
                page_changes[page_num] = changes_with_locations
                all_changes.extend(changes_with_locations)
                filtered_count = len(changes) - len(changes_with_locations)
                if filtered_count > 0:
                    print(f"    Found {len(changes)} changes, {len(changes_with_locations)} meaningful (filtered {filtered_count} meaningless)")
                else:
                    print(f"    Found {len(changes)} changes, all meaningful")
        except Exception as e:
            print(f"    Error comparing page {page_num}: {e}")
    
    # Convert to old format for compatibility
    modified = []
    for change in all_changes:
        modified.append({
            "field": change.get("context", "Unknown"),
            "old": change.get("old", ""),
            "new": change.get("new", "")
        })
    
    print(f"\n  Results: {len(all_changes)} total changes across {len(page_changes)} pages")
    
    return {
        "old_markdown": old_markdown,
        "new_markdown": new_markdown,
        "changes": {
            "modified": modified,
            "added": [],
            "removed": []
        },
        "page_changes": page_changes,  # New: page-by-page changes
        "method": lib
    }


def is_marker_available() -> bool:
    """Check if any library is available."""
    return len(AVAILABLE_LIBRARIES) > 0


def get_extraction_method() -> str:
    """Get the active extraction method name."""
    return ACTIVE_LIBRARY or "None"


def get_available_libraries() -> list:
    """Get list of available libraries."""
    return AVAILABLE_LIBRARIES.copy()


def set_active_library(library: str):
    """Set the active library."""
    global ACTIVE_LIBRARY
    if library in AVAILABLE_LIBRARIES:
        ACTIVE_LIBRARY = library
        print(f"[PDF2MD] Switched to: {library}")
    else:
        raise ValueError(f"Library not available: {library}. Available: {AVAILABLE_LIBRARIES}")
