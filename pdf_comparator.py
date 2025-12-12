"""
PDF comparison module for the Version Comparison Agent.
Robust text/table extraction and LLM-powered comparison with highlighting.
"""
import fitz  # PyMuPDF
import io
import re
from dataclasses import dataclass
from llm_client import LLMClient


@dataclass
class ComparisonColors:
    """Colors for highlighting changes (RGB tuples normalized to 0-1)."""
    ADDED = (0.4, 0.9, 0.4)      # Bright green  
    MODIFIED = (1, 0.9, 0.4)     # Yellow/amber


class PDFComparator:
    """Handles PDF version comparison and highlighted output generation."""
    
    def __init__(self):
        self.llm_client = LLMClient()
        self.colors = ComparisonColors()
        
        # Patterns to ignore (watermarks, etc.)
        self.ignore_patterns = [
            r'^draft$', r'^confidential$', r'^page\s*\d*$',
            r'^\d+$',  # Standalone page numbers
        ]
    
    def _normalize_text(self, text: str) -> str:
        """Clean and normalize extracted text."""
        if not text:
            return ""
        
        # Fix hyphenation at line breaks
        text = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', text)
        
        # Normalize whitespace
        text = re.sub(r'[ \t]+', ' ', text)  # Multiple spaces to single
        text = re.sub(r'\n\s*\n', '\n\n', text)  # Multiple newlines to double
        
        # Clean each line
        lines = [line.strip() for line in text.split('\n')]
        text = '\n'.join(line for line in lines if line)
        
        return text.strip()
    
    def _should_ignore(self, text: str) -> bool:
        """Check if text should be ignored (watermarks, page numbers, etc.)."""
        if not text:
            return True
        
        text_lower = text.lower().strip()
        
        for pattern in self.ignore_patterns:
            if re.match(pattern, text_lower, re.IGNORECASE):
                return True
        
        return False
    
    def _is_inside_any_bbox(self, inner: tuple, boxes: list, margin: float = 2) -> bool:
        """Check if inner bbox is inside any of the given bboxes."""
        ix0, iy0, ix1, iy1 = inner
        for ox0, oy0, ox1, oy1 in boxes:
            if (ix0 >= ox0 - margin and iy0 >= oy0 - margin and 
                ix1 <= ox1 + margin and iy1 <= oy1 + margin):
                return True
        return False
    
    def _clean_cell_value(self, cell) -> str:
        """Clean and normalize a table cell value."""
        if cell is None:
            return ""
        
        text = str(cell).strip()
        
        # Normalize whitespace
        text = re.sub(r'\s+', ' ', text)
        
        # Remove common artifacts
        text = text.replace('\n', ' ').replace('\r', '')
        
        return text.strip()
    
    def _parse_numeric(self, text: str) -> tuple[bool, str]:
        """
        Try to parse text as a number and format it consistently.
        Returns: (is_numeric, formatted_value)
        """
        if not text:
            return False, text
        
        # Remove currency symbols and common formatting
        clean = re.sub(r'[£$€¥,\s]', '', text)
        clean = clean.replace('(', '-').replace(')', '')  # Handle accounting negatives
        
        try:
            # Try to parse as number
            num = float(clean)
            # Keep original formatting but mark as numeric
            return True, text
        except ValueError:
            return False, text
    
    def _detect_header_row(self, rows: list) -> int:
        """
        Detect which row is the header row.
        Returns the index of the header row (usually 0), or -1 if no clear header.
        """
        if not rows or len(rows) < 2:
            return -1
        
        first_row = rows[0]
        second_row = rows[1] if len(rows) > 1 else []
        
        # Heuristics for header detection:
        # 1. First row has no numbers, second row has numbers
        # 2. First row cells are shorter (labels vs data)
        
        first_has_numbers = any(
            self._parse_numeric(cell)[0] for cell in first_row if cell
        )
        second_has_numbers = any(
            self._parse_numeric(cell)[0] for cell in second_row if cell
        )
        
        if not first_has_numbers and second_has_numbers:
            return 0
        
        # Check if first row looks like headers (shorter text, no special chars)
        first_avg_len = sum(len(str(c)) for c in first_row if c) / max(len(first_row), 1)
        second_avg_len = sum(len(str(c)) for c in second_row if c) / max(len(second_row), 1)
        
        if first_avg_len < second_avg_len * 0.7:  # First row significantly shorter
            return 0
        
        return -1  # No clear header
    
    def _extract_tables(self, page) -> tuple[list, list]:
        """
        Extract tables from page with improved handling.
        
        Features:
        - Detects header rows
        - Cleans cell values
        - Handles merged cells
        - Preserves numeric formatting
        - Returns structured table data with metadata
        
        Returns: (table_data_list, table_bbox_list)
        Each table_data is: {"headers": [...], "rows": [[...], ...], "has_header": bool}
        """
        tables_data = []
        table_bboxes = []
        
        try:
            # Use PyMuPDF's table finder with different strategies
            tables = page.find_tables(
                snap_tolerance=3,      # Tolerance for snapping to lines
                snap_x_tolerance=3,
                snap_y_tolerance=3,
                join_tolerance=3,      # Tolerance for joining broken lines
                edge_min_length=3,     # Minimum line length
                min_words_vertical=1,  # Minimum words to detect vertical lines
                min_words_horizontal=1
            )
            
            for table in tables:
                raw_data = table.extract()
                
                if not raw_data:
                    continue
                
                # Clean all cells
                cleaned_rows = []
                for row in raw_data:
                    cleaned_row = [self._clean_cell_value(cell) for cell in row]
                    # Keep row if it has any content
                    if any(cleaned_row):
                        cleaned_rows.append(cleaned_row)
                
                if not cleaned_rows:
                    continue
                
                # Detect header row
                header_idx = self._detect_header_row(cleaned_rows)
                
                # Structure the table data
                if header_idx >= 0 and header_idx < len(cleaned_rows):
                    table_info = {
                        "headers": cleaned_rows[header_idx],
                        "rows": cleaned_rows[header_idx + 1:],
                        "has_header": True,
                        "all_rows": cleaned_rows  # Keep all rows for formatting
                    }
                else:
                    table_info = {
                        "headers": [],
                        "rows": cleaned_rows,
                        "has_header": False,
                        "all_rows": cleaned_rows
                    }
                
                # Only add tables with actual data rows
                if table_info["rows"] or table_info["all_rows"]:
                    tables_data.append(table_info)
                    table_bboxes.append(table.bbox)
                    
        except Exception as e:
            # Fallback: try simpler extraction
            try:
                tables = page.find_tables()
                for table in tables:
                    raw_data = table.extract()
                    if raw_data and any(any(cell for cell in row) for row in raw_data):
                        cleaned_rows = []
                        for row in raw_data:
                            cleaned_row = [self._clean_cell_value(cell) for cell in row]
                            if any(cleaned_row):
                                cleaned_rows.append(cleaned_row)
                        
                        if cleaned_rows:
                            tables_data.append({
                                "headers": [],
                                "rows": cleaned_rows,
                                "has_header": False,
                                "all_rows": cleaned_rows
                            })
                            table_bboxes.append(table.bbox)
            except Exception:
                pass
        
        return tables_data, table_bboxes
    
    def _extract_text_blocks(self, page, exclude_bboxes: list) -> list:
        """
        Extract text blocks from page, excluding areas covered by tables.
        Returns list of {"text": str, "bbox": tuple, "y_pos": float}
        """
        text_blocks = []
        
        # Get detailed text extraction
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        
        for block in blocks:
            if "lines" not in block:
                continue
            
            block_bbox = block["bbox"]
            
            # Skip if inside a table
            if self._is_inside_any_bbox(block_bbox, exclude_bboxes):
                continue
            
            # Extract text from all lines in block
            block_text = ""
            for line in block["lines"]:
                line_text = ""
                for span in line["spans"]:
                    span_text = span.get("text", "")
                    line_text += span_text
                
                line_text = line_text.strip()
                if line_text:
                    block_text += line_text + "\n"
            
            block_text = self._normalize_text(block_text)
            
            # Skip empty or ignorable text
            if not block_text or self._should_ignore(block_text):
                continue
            
            text_blocks.append({
                "text": block_text,
                "bbox": block_bbox,
                "y_pos": block_bbox[1]  # Top position for sorting
            })
        
        # Sort by reading order (top to bottom, left to right)
        text_blocks.sort(key=lambda b: (round(b["y_pos"] / 20) * 20, b["bbox"][0]))
        
        return text_blocks
    
    def extract_content_structured(self, pdf_bytes: bytes) -> dict:
        """
        Extract all content from PDF with robust structure preservation.
        
        - Extracts tables separately (avoids duplicate text)
        - Preserves reading order
        - Normalizes all text
        - Filters out watermarks and noise
        """
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        result = {
            "pages": [],
            "full_text": "",
            "all_tables": []
        }
        
        for page_num, page in enumerate(doc):
            # Extract tables first (and get their bboxes to exclude from text)
            tables_data, table_bboxes = self._extract_tables(page)
            
            # Extract text blocks (excluding table areas)
            text_blocks = self._extract_text_blocks(page, table_bboxes)
            
            page_content = {
                "page_num": page_num,
                "text_blocks": text_blocks,
                "tables": tables_data
            }
            
            result["pages"].append(page_content)
            result["all_tables"].extend(tables_data)
            
            # Build full text for this page
            page_text = "\n".join(block["text"] for block in text_blocks)
            result["full_text"] += f"\n{page_text}\n"
        
        doc.close()
        return result
    
    def format_content_for_comparison(self, content: dict) -> str:
        """
        Format extracted content into a clear, structured string for LLM.
        
        Features:
        - Clear PAGE markers
        - TABLE markers with headers and labeled rows
        - Row labels from first column when available
        - Clear cell separators
        """
        formatted_parts = []
        
        for page in content["pages"]:
            page_num = page["page_num"] + 1
            formatted_parts.append(f"\n══════ PAGE {page_num} ══════\n")
            
            # Add text blocks
            for block in page["text_blocks"]:
                text = block["text"]
                if text:
                    formatted_parts.append(text)
                    formatted_parts.append("")  # Empty line between blocks
            
            # Add tables with improved structure
            for table_idx, table_info in enumerate(page["tables"], 1):
                formatted_parts.append(f"\n┌─ TABLE {table_idx} ─┐")
                
                # Handle both old format (list) and new format (dict)
                if isinstance(table_info, dict):
                    headers = table_info.get("headers", [])
                    rows = table_info.get("all_rows", table_info.get("rows", []))
                    has_header = table_info.get("has_header", False)
                    
                    # Format header row if present
                    if has_header and headers:
                        header_str = " │ ".join(h if h else "—" for h in headers)
                        formatted_parts.append(f"  [HEADER]: {header_str}")
                        formatted_parts.append("  " + "─" * 50)
                        # Skip header in data rows
                        data_rows = rows[1:] if rows and rows[0] == headers else rows
                    else:
                        data_rows = rows
                    
                    # Format data rows with row labels
                    for row_idx, row in enumerate(data_rows, 1):
                        if not row:
                            continue
                        
                        # Use first column as row label if it looks like a label
                        first_col = row[0] if row else ""
                        is_label = first_col and not self._parse_numeric(first_col)[0]
                        
                        if is_label and len(row) > 1:
                            # First column is label, rest is data
                            label = first_col[:30]  # Truncate long labels
                            data_cells = row[1:]
                            data_str = " │ ".join(cell if cell else "—" for cell in data_cells)
                            formatted_parts.append(f"  [{label}]: {data_str}")
                        else:
                            # No clear label, use row number
                            row_str = " │ ".join(cell if cell else "—" for cell in row)
                            formatted_parts.append(f"  Row {row_idx}: {row_str}")
                else:
                    # Fallback for old list format
                    for row_idx, row in enumerate(table_info, 1):
                        row_str = " │ ".join(cell if cell else "—" for cell in row)
                        formatted_parts.append(f"  Row {row_idx}: {row_str}")
                
                formatted_parts.append(f"└─ /TABLE {table_idx} ─┘\n")
        
        return "\n".join(formatted_parts)
    
    def highlight_text_in_pdf(self, doc: fitz.Document, text: str, color: tuple) -> bool:
        """Highlight specific text in the PDF. Returns True if successful."""
        if not text or len(text.strip()) < 2:
            return False
        
        text = text.strip()
        
        # Skip watermarks and common noise
        if self._should_ignore(text):
            return False
        
        for page in doc:
            instances = page.search_for(text, quads=True)
            
            if instances:
                try:
                    annot = page.add_highlight_annot(instances[0])
                    annot.set_colors(stroke=color)
                    annot.update()
                    return True
                except Exception:
                    pass
        
        return False
    
    def highlight_value_with_context(self, doc: fitz.Document, value: str, 
                                      context: str, color: tuple) -> bool:
        """
        Highlight a value using context to find the correct instance.
        Context is typically the field name or nearby text.
        """
        if not value or len(str(value).strip()) < 1:
            return False
        
        value = str(value).strip()
        
        if self._should_ignore(value):
            return False
        
        for page in doc:
            value_instances = page.search_for(value, quads=True)
            
            if not value_instances:
                continue
            
            # If we have context, find the value closest to it
            if context:
                context = str(context).strip()
                context_instances = page.search_for(context)
                
                if context_instances:
                    context_rect = context_instances[0]
                    
                    # Find value instance closest to context
                    best_instance = None
                    best_score = float('inf')
                    
                    for inst in value_instances:
                        inst_rect = inst.rect if hasattr(inst, 'rect') else fitz.Rect(inst)
                        
                        # Score: prefer same line (y), then close horizontally
                        y_diff = abs(inst_rect.y0 - context_rect.y0)
                        x_diff = abs(inst_rect.x0 - context_rect.x1)
                        
                        # Heavy penalty for different lines
                        score = y_diff * 10 + x_diff
                        
                        if score < best_score:
                            best_score = score
                            best_instance = inst
                    
                    if best_instance and best_score < 500:
                        try:
                            annot = page.add_highlight_annot(best_instance)
                            annot.set_colors(stroke=color)
                            annot.update()
                            return True
                        except Exception:
                            pass
            
            # Fallback: highlight first instance
            try:
                annot = page.add_highlight_annot(value_instances[0])
                annot.set_colors(stroke=color)
                annot.update()
                return True
            except Exception:
                pass
        
        return False
    
    def compare_pdfs(self, old_pdf_bytes: bytes, new_pdf_bytes: bytes) -> dict:
        """
        Compare two PDF files and generate highlighted output.
        
        Output highlights:
        - GREEN: Genuinely new content
        - YELLOW: Modified values (only the changed value, not whole line)
        
        Removed content is listed in changes but not highlighted 
        (since it doesn't exist in the new PDF).
        """
        # Extract content from both PDFs
        old_content = self.extract_content_structured(old_pdf_bytes)
        new_content = self.extract_content_structured(new_pdf_bytes)
        
        # Format for LLM comparison
        old_formatted = self.format_content_for_comparison(old_content)
        new_formatted = self.format_content_for_comparison(new_content)
        
        # Get changes from LLM
        changes = self.llm_client.compare_pdf_content(old_formatted, new_formatted)
        
        # Open new PDF for highlighting
        doc = fitz.open(stream=new_pdf_bytes, filetype="pdf")
        
        # Highlight ADDED content (green)
        for item in changes.get("added", []):
            if isinstance(item, dict):
                text = item.get("text", "")
                if text:
                    self.highlight_text_in_pdf(doc, text, self.colors.ADDED)
            elif item:
                self.highlight_text_in_pdf(doc, str(item), self.colors.ADDED)
        
        # Highlight MODIFIED values (yellow) - only the new value
        for item in changes.get("modified", []):
            if isinstance(item, dict):
                new_val = item.get("new", "")
                field = item.get("field", item.get("context", ""))
                
                if new_val:
                    self.highlight_value_with_context(
                        doc, str(new_val), str(field), self.colors.MODIFIED
                    )
        
        # Save result
        output = io.BytesIO()
        doc.save(output)
        doc.close()
        
        return {
            "highlighted_pdf": output.getvalue(),
            "changes": changes
        }
    
    def get_change_summary(self, changes: dict) -> str:
        """Generate human-readable summary."""
        removed = len(changes.get("removed", []))
        added = len(changes.get("added", []))
        modified = len(changes.get("modified", []))
        
        parts = []
        if removed > 0:
            parts.append(f"🔴 {removed} removed")
        if added > 0:
            parts.append(f"🟢 {added} added")
        if modified > 0:
            parts.append(f"🟡 {modified} modified")
        
        return " | ".join(parts) if parts else "No changes detected"
