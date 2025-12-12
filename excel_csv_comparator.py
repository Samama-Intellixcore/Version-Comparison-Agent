"""
Excel and CSV comparison module for the Version Comparison Agent.
Handles simple row-by-row comparison using LLM and generates Change table.
"""
import pandas as pd
import io
import json
import re


class ExcelCSVComparator:
    """Handles Excel and CSV version comparison with simple row-by-row approach using LLM."""
    
    def __init__(self):
        from llm_client import LLMClient
        self.llm_client = LLMClient()
    
    def read_csv(self, file_bytes: bytes) -> pd.DataFrame:
        """Read CSV file into DataFrame."""
        for encoding in ['utf-8', 'latin-1', 'cp1252']:
            try:
                return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError("Could not decode CSV file with supported encodings")
    
    def read_excel(self, file_bytes: bytes) -> dict[str, pd.DataFrame]:
        """Read Excel file into dictionary of DataFrames (one per sheet)."""
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets = {}
        for sheet_name in excel_file.sheet_names:
            sheets[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name)
        return sheets
    
    def _compare_batch_with_llm(self, old_rows: list[dict], new_rows: list[dict], batch_start: int) -> list[str]:
        """
        Use LLM to compare a batch of rows and calculate changes.
        
        Args:
            old_rows: List of dictionaries with row data from old version
            new_rows: List of dictionaries with row data from new version
            batch_start: Starting index for this batch
            
        Returns:
            List of change descriptions as strings
        """
        system_prompt = """You are a precise financial data comparison assistant. Compare two datasets row-by-row and identify ALL changes.

**CRITICAL INSTRUCTIONS - EXPLICIT INDEX REQUIRED:**
1. For EACH row pair, you MUST return the exact index number along with the change
2. Check EVERY column for changes - don't miss any
3. For each changed column, include: column name, old value, new value, and difference
4. Return ALL row pairs, even if there's no change (use empty string "")

**INDEX REQUIREMENT:** Each change entry MUST include the exact index number from the input row pair.

Given pairs of old and new rows (which may have multiple columns), determine the change type for each row:

1. **No Change**: If all values in ALL columns are identical → return ""
2. **Value Changed**: If ANY column values changed, you MUST provide:
   - Column name (use the EXACT column name from the data)
   - Previous value (old value)
   - New value
   - The change (difference: new - old)
   
   Format for single column change: "ColumnName: old_value → new_value (change)"
   Format for multiple column changes: "Column1: old1 → new1 (diff1); Column2: old2 → new2 (diff2)"
   
   - Check EVERY column, not just the last one
   - Use the EXACT column name as it appears in the data
   - Parse numbers correctly - handle commas, parentheses (negative values), and decimal points
   - Format all values with commas and parentheses for negatives (accounting format)
   - For text changes (non-numeric), show: "ColumnName: 'old_text' → 'new_text'"
3. **Row Added**: If old row is empty/blank but new row has data → return "ROW ADDED"
4. **Row Removed**: If old row has data but new row is empty/blank → return "ROW REMOVED"

**IMPORTANT FORMATTING RULES:**
- ALWAYS include the EXACT column name from the data in the change description
- For numeric changes: "ColumnName: old_value → new_value (difference)"
- For text changes: "ColumnName: 'old_text' → 'new_text'"
- If multiple columns change, separate with semicolons: "Col1: ...; Col2: ..."
- Use the exact format as shown in the data (commas, parentheses)
- Calculate difference correctly: new - old
- Format difference with commas and parentheses for negatives
- MAINTAIN ORDER: changes[0] must correspond to row_index 0, changes[1] to row_index 1, etc.

Examples:
- old={"Name": "John", "Theory Number": "30"}, new={"Name": "John", "Theory Number": "31"} 
  → "Theory Number: 30 → 31 (1)"
  
- old={"Name": "John", "Value": "45"}, new={"Name": "John", "Value": "489"} 
  → "Value: 45 → 489 (444)"
  
- old={"Account": "Cash", "Value": "(2,043,420)"}, new={"Account": "Cash", "Value": "(5,766,604)"} 
  → "Value: (2,043,420) → (5,766,604) ((3,723,184))"
  
- old={"Name": "John", "Value": "150,000"}, new={"Name": "Jane", "Value": "150,647"} 
  → "Name: 'John' → 'Jane'; Value: 150,000 → 150,647 (647)"
  
- old={"Account": "Asset", "Value": "32,302,655"}, new={"Account": "Asset", "Value": "32,302,655"} 
  → "" (no change)
  
- old={"Account": "", "Value": ""}, new={"Account": "New Item", "Value": "100,000"} 
  → "ROW ADDED"
  
- old={"Account": "Old Item", "Value": "50,000"}, new={"Account": "", "Value": ""} 
  → "ROW REMOVED"

Return a JSON object with an array of change objects, each containing the index and change description:
{
  "changes": [
    {"index": 0, "change": ""},
    {"index": 1, "change": "Theory Number: 30 → 31 (1)"},
    {"index": 2, "change": "Value: 45 → 489 (444)"},
    {"index": 3, "change": "ROW ADDED"},
    {"index": 4, "change": "ROW REMOVED"},
    ...
  ]
}

**CRITICAL:** 
- Each object MUST have "index" (the exact index from input) and "change" (the change description or "")
- Include ALL row pairs, even if no change (use "" for change)
- The index must match exactly the index from the input row pair
- Return ONLY valid JSON."""

        # Prepare data for LLM - simple index-based matching
        rows_data = []
        for i in range(len(old_rows)):
            rows_data.append({
                "index": batch_start + i,
                "old": old_rows[i],
                "new": new_rows[i]
            })
        
        user_prompt = f"""Compare these {len(rows_data)} row pairs. For EACH row pair:

1. Check EVERY column for changes - don't miss any
2. If ANY column changed, provide: "ColumnName: old_value → new_value (change)"
3. If MULTIPLE columns changed, list all: "Col1: old1 → new1 (diff1); Col2: old2 → new2 (diff2)"
4. If no change in any column, use "" for change
5. If row added (old is empty, new has data), use "ROW ADDED" for change
6. If row removed (old has data, new is empty), use "ROW REMOVED" for change

**CRITICAL: Return an array of objects with "index" and "change". Include ALL {len(rows_data)} row pairs with their exact index numbers.**

Row pairs (index, old, new):
{json.dumps(rows_data, indent=2)}

Return JSON: {{"changes": [{{"index": 0, "change": ""}}, {{"index": 1, "change": "ColumnName: old → new (diff)"}}, ...]}} with exactly {len(rows_data)} objects, each with the correct index."""

        response = self.llm_client.client.chat.completions.create(
            model=self.llm_client.deployment,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0,
            response_format={"type": "json_object"},
            max_tokens=32000
        )
        
        content = response.choices[0].message.content.strip()
        
        # Try to extract JSON if it's wrapped in markdown code blocks
        if content.startswith("```"):
            lines = content.split("\n")
            json_start = None
            json_end = None
            for i, line in enumerate(lines):
                if line.strip().startswith("```"):
                    if json_start is None:
                        json_start = i + 1
                    else:
                        json_end = i
                        break
            if json_start and json_end:
                content = "\n".join(lines[json_start:json_end])
        
        result = json.loads(content)
        changes_list = result.get("changes", [])
        # import pdb; pdb.set_trace()

        def _strip_leading_index_from_change(change_str: str) -> str:
            """
            Some LLM responses include a row identifier/index prefix (e.g. "32302655:" or a first line "32302655").
            Strip that so the worksheet Change cell only shows the actual change description.
            """
            s = str(change_str).strip()
            if not s:
                return ""

            # If it starts as "12345: rest...", drop the "12345:" prefix
            m = re.match(r"^\s*[\d,]+\s*:\s*(.+)$", s, flags=re.DOTALL)
            if m:
                s = m.group(1).strip()

            # If first non-empty line is just an index/id (e.g. "32302655" or "32,302,655:"), drop that line.
            lines = [ln.strip() for ln in s.splitlines() if str(ln).strip()]
            if lines and re.fullmatch(r"[\d,]+:?", lines[0]):
                lines = lines[1:]

            return "\n".join(lines).strip()
        
        # Create a dictionary mapping index to change
        index_to_change = {}
        for change_obj in changes_list:
            if isinstance(change_obj, dict):
                idx = change_obj.get("index")
                change_val = change_obj.get("change", "")
                if idx is not None:
                    # Clean up the change value
                    change_str = _strip_leading_index_from_change(change_val)
                    if change_str.lower() in ['', 'no change', 'none', 'null', '""', "''", '0', '0.0']:
                        index_to_change[idx] = ""
                    else:
                        index_to_change[idx] = change_str
            elif isinstance(change_obj, str):
                # Fallback: if LLM returns array of strings, use position
                idx = len(index_to_change)
                change_str = _strip_leading_index_from_change(change_obj)
                if change_str.lower() in ['', 'no change', 'none', 'null', '""', "''", '0', '0.0']:
                    index_to_change[idx] = ""
                else:
                    index_to_change[idx] = change_str
        
        # Build result array in correct order using explicit indices
        cleaned_changes = []
        for i in range(len(rows_data)):
            expected_index = batch_start + i
            # Use explicit index from LLM, or fallback to position
            if expected_index in index_to_change:
                cleaned_changes.append(index_to_change[expected_index])
            elif i in index_to_change:
                cleaned_changes.append(index_to_change[i])
            else:
                cleaned_changes.append("")
        
        return cleaned_changes
    
    def _compare_all_rows_with_llm(self, old_rows: list[dict], new_rows: list[dict]) -> list[str]:
        """
        Use LLM to compare all rows in batches and calculate changes.
        
        Args:
            old_rows: List of dictionaries with row data from old version
            new_rows: List of dictionaries with row data from new version
            
        Returns:
            List of change descriptions as strings:
            - Empty string "" if no change
            - Difference value if value changed (e.g., "647", "(3,723,184)")
            - "ROW ADDED" if row is new (not in old version)
            - "ROW REMOVED" if row was removed (not in new version)
        """
        total_rows = max(len(old_rows), len(new_rows))
        batch_size = 50  # Process 50 rows at a time to avoid token limits
        all_changes = []
        
        for batch_start in range(0, total_rows, batch_size):
            batch_end = min(batch_start + batch_size, total_rows)
            
            # Extract batch
            batch_old = []
            batch_new = []
            for i in range(batch_start, batch_end):
                batch_old.append(old_rows[i] if i < len(old_rows) else {})
                batch_new.append(new_rows[i] if i < len(new_rows) else {})
            
            # Compare batch
            batch_changes = self._compare_batch_with_llm(batch_old, batch_new, batch_start)
            all_changes.extend(batch_changes)
        
        return all_changes
    
    def _align_dataframes(self, old_df: pd.DataFrame, new_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
        """
        Align two dataframes to the same number of rows by padding with empty rows.
        
        Returns:
            Tuple of (aligned_old_df, aligned_new_df) with same number of rows
        """
        max_rows = max(len(old_df), len(new_df))
        
        # Get all unique columns from both dataframes
        all_columns = list(dict.fromkeys(list(old_df.columns) + list(new_df.columns)))
        
        # Create aligned old_df
        aligned_old_rows = []
        for i in range(max_rows):
            if i < len(old_df):
                row_dict = {}
                for col in all_columns:
                    if col in old_df.columns:
                        val = old_df.iloc[i][col]
                        row_dict[col] = val if not pd.isna(val) else ""
                    else:
                        row_dict[col] = ""
                aligned_old_rows.append(row_dict)
            else:
                aligned_old_rows.append({col: "" for col in all_columns})
        
        aligned_old_df = pd.DataFrame(aligned_old_rows)
        
        # Create aligned new_df
        aligned_new_rows = []
        for i in range(max_rows):
            if i < len(new_df):
                row_dict = {}
                for col in all_columns:
                    if col in new_df.columns:
                        val = new_df.iloc[i][col]
                        row_dict[col] = val if not pd.isna(val) else ""
                    else:
                        row_dict[col] = ""
                aligned_new_rows.append(row_dict)
            else:
                aligned_new_rows.append({col: "" for col in all_columns})
        
        aligned_new_df = pd.DataFrame(aligned_new_rows)
        
        return aligned_old_df, aligned_new_df
    
    def compare_dataframes_simple(self, old_df: pd.DataFrame, new_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """
        Simple row-by-row comparison using LLM.
        Returns aligned old_df, new_df, and change_df with Change column.
        
        Returns:
            Tuple of (aligned_old_df, aligned_new_df, change_df) where change_df has Change column
        """
        # Determine the maximum number of rows
        max_rows = max(len(old_df), len(new_df))
        
        # Build row data for LLM - preserve all columns from new_df
        old_rows = []
        new_rows = []
        
        # Get all unique columns from both dataframes
        all_columns = list(dict.fromkeys(list(old_df.columns) + list(new_df.columns)))
        
        for i in range(max_rows):
            # Get old row data (as dict with all columns) - pass raw values to LLM
            if i < len(old_df):
                old_row = {}
                for col in all_columns:
                    if col in old_df.columns:
                        val = old_df.iloc[i][col]
                        # Convert to string, let LLM handle NaN/None
                        old_row[col] = "" if pd.isna(val) else str(val)
                    else:
                        old_row[col] = ""
            else:
                old_row = {col: "" for col in all_columns}
            
            # Get new row data (as dict with all columns) - pass raw values to LLM
            if i < len(new_df):
                new_row = {}
                for col in all_columns:
                    if col in new_df.columns:
                        val = new_df.iloc[i][col]
                        # Convert to string, let LLM handle NaN/None
                        new_row[col] = "" if pd.isna(val) else str(val)
                    else:
                        new_row[col] = ""
            else:
                new_row = {col: "" for col in all_columns}
            
            old_rows.append(old_row)
            new_rows.append(new_row)
        
        # Use LLM to calculate all changes at once
        change_values = self._compare_all_rows_with_llm(old_rows, new_rows)
        
        # Align dataframes
        aligned_old_df, aligned_new_df = self._align_dataframes(old_df, new_df)
        
        # Create change DataFrame: aligned_new_df with Change column added
        change_df = aligned_new_df.copy()
        
        # Ensure we have enough change values (pad with empty strings if needed)
        while len(change_values) < len(change_df):
            change_values.append("")
        
        # Add Change column
        change_df['Change'] = change_values[:len(change_df)]
        
        return aligned_old_df, aligned_new_df, change_df
    
    def _write_three_tables_to_sheet(self, worksheet, workbook, aligned_old_df, aligned_new_df, change_df):
        """
        Write three tables side-by-side in a single sheet: v0.5, v1, and Change.
        Matches the format shown in the image.
        """
        # Calculate column offsets
        v05_cols = len(aligned_old_df.columns)
        v1_cols = len(aligned_new_df.columns)
        change_cols = len(change_df.columns)
        
        v05_start_col = 0
        v1_start_col = v05_start_col + v05_cols + 2  # +2 for spacing
        change_start_col = v1_start_col + v1_cols + 2  # +2 for spacing
        
        # Define formats
        header_format = workbook.add_format({
            'bg_color': '#90EE90',  # Light green
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })
        row_removed_format = workbook.add_format({
            'bg_color': '#FFE6E6',  # Light red
            'font_color': '#CC0000',  # Dark red
            'text_wrap': True
        })
        row_added_format = workbook.add_format({
            'bg_color': '#E6F7E6',  # Light green
            'font_color': '#006600',  # Dark green
            'text_wrap': True
        })
        value_changed_format = workbook.add_format({
            'bg_color': '#FFF9E6',  # Light yellow
            'font_color': '#CC9900',  # Dark yellow/orange
            'text_wrap': True
        })
        
        # Write headers (row 0)
        # v0.5 header
        worksheet.merge_range(0, v05_start_col, 0, v05_start_col + v05_cols - 1, 'v0.5', header_format)
        # v1 header
        worksheet.merge_range(0, v1_start_col, 0, v1_start_col + v1_cols - 1, 'v1', header_format)
        # Change header
        worksheet.merge_range(0, change_start_col, 0, change_start_col + change_cols - 1, 'Change', header_format)
        
        # Write column headers (row 1)
        # v0.5 column headers
        for col_idx, col_name in enumerate(aligned_old_df.columns):
            worksheet.write(1, v05_start_col + col_idx, col_name)
        
        # v1 column headers
        for col_idx, col_name in enumerate(aligned_new_df.columns):
            worksheet.write(1, v1_start_col + col_idx, col_name)
        
        # Change column headers
        for col_idx, col_name in enumerate(change_df.columns):
            worksheet.write(1, change_start_col + col_idx, col_name)
        
        # Find Change column index in change_df
        change_col_idx = None
        for col_idx, col_name in enumerate(change_df.columns):
            if col_name == 'Change':
                change_col_idx = col_idx
                break
        
        # Write data rows (starting from row 2)
        max_rows = max(len(aligned_old_df), len(aligned_new_df), len(change_df))
        
        for row_idx in range(max_rows):
            # Write v0.5 data
            if row_idx < len(aligned_old_df):
                for col_idx, col_name in enumerate(aligned_old_df.columns):
                    val = aligned_old_df.iloc[row_idx, col_idx]
                    worksheet.write(row_idx + 2, v05_start_col + col_idx, val)
            
            # Write v1 data
            if row_idx < len(aligned_new_df):
                for col_idx, col_name in enumerate(aligned_new_df.columns):
                    val = aligned_new_df.iloc[row_idx, col_idx]
                    worksheet.write(row_idx + 2, v1_start_col + col_idx, val)
            
            # Write Change data with color formatting
            if row_idx < len(change_df):
                change_val = change_df.iloc[row_idx]['Change']
                has_change = change_val and str(change_val).strip()
                
                if has_change:
                    change_str = str(change_val).strip().upper()
                    
                    # Determine format based on change type
                    if 'ROW REMOVED' in change_str:
                        # Format entire row in red
                        for col_idx, col_name in enumerate(change_df.columns):
                            val = change_df.iloc[row_idx, col_idx]
                            worksheet.write(row_idx + 2, change_start_col + col_idx, val, row_removed_format)
                    elif 'ROW ADDED' in change_str:
                        # Format entire row in green
                        for col_idx, col_name in enumerate(change_df.columns):
                            val = change_df.iloc[row_idx, col_idx]
                            worksheet.write(row_idx + 2, change_start_col + col_idx, val, row_added_format)
                    else:
                        # Value changed - format only the Change column cell in yellow
                        for col_idx, col_name in enumerate(change_df.columns):
                            val = change_df.iloc[row_idx, col_idx]
                            if col_idx == change_col_idx:
                                # Highlight only the Change column cell
                                worksheet.write(row_idx + 2, change_start_col + col_idx, val, value_changed_format)
                            else:
                                # Other columns without formatting
                                worksheet.write(row_idx + 2, change_start_col + col_idx, val)
                else:
                    # No change - write all columns without formatting
                    for col_idx, col_name in enumerate(change_df.columns):
                        val = change_df.iloc[row_idx, col_idx]
                        worksheet.write(row_idx + 2, change_start_col + col_idx, val)
    
    def compare_csv_files(self, old_csv_bytes: bytes, new_csv_bytes: bytes) -> dict:
        """
        Compare two CSV files using simple row-by-row comparison.
        
        Returns:
        {
            "result_df": DataFrame with new version + Change column,
            "result_bytes": bytes of the output Excel file,
            "summary": change summary string
        }
        """
        old_df = self.read_csv(old_csv_bytes)
        new_df = self.read_csv(new_csv_bytes)
        
        # Get aligned dataframes and change information
        aligned_old_df, aligned_new_df, change_df = self.compare_dataframes_simple(old_df, new_df)
        
        # Generate output Excel with single sheet containing 3 tables side-by-side
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Comparison')
            
            # Write three tables side-by-side
            self._write_three_tables_to_sheet(worksheet, workbook, aligned_old_df, aligned_new_df, change_df)
        
        # Generate summary
        changes_count = sum(1 for c in change_df['Change'] if c and str(c).strip())
        total_rows = len(change_df)
        summary = f"Compared {total_rows} rows: {changes_count} rows with changes, {total_rows - changes_count} unchanged"
        
        return {
            "result_df": change_df,
            "result_bytes": output.getvalue(),
            "summary": summary,
            "file_type": "csv"
        }
    
    def compare_excel_files(self, old_excel_bytes: bytes, new_excel_bytes: bytes) -> dict:
        """
        Compare two Excel files using simple row-by-row comparison.
        Creates sheets equal to max number of sheets between old and new files.
        Each sheet contains three side-by-side tables: v0.5, v1, and Change.
        
        Returns:
        {
            "result_sheets": {sheet_name: DataFrame with Change column},
            "result_bytes": bytes of the output Excel file,
            "summary": change summary string
        }
        """
        old_sheets = self.read_excel(old_excel_bytes)
        new_sheets = self.read_excel(new_excel_bytes)
        
        # Get all sheet names and determine max number of sheets
        all_sheet_names = set(old_sheets.keys()) | set(new_sheets.keys())
        max_sheets = max(len(old_sheets), len(new_sheets))
        
        # Find matching sheet names for comparison
        matching_sheets = set(old_sheets.keys()) & set(new_sheets.keys())
        
        result_sheets = {}
        total_changes = 0
        total_rows = 0
        
        # Generate output Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Process each sheet (use max number of sheets)
            # For matching sheets, create comparison with 3 tables
            # For non-matching sheets, create empty or single table
            processed_sheets = set()
            
            # Process matching sheets first
            for sheet_name in sorted(matching_sheets):
                old_df = old_sheets[sheet_name]
                new_df = new_sheets[sheet_name]
                
                # Get aligned dataframes and change information
                aligned_old_df, aligned_new_df, change_df = self.compare_dataframes_simple(old_df, new_df)
                result_sheets[sheet_name] = change_df
                
                # Create worksheet with three side-by-side tables
                worksheet = workbook.add_worksheet(sheet_name)
                self._write_three_tables_to_sheet(worksheet, workbook, aligned_old_df, aligned_new_df, change_df)
                
                processed_sheets.add(sheet_name)
                
                # Count changes
                changes_count = sum(1 for c in change_df['Change'] if c and str(c).strip())
                total_changes += changes_count
                total_rows += len(change_df)
            
            # Process remaining sheets (only in old or only in new) up to max_sheets
            remaining_old = set(old_sheets.keys()) - processed_sheets
            remaining_new = set(new_sheets.keys()) - processed_sheets
            
            # Add remaining old sheets
            for sheet_name in sorted(remaining_old)[:max_sheets - len(processed_sheets)]:
                old_df = old_sheets[sheet_name]
                # Create empty new_df with same columns
                empty_new_df = pd.DataFrame(columns=old_df.columns)
                aligned_old_df, aligned_new_df, change_df = self.compare_dataframes_simple(old_df, empty_new_df)
                result_sheets[sheet_name] = change_df
                
                worksheet = workbook.add_worksheet(sheet_name)
                self._write_three_tables_to_sheet(worksheet, workbook, aligned_old_df, aligned_new_df, change_df)
                processed_sheets.add(sheet_name)
                
                # Count changes
                changes_count = sum(1 for c in change_df['Change'] if c and str(c).strip())
                total_changes += changes_count
                total_rows += len(change_df)
            
            # Add remaining new sheets
            for sheet_name in sorted(remaining_new)[:max_sheets - len(processed_sheets)]:
                new_df = new_sheets[sheet_name]
                # Create empty old_df with same columns
                empty_old_df = pd.DataFrame(columns=new_df.columns)
                aligned_old_df, aligned_new_df, change_df = self.compare_dataframes_simple(empty_old_df, new_df)
                result_sheets[sheet_name] = change_df
                
                worksheet = workbook.add_worksheet(sheet_name)
                self._write_three_tables_to_sheet(worksheet, workbook, aligned_old_df, aligned_new_df, change_df)
                processed_sheets.add(sheet_name)
                
                # Count changes
                changes_count = sum(1 for c in change_df['Change'] if c and str(c).strip())
                total_changes += changes_count
                total_rows += len(change_df)
        
        # Build summary
        summary_parts = [f"Processed {len(processed_sheets)} sheet(s): {', '.join(sorted(processed_sheets))}"]
        summary_parts.append(f"Total: {total_rows} rows, {total_changes} rows with changes")
        
        return {
            "result_sheets": result_sheets,
            "result_bytes": output.getvalue(),
            "summary": " | ".join(summary_parts),
            "file_type": "excel"
        }
    
    def get_preview_df(self, df: pd.DataFrame, max_rows: int = 20) -> pd.DataFrame:
        """Get a preview of the DataFrame for display."""
        return df.head(max_rows)
