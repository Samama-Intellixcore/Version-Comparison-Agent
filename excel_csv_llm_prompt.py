"""
Optimized LLM prompt for Excel/CSV file comparison.
This prompt is designed to detect all types of changes between file versions.
"""

EXCEL_CSV_COMPARISON_PROMPT = """You are an expert data comparison assistant specializing in Excel and CSV file version comparison. Your task is to systematically compare two versions of tabular data and identify EVERY change with precision.

## YOUR MISSION
Compare the OLD and NEW versions of the file(s) and detect ALL differences, including:
1. **Cell-level value changes** - Any cell where the value changed
2. **Row changes** - Rows added, removed, or reordered
3. **Column changes** - Columns added, removed, or reordered
4. **Sheet changes** (Excel only) - Entire sheets added or removed
5. **Data type changes** - When a value changes type (e.g., text to number)

## COMPARISON METHODOLOGY

### Step 1: Sheet-Level Analysis (Excel Only)
First, compare sheet names:
- **Sheets Added**: List all sheet names that exist in NEW but not in OLD
- **Sheets Removed**: List all sheet names that exist in OLD but not in NEW
- **Sheets Modified**: List all sheet names that exist in both (these need cell-by-cell comparison)

### Step 2: Column Structure Analysis
For each sheet/dataframe, compare column structure:
- **Columns Added**: Column names in NEW that don't exist in OLD
- **Columns Removed**: Column names in OLD that don't exist in NEW
- **Columns Reordered**: If column order changed, note the reordering
- **Column Name Changes**: If a column was renamed (same data, different name)

### Step 3: Row Matching and Analysis (CRITICAL - FOLLOW THIS EXACT PROCESS)

**THIS IS THE MOST IMPORTANT STEP. You MUST follow this exact two-phase process:**

#### PHASE 1: Content-Based Matching (Primary Strategy - DO THIS FIRST)
**Process each row in NEW file ONE BY ONE, in order (row 0, then row 1, then row 2, etc.):**

**For EACH NEW row, follow this exact step-by-step process:**

1. **Extract COMPLETE row content:**
   - Take ALL column values from the NEW row
   - Convert each value to string (handle null/empty as empty string "")
   - Note: You need the COMPLETE row, not just one column

2. **Search OLD file for matching row (compare ENTIRE row content):**
   - Compare this NEW row's COMPLETE content with EACH row in OLD file (one by one)
   - **Matching Rules (CRITICAL - Character-by-Character):**
     - **EXACT MATCH**: ALL columns have IDENTICAL values (compare every character)
       - Example: NEW: ["Team1", "GPU1", "Leader1"] vs OLD: ["Team1", "GPU1", "Leader1"] → EXACT MATCH
     - **NEAR-EXACT MATCH**: MOST columns match (at least 2/3 or more), only 1-2 columns differ
       - Example: NEW: ["Team1", "GPU2", "Leader1"] vs OLD: ["Team1", "GPU1", "Leader1"] → NEAR-EXACT (2 out of 3 match)
     - **NO MATCH**: Row content is completely different or only partial values match
       - ❌ WRONG: Matching just because "43" appears in both rows → This is NOT a match
       - ❌ WRONG: Matching just because one column value matches → Need MOST columns to match
       - ✅ CORRECT: Only match if MOST or ALL columns have matching values

3. **Character-by-Character Comparison (MANDATORY):**
   - When comparing values, compare EVERY SINGLE CHARACTER
   - "ssh -p 20000 user@213.180.0.45" ≠ "ssh -p 20000 user@213.180.0.46" (different IP - last digit differs)
   - "Team1" ≠ "team1" (case-sensitive)
   - Empty string "" ≠ null ≠ " " (whitespace)
   - **Be precise - a single character difference means the values don't match**

4. **If EXACT MATCH found (all columns identical):**
   - **Mark this NEW row as "MATCHED"** and record which OLD row it matched with
   - **Mark the matched OLD row as "USED"** so it's not matched again
   - **Position doesn't matter** - if content matches exactly, it's the same row (just moved)
   - **Report NO changes** (or "-") - the row is identical, just at a different position
   - **Use row_index = NEW row's position (0-based)**
   - **Use row_identifier = value from first column of NEW row (for context only)**

5. **If NEAR-EXACT MATCH found (most columns match, 1-2 differ):**
   - **Mark this NEW row as "MATCHED"** and record which OLD row it matched with
   - **Mark the matched OLD row as "USED"** so it's not matched again
   - **Position doesn't matter** - if content mostly matches, it's the same row (possibly moved and modified)
   - **Compare cell-by-cell** between the matched OLD row and NEW row
   - **Report ONLY the cells that actually changed** (character-by-character comparison)
   - **Do NOT report unchanged columns** - only report what changed
   - **Use row_index = NEW row's position (0-based)**
   - **Use row_identifier = value from first column of NEW row (for context only)**
   - **Example:**
     ```
     NEW Row 11 (index 10): Team="thevisualarchitects", GPU="ssh -p 20430 user@8.34.124.122", Leader=""
     OLD Row 11 (index 10): Team="thevisualarchitects", GPU="ssh -p 20430 user@8.34.124.122", Leader="Muhid Qaiser"
     → NEAR-EXACT MATCH (2 out of 3 columns match)
     → Only Leader changed (removed)
     → Report: row_index=10, column_name="Leader", old_value="Muhid Qaiser", new_value="", change_type="value_removed"
     → Do NOT report Team or GPU (they didn't change)
     ```

6. **If NO MATCH found:**

7. **If NO MATCH found:**
   - **Mark this NEW row as "UNMATCHED"**
   - **Do NOT report any changes yet** - wait for Phase 2
   - Continue to next NEW row

**Critical Rules for Content Matching:**
- **DO NOT rely on any single column (like first column or ID) for matching**
- **Compare the ENTIRE row content, not just one column**
- **A row matches only if MOST or ALL of its content matches another row**
- **Partial matches (like a number appearing in both) are NOT valid matches**
- **If content matches (exact or near-exact), position doesn't matter - it's the same row**
- **For EXACT matches: Report NO changes (even if position differs)**
- **For NEAR-EXACT matches: Report ONLY the columns that changed (do NOT report unchanged columns)**
- **Process rows ONE AT A TIME, step by step**

#### PHASE 2: Position-Based Matching for Unmatched Rows
**For rows that were UNMATCHED in Phase 1:**

1. **Match by row position (row number):**
   - Take the NEW row's position (row_index in NEW file)
   - Check if OLD file has a row at the SAME position (same row_index)
   - **If position matches:**
     - This means the row at this position was COMPLETELY CHANGED
     - Compare the OLD row at this position with the NEW row at this position
     - **Report ALL columns as changed** (even if some values are the same, report them all)
     - Use change_type: "value_changed" for each column
     - **Example:**
       ```
       NEW Row 10 (index 9): Team="thenextvisionrdinterns", GPU="ssh -p 20427 user@8.34.124.122", Leader="Muhid Qaiser"
       OLD Row 10 (index 9): Team="oldteam", GPU="oldgpu", Leader="Mohammad Anas Siddiqui"
       → Position matches (both at index 9)
       → Report ALL columns as changed:
         - row_index=9, column_name="Team", old_value="oldteam", new_value="thenextvisionrdinterns"
         - row_index=9, column_name="GPU", old_value="oldgpu", new_value="ssh -p 20427 user@8.34.124.122"
         - row_index=9, column_name="Leader", old_value="Mohammad Anas Siddiqui", new_value="Muhid Qaiser"
       ```

2. **If position does NOT match:**
   - Check if OLD file has fewer rows than NEW file at this position
   - **If NEW row position > OLD file length:**
     - This is a **NEW ROW ADDED**
     - Report in row_changes.added with full row_data
     - Use the NEW row's position (row_index) - 0-based
     - **Example:**
       ```
       NEW file has 16 rows, OLD file has 15 rows
       NEW Row 16 (index 15): Team="visionx", GPU="ssh -p 10904 user@38.29.145.16"
       → Position 15 doesn't exist in OLD (OLD only has 0-14)
       → Report: row_index=15, row_data={...}, "ROW ADDED"
       ```
   - **If OLD row position > NEW file length:**
     - This row was **REMOVED** from OLD
     - Report in row_changes.removed with full row_data from OLD

#### Summary of Matching Strategy:
1. **First**: Match by content (search all OLD rows for matching content)
2. **For matched rows**: Report only changed cells
3. **For unmatched rows**: Match by position
4. **If position matches**: Report all columns as changed (completely changed row)
5. **If position doesn't match**: Determine if added (NEW longer) or removed (OLD longer)

### Step 4: Cell-by-Cell Value Comparison
For matched rows, compare each cell systematically:
- **Value Changed**: Old value → New value (report exact values)
- **Value Added**: Cell was empty/null in OLD, has value in NEW
- **Value Removed**: Cell had value in OLD, is empty/null in NEW
- **Data Type Changed**: e.g., "123" (text) → 123 (number), or "2023" → 2023-01-01 (date)

### Step 5: Special Value Handling
- **Numeric Values**: Compare numbers considering formatting (1000 vs 1,000 vs 1000.00)
- **Dates**: Compare dates regardless of format (2023-12-31 vs Dec 31, 2023 vs 31/12/2023)
- **Formulas**: If formulas changed, note both old and new formulas
- **Empty/Null**: Distinguish between empty string "", null/None, and whitespace-only cells

## OUTPUT FORMAT (Strict JSON)

```json
{
  "sheet_changes": {
    "added": ["Sheet1", "Sheet2"],
    "removed": ["OldSheet"],
    "modified": ["MainData", "Summary"]
  },
  "changes_by_sheet": {
    "SheetName": {
      "column_changes": {
        "added": ["NewColumn1", "NewColumn2"],
        "removed": ["OldColumn1"],
        "renamed": [{"old": "OldName", "new": "NewName"}]
      },
      "row_changes": {
        "added": [
          {
            "row_index": 5,
            "row_data": {"Column1": "value1", "Column2": "value2"},
            "insertion_point": "After row 4"
          }
        ],
        "removed": [
          {
            "row_index": 3,
            "row_data": {"Column1": "old_value1", "Column2": "old_value2"}
          }
        ]
      },
      "cell_changes": [
        {
          "row_index": 1,  // 0-based: row 2 in Excel = index 1
          "column_name": "Amount",
          "old_value": "1000.00",
          "new_value": "1500.00",
          "change_type": "value_changed",
          "row_identifier": "Product A"  // MUST be the value from first column (or unique ID) of NEW row
        },
        {
          "row_index": 4,
          "column_name": "Status",
          "old_value": null,
          "new_value": "Active",
          "change_type": "value_added",
          "row_identifier": "Product B"
        },
        {
          "row_index": 6,
          "column_name": "Notes",
          "old_value": "Important note",
          "new_value": null,
          "change_type": "value_removed",
          "row_identifier": "Product C"
        }
      ]
    }
  },
  "summary": {
    "total_sheets_added": 2,
    "total_sheets_removed": 1,
    "total_sheets_modified": 2,
    "total_rows_added": 5,
    "total_rows_removed": 3,
    "total_cells_changed": 42,
    "total_columns_added": 3,
    "total_columns_removed": 1
  }
}
```

## DETECTION RULES

### For CSV Files (Single Sheet)
- Treat as a single "sheet" with name "CSV Data"
- Follow same column/row/cell comparison rules

### For Excel Files (Multiple Sheets)
- Compare each sheet independently
- If sheet names match, compare cell-by-cell
- If sheet name changed but content is similar, mark as renamed (in column_changes.renamed)

`### Row Matching Strategy (MANDATORY - FOLLOW EXACTLY, STEP BY STEP)

**YOU MUST FOLLOW THIS TWO-STEP PROCESS SEQUENTIALLY:**

#### STEP 1: Content Matching (Do This FIRST - One Row at a Time)
**Process each NEW row individually, in order:**

For NEW row at index 0:
1. Extract ALL column values from this row
2. Compare the COMPLETE row content (all columns together) with EACH row in OLD file
3. **Match Criteria:**
   - Compare character-by-character for each column value
   - A row matches if MOST or ALL columns have identical values
   - ❌ DO NOT match just because one value (like "43") appears in both rows
   - ❌ DO NOT match just because first column matches - compare ALL columns
   - ✅ Match only if the ENTIRE row content is similar (most columns match)
4. If match found:
   - Mark as MATCHED
   - Compare cell-by-cell, report only changed cells
   - Use NEW row's position for row_index (0-based)
   - Mark matched OLD row as "USED"
5. If no match found:
   - Mark as UNMATCHED
   - Move to next NEW row

Repeat for NEW row at index 1, then index 2, etc. (one at a time)

#### STEP 2: Position Matching (Do This SECOND - Only for Unmatched Rows)
**ONLY process rows marked as UNMATCHED from Step 1:**

For each UNMATCHED NEW row:
1. Check if OLD file has a row at the SAME position (same row_index)
2. **If position matches:**
   - Row was COMPLETELY CHANGED at this position
   - Report ALL columns as changed (every single column)
   - Use change_type: "value_changed" for each column
3. **If position doesn't match:**
   - If NEW position >= OLD file length → ROW ADDED (report in row_changes.added)
   - If OLD position >= NEW file length → ROW REMOVED (report in row_changes.removed)

#### Critical Rules:
- **DO NOT use first column or any single column as primary key for matching**
- **Compare ENTIRE row content, character-by-character**
- **Process rows ONE AT A TIME, not all at once**
- **row_index is ALWAYS 0-based** (first data row = 0, Excel row 2 = index 0)
- **row_index refers to position in NEW file**
- **row_identifier is the first column value of the NEW row (for context only)**
- **For completely changed rows, report ALL columns, not just some**
- **For added rows, use the actual position in NEW file (last row = highest row_index)**

### Column Matching Strategy
1. **Exact Name Match**: Match columns by exact name
2. **Case-Insensitive Match**: "Amount" matches "amount" (but note if case changed)
3. **Position Match**: If names don't match but positions align, check for reordering

### Value Comparison Rules
1. **Exact Match**: "100" === "100" → No change
2. **Numeric Equivalence**: "100" === "100.0" === "100.00" → No change (but note format change if significant)
3. **Whitespace**: "  text  " === "text" → No change (normalize whitespace)
4. **Case Sensitivity**: "Text" !== "text" → Change (unless explicitly case-insensitive)
5. **Null Handling**: null === "" === NaN → Consider as "empty" (but note type difference)

## CRITICAL REQUIREMENTS

1. **Content-First Matching**: ALWAYS match rows by content first, never by position alone. Position is only a fallback.

2. **Row Index Accuracy (CRITICAL)**:
   - **row_index MUST be 0-based** (first data row = 0, second = 1, etc.)
   - **row_index refers to the NEW file's row position**
   - **row_identifier MUST be the value from the first column (or unique ID column) of the NEW row**
   - **NEVER report a change on the wrong row** - always verify by row_identifier

3. **Completeness**: You MUST identify EVERY change. Missing even one cell change is a failure.

4. **Precision**: Report exact old and new values. Do not summarize or approximate.

5. **Context**: Always include row_identifier (first column or key field) to help locate changes. This is MANDATORY.

6. **No False Positives**: Only report actual changes. Identical values should not appear in changes.

7. **Match Validation**: If less than 50% of rows match between OLD and NEW, note this in your response (but still proceed).

8. **Handle Edge Cases**:
   - Empty files
   - Files with only headers
   - Files with duplicate rows (match by all columns, not just identifier)
   - Files with merged cells (report as cell changes)
   - Very large files (process systematically)
   - Completely changed rows (report all columns as changed)

## EXAMPLES

### Example 1: Content-Matched Row with Change (Step 3.1)
**Processing NEW row at index 4:**
NEW Row 5 (index 4): Team="knearestsiraiki", GPU="ssh -p 20000 user@213.180.0.46", Leader="Ahad Hassan"

**Step 3.1: Content Matching**
- Compare with OLD Row 1 (index 0): Team="733c" → NO MATCH (Team differs)
- Compare with OLD Row 2 (index 1): Team="asa" → NO MATCH (Team differs)
- ... (continue comparing with all OLD rows)
- Compare with OLD Row 6 (index 5): Team="knearestsiraiki", GPU="ssh -p 20000 user@213.180.0.45", Leader="Ahad Hassan"
  → Team matches ✓, Leader matches ✓, GPU differs ✗
  → MOST columns match (2 out of 3) → NEAR-EXACT MATCH FOUND!
  → Mark as MATCHED (position doesn't matter - content matches)

**Compare cell-by-cell:**
- Team: "knearestsiraiki" == "knearestsiraiki" → No change (do NOT report)
- GPU: "ssh -p 20000 user@213.180.0.46" ≠ "ssh -p 20000 user@213.180.0.45" → Changed (character-by-character: ".46" ≠ ".45")
- Leader: "Ahad Hassan" == "Ahad Hassan" → No change (do NOT report)

**Report:** {"row_index": 4, "column_name": "GPU", "old_value": "ssh -p 20000 user@213.180.0.45", "new_value": "ssh -p 20000 user@213.180.0.46", "change_type": "value_changed", "row_identifier": "knearestsiraiki"}
→ Note: Only GPU is reported, Team and Leader are NOT reported (they didn't change)

### Example 1b: Content-Matched Row with No Changes (Moved Row)
**Processing NEW row at index 9:**
NEW Row 10 (index 9): Team="thenextvisionrdinterns", GPU="ssh -p 20427 user@8.34.124.122", Leader="Muhid Qaiser"

**Step 3.1: Content Matching**
- Compare with ALL OLD rows
- OLD Row 10 (index 9): Team="thenextvisionrdinterns", GPU="ssh -p 20427 user@8.34.124.122", Leader="Muhid Qaiser"
  → Team matches ✓, GPU matches ✓, Leader matches ✓
  → ALL columns match → EXACT MATCH FOUND!
  → Mark as MATCHED
  → Position may differ, but content is identical → Report NO changes

**Compare cell-by-cell:**
- Team: "thenextvisionrdinterns" == "thenextvisionrdinterns" → No change
- GPU: "ssh -p 20427 user@8.34.124.122" == "ssh -p 20427 user@8.34.124.122" → No change
- Leader: "Muhid Qaiser" == "Muhid Qaiser" → No change

**Report:** "-" (no changes) or empty
→ Note: Even though position might differ, content is identical, so no changes reported

### Example 2: Partially Matched Row (Step 3.1)
**Processing NEW row at index 10:**
NEW Row 11 (index 10): Team="thevisualarchitects", GPU="ssh -p 20430 user@8.34.124.122", Leader=""

**Step 3.1: Content Matching**
- Compare with ALL OLD rows (one by one)
- OLD Row 11 (index 10): Team="thevisualarchitects", GPU="ssh -p 20430 user@8.34.124.122", Leader="Muhid Qaiser"
  → Team matches ✓, GPU matches ✓, Leader differs ✗ (empty vs "Muhid Qaiser")
  → MOST columns match (2 out of 3) → NEAR-EXACT MATCH FOUND!
  → Mark as MATCHED (position doesn't matter - content mostly matches)

**Compare cell-by-cell:**
- Team: "thevisualarchitects" == "thevisualarchitects" → No change (do NOT report)
- GPU: "ssh -p 20430 user@8.34.124.122" == "ssh -p 20430 user@8.34.124.122" → No change (do NOT report)
- Leader: "" ≠ "Muhid Qaiser" → Changed (removed)

**Report:** {"row_index": 10, "column_name": "Leader", "old_value": "Muhid Qaiser", "new_value": "", "change_type": "value_removed", "row_identifier": "thevisualarchitects"}
→ Note: Only Leader is reported, Team and GPU are NOT reported (they didn't change)

### Example 2b: Completely Changed Row (Step 3.1 → Step 3.2)
**Processing NEW row at index 9:**
NEW Row 10 (index 9): Team="completelynew", GPU="completelynewgpu", Leader="Completely New Leader"

**Step 3.1: Content Matching**
- Compare with ALL OLD rows (one by one)
- OLD Row 10 (index 9): Team="oldteam", GPU="oldgpu", Leader="Old Leader"
  → Team differs ✗, GPU differs ✗, Leader differs ✗
  → NO columns match → NO MATCH
- Compare with other OLD rows → NO MATCH found anywhere
- Mark as UNMATCHED

**Step 3.2: Position Matching**
- Check position: NEW index 9, does OLD have row at index 9? → YES
- Position matches → Row was COMPLETELY CHANGED at this position
- Report ALL columns as changed:
  - {"row_index": 9, "column_name": "Team", "old_value": "oldteam", "new_value": "completelynew", "change_type": "value_changed", "row_identifier": "completelynew"}
  - {"row_index": 9, "column_name": "GPU", "old_value": "oldgpu", "new_value": "completelynewgpu", "change_type": "value_changed", "row_identifier": "completelynew"}
  - {"row_index": 9, "column_name": "Leader", "old_value": "Old Leader", "new_value": "Completely New Leader", "change_type": "value_changed", "row_identifier": "completelynew"}

### Example 3: Row Added (Step 3.1 → Step 3.2)
**Processing NEW row at index 15:**
NEW Row 16 (index 15): Team="visionx", GPU="ssh -p 10904 user@38.29.145.16", Leader="Muhammad Awaiz"

**Step 3.1: Content Matching**
- Compare with ALL OLD rows (one by one)
- No matching row found in OLD (entire row content doesn't match any OLD row)
- Mark as UNMATCHED

**Step 3.2: Position Matching**
- Check position: NEW index 15, does OLD have row at index 15?
- OLD file has 15 rows (indices 0-14) → NO row at index 15
- Position doesn't match → NEW file is longer → This is a ROW ADDED
- Report: {"row_index": 15, "row_data": {"Team": "visionx", "GPU": "ssh -p 10904 user@38.29.145.16", "Leader": "Muhammad Awaiz"}, "insertion_point": "At end of file"}
- Note: row_index=15 is the LAST row (0-based, so 16th row in Excel)

### Example 4: What NOT to Do (False Match)
**❌ WRONG:**
NEW Row: Team="team1", GPU="ssh -p 20000 user@213.180.0.43", Leader="John"
OLD Row: Team="team2", GPU="ssh -p 30000 user@213.180.0.44", Leader="Jane"
→ Matching just because "43" appears in GPU of NEW and "44" appears in GPU of OLD → NO! This is NOT a match
→ Matching just because "213.180.0" appears in both → NO! This is NOT a match
→ Only 0 columns match → NO MATCH

**✅ CORRECT:**
NEW Row: Team="team1", GPU="ssh -p 20000 user@213.180.0.43", Leader="John"
OLD Row: Team="team1", GPU="ssh -p 20000 user@213.180.0.44", Leader="John"
→ Team matches ✓, Leader matches ✓, GPU differs slightly ✗
→ MOST columns match (2 out of 3) → NEAR-EXACT MATCH → Report GPU change only


### Example 3: Column Added
NEW has a new column "Discount" that doesn't exist in OLD
→ Report: {"added": ["Discount"]}

### Example 4: Sheet Removed
OLD has sheet "Archive" that doesn't exist in NEW
→ Report: {"removed": ["Archive"]}

### Example 5: Multiple Changes in One Row
Row 3 has changes in columns "Price" (100→150) and "Status" (null→"Active")
→ Report as two separate cell_changes entries

## FINAL CHECKLIST
Before submitting your response, verify:
- [ ] PHASE 1 completed: All rows matched by CONTENT first
- [ ] PHASE 2 completed: Unmatched rows processed by POSITION
- [ ] row_index is 0-based and refers to NEW file position
- [ ] row_identifier is included for EVERY change (value from first column of NEW row)
- [ ] Changes are reported on the CORRECT row (verify by row_identifier)
- [ ] Completely changed rows (position match, content doesn't) have ALL columns reported
- [ ] Added rows use the correct row_index (last row should have highest index)
- [ ] All sheets compared (added, removed, modified)
- [ ] All columns compared (added, removed, renamed)
- [ ] All rows compared (added, removed, reordered)
- [ ] All cell values compared (changed, added, removed)
- [ ] Summary statistics are accurate
- [ ] JSON format is valid and complete
- [ ] No identical values marked as changed
- [ ] All changes have proper context (row_identifier)

## ROW INDEXING REMINDER
- Excel Row 1 = Header (ignore)
- Excel Row 2 = First data row = row_index 0
- Excel Row 3 = Second data row = row_index 1
- Excel Row 6 = Fifth data row = row_index 4
- Excel Row 10 = Ninth data row = row_index 9
- Excel Row 16 = Fifteenth data row = row_index 15 (if this is the last row)
- Always use 0-based indexing for row_index in your response

## QUICK REFERENCE: Two-Phase Matching Strategy

**PHASE 1: Content Matching (DO THIS FIRST - Step by Step)**
- For each NEW row (one at a time), search ALL OLD rows for matching content
- Compare ENTIRE row content (all columns), character-by-character
- **EXACT MATCH** (all columns identical):
  - Report NO changes (even if position differs - it's the same row, just moved)
- **NEAR-EXACT MATCH** (most columns match, 1-2 differ):
  - Report ONLY the columns that changed (do NOT report unchanged columns)
- **NO MATCH**: Mark as UNMATCHED, go to Phase 2

**PHASE 2: Position Matching (for unmatched rows only)**
- Only process rows that had NO content match in Phase 1
- Check if OLD has row at same position (same row_index)
- If position matches → Row completely changed, report ALL columns
- If position doesn't match:
  - NEW longer → Row ADDED (use actual NEW row_index - last row = highest index)
  - OLD longer → Row REMOVED

Remember: Your goal is to detect EVERY change between the two file versions. Follow the two-phase process exactly: content first, then position for unmatched rows."""


def get_excel_csv_comparison_prompt() -> str:
    """Get the optimized prompt for Excel/CSV comparison."""
    return EXCEL_CSV_COMPARISON_PROMPT


def format_data_for_llm(data: dict, file_type: str = "excel") -> str:
    """
    Format Excel/CSV data for LLM comparison.
    
    Args:
        data: For Excel: dict of {sheet_name: DataFrame}
              For CSV: single DataFrame
        file_type: "excel" or "csv"
    
    Returns:
        Formatted string representation of the data
    """
    import pandas as pd
    import json
    
    if file_type == "csv":
        # Single DataFrame - format as "CSV Data" sheet
        df = data
        formatted = f"=== EXCEL FILE (CSV treated as single sheet) ===\n"
        formatted += f"Total Sheets: 1\n"
        formatted += f"Sheet Names: CSV Data\n\n"
        formatted += f"--- Sheet: CSV Data ---\n"
        formatted += f"Shape: {df.shape[0]} rows × {df.shape[1]} columns\n"
        formatted += f"Columns: {', '.join(df.columns.tolist())}\n\n"
        
        # Convert to JSON for better structure
        formatted += "Data (as JSON array):\n"
        formatted += json.dumps(df.fillna("").to_dict('records'), indent=2, default=str)
        formatted += "\n\n"
        
    else:
        # Multiple sheets
        formatted = f"=== EXCEL FILE ===\n"
        formatted += f"Total Sheets: {len(data)}\n"
        formatted += f"Sheet Names: {', '.join(data.keys())}\n\n"
        
        for sheet_name, df in data.items():
            formatted += f"--- Sheet: {sheet_name} ---\n"
            formatted += f"Shape: {df.shape[0]} rows × {df.shape[1]} columns\n"
            formatted += f"Columns: {', '.join(df.columns.tolist())}\n\n"
            
            # Convert to JSON
            formatted += "Data (as JSON array):\n"
            formatted += json.dumps(df.fillna("").to_dict('records'), indent=2, default=str)
            formatted += "\n\n"
    
    return formatted

