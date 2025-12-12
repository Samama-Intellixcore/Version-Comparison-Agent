"""
Azure OpenAI LLM client for the Version Comparison Agent.
"""
import json
import httpx
from openai import AzureOpenAI
from config import Config


class LLMClient:
    """Client for Azure OpenAI interactions."""
    
    def __init__(self):
        # Create httpx client without proxies to avoid compatibility issues
        http_client = httpx.Client()
        
        self.client = AzureOpenAI(
            api_key=Config.AZURE_OPENAI_API_KEY,
            api_version=Config.AZURE_OPENAI_API_VERSION,
            azure_endpoint=Config.AZURE_OPENAI_ENDPOINT,
            http_client=http_client
        )
        self.deployment = Config.AZURE_OPENAI_DEPLOYMENT
    
    def compare_text_content(self, old_text: str, new_text: str) -> dict:
        """
        Compare two text contents and identify changes.
        
        Returns a structured response with:
        - removed: List of text segments removed from old version
        - added: List of text segments added in new version
        - modified: List of {old: str, new: str} for modified segments
        """
        system_prompt = """You are a precise document comparison assistant. Your task is to compare two versions of a document and identify all changes.

Analyze the OLD and NEW document texts and return a JSON object with:
1. "removed": Array of text segments that exist in OLD but not in NEW
2. "added": Array of text segments that exist in NEW but not in OLD  
3. "modified": Array of objects with {"old": "original text", "new": "modified text"} for text that was changed

Rules:
- Be precise and capture exact text differences
- For modified text, match corresponding segments that were changed (not completely removed/added)
- Ignore minor whitespace differences
- Return valid JSON only, no additional text
- Group contiguous changes together when possible"""

        user_prompt = f"""Compare these two document versions:

=== OLD VERSION ===
{old_text}

=== NEW VERSION ===
{new_text}

Return the comparison as JSON with "removed", "added", and "modified" arrays."""

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0,
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            return {
                "removed": result.get("removed", []),
                "added": result.get("added", []),
                "modified": result.get("modified", [])
            }
        except Exception as e:
            raise Exception(f"LLM comparison failed: {str(e)}")
    
    def compare_tabular_data(self, old_data: list[list], new_data: list[list], 
                             headers: list[str] = None) -> list[list]:
        """
        Compare two tabular datasets row by row.
        
        Returns a change matrix where:
        - "-" means no change
        - Value means the difference/change description
        """
        system_prompt = """You are a precise data comparison assistant. Compare two tabular datasets row by row and cell by cell.

For each cell, determine if there's a change between the old and new version:
- If no change: output "-"
- If changed: output the change description (e.g., the difference value or a short description)

Return a JSON object with:
- "changes": A 2D array (list of rows, each row is a list of cells) containing the change indicators
- Each cell should be "-" for no change or the change value/description for changes"""

        user_prompt = f"""Compare these datasets row by row:

=== OLD DATA ===
{json.dumps(old_data, indent=2)}

=== NEW DATA ===
{json.dumps(new_data, indent=2)}

{f"Headers: {headers}" if headers else ""}

Return JSON with "changes" as a 2D array matching the structure of the new data."""

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0,
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            return result.get("changes", [])
        except Exception as e:
            raise Exception(f"LLM comparison failed: {str(e)}")
    
    def compare_tabular_batch(self, old_data: list[list], new_data: list[list],
                               headers: list[str] = None, batch_size: int = 50) -> list[list]:
        """
        Compare tabular data in batches to handle large datasets.
        """
        all_changes = []
        
        # Determine the max rows to process
        max_rows = max(len(old_data), len(new_data))
        
        for i in range(0, max_rows, batch_size):
            old_batch = old_data[i:i + batch_size] if i < len(old_data) else []
            new_batch = new_data[i:i + batch_size] if i < len(new_data) else []
            
            # Handle case where one dataset is shorter
            if not old_batch:
                old_batch = [[""] * len(new_batch[0])] * len(new_batch) if new_batch else []
            if not new_batch:
                new_batch = [[""] * len(old_batch[0])] * len(old_batch) if old_batch else []
            
            batch_changes = self.compare_tabular_data(
                old_batch, new_batch, 
                headers if i == 0 else None
            )
            all_changes.extend(batch_changes)
        
        return all_changes
    
    def compare_pdf_content(self, old_content: str, new_content: str) -> dict:
        """
        Compare two PDF document contents that may include text and tables.
        Returns structured changes for highlighting.
        """
        system_prompt = """You are a document comparison expert. Compare OLD and NEW versions of the SAME document.

## DOCUMENT STRUCTURE
Documents contain:
- Text paragraphs (headings, body text, signatures)
- Tables with structure:
  ┌─ TABLE N ─┐
    [HEADER]: Column1 │ Column2 │ Column3
    ──────────────────────────────────────
    [Row Label]: Value1 │ Value2 │ Value3
    [Another Row]: Value1 │ Value2 │ Value3
  └─ /TABLE N ─┘

## OUTPUT FORMAT (strict JSON)

{
  "modified": [
    {"field": "context/label", "old": "old value only", "new": "new value only"}
  ],
  "added": [
    {"text": "new content with no equivalent in OLD"}
  ],
  "removed": [
    {"text": "deleted content with no equivalent in NEW"}
  ]
}

## CLASSIFICATION RULES

### MODIFIED (most common)
When SAME field/position exists in both but VALUE changed.

TEXT EXAMPLES:
- "Company No. 00802030" → "Company No. 00802031"
  → {"field": "Company Registration No.", "old": "00802030", "new": "00802031"}
  
- "John Smith (Senior Auditor)" → "Marc Cowell CA (Senior Statutory Auditor)"  
  → {"field": "Auditor Name", "old": "John Smith", "new": "Marc Cowell CA"}

TABLE EXAMPLES:
- Row labeled [Gas combustion]: OLD has "416,000" | NEW has "416,011"
  → {"field": "Gas combustion", "old": "416,000", "new": "416,011"}

- Row labeled [Total gross emissions]: OLD has "450.00" | NEW has "476.85"
  → {"field": "Total gross emissions", "old": "450.00", "new": "476.85"}

- Row 5 (no label): OLD has "100,000" | NEW has "150,000"
  → {"field": "Table Row 5", "old": "100,000", "new": "150,000"}

DATE/YEAR:
- "31 December 2023" → "31 December 2024"
  → {"field": "Year End Date", "old": "2023", "new": "2024"}

CRITICAL: Return ONLY the specific changed value, NOT the entire line/row.

### ADDED (rare)
ONLY for genuinely NEW content that has NO equivalent in OLD:
- A completely new paragraph
- A new table row that didn't exist
- A new section

Do NOT mark as added if similar content exists with different values (that's MODIFIED).

### REMOVED (rare)
ONLY for content completely DELETED from OLD with NO equivalent in NEW:
- A paragraph that no longer exists
- A table row that was removed
- A section that was deleted

### REMOVED (rare - only for completely DELETED content)  
Use ONLY when OLD has content that NEW does not have AT ALL.
- A deleted paragraph
- A removed table row
- A removed section

## MUST IGNORE (never report these)
- "DRAFT" watermark
- Page numbers
- Headers/footers that appear in both
- Identical text in both documents
- Formatting/whitespace differences
- Structural markers (PAGE, TABLE labels)

## VERIFICATION
Before including ANY item:
1. Is it actually DIFFERENT between OLD and NEW? If identical → skip
2. Am I returning just the VALUE, not the whole line? If whole line → extract just the value
3. For ADDED: Does OLD truly have NOTHING similar? If OLD has it → probably MODIFIED
4. For REMOVED: Does NEW truly have NOTHING similar? If NEW has it → probably MODIFIED"""

        user_prompt = f"""Compare these two document versions line by line:

=== OLD VERSION ===
{old_content}

=== NEW VERSION ===
{new_content}

Find ALL differences. For each:
1. Identify what changed
2. Extract ONLY the specific value that changed (not whole lines)
3. Classify correctly (modified/added/removed)
4. Skip identical content and watermarks

Return JSON with "modified", "added", "removed" arrays."""

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0,
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            return {
                "removed": result.get("removed", []),
                "added": result.get("added", []),
                "modified": result.get("modified", [])
            }
        except Exception as e:
            raise Exception(f"LLM comparison failed: {str(e)}")
    
    def compare_excel_csv_files(self, old_data_formatted: str, new_data_formatted: str) -> dict:
        """
        Compare Excel/CSV files using the optimized comparison prompt.
        
        Args:
            old_data_formatted: Formatted string representation of old file data
            new_data_formatted: Formatted string representation of new file data
        
        Returns:
            Dictionary with structured changes including:
            - sheet_changes: Added/removed/modified sheets (Excel only)
            - changes_by_sheet: Detailed changes per sheet
            - summary: Summary statistics
        """
        from excel_csv_llm_prompt import get_excel_csv_comparison_prompt
        
        system_prompt = get_excel_csv_comparison_prompt()
        
        user_prompt = f"""Compare these two file versions and identify ALL changes.

{old_data_formatted}

{new_data_formatted}

Follow the comparison methodology systematically:
1. First, identify sheet-level changes (if Excel)
2. Then, analyze column structure changes
3. Next, analyze row structure changes
4. Finally, perform cell-by-cell value comparison
5. Generate comprehensive summary statistics

Return the complete JSON response with all detected changes."""

        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0,
                response_format={"type": "json_object"}
            )
            
            result = json.loads(response.choices[0].message.content)
            return result
        except Exception as e:
            raise Exception(f"LLM Excel/CSV comparison failed: {str(e)}")
    
    def process_pdf_base64(self, base64_pdf: str, prompt: str = None) -> str:
        """
        Process a base64-encoded PDF file using the LLM vision capabilities.
        
        Args:
            base64_pdf: Base64-encoded string of the PDF file
            prompt: Optional custom prompt. If None, uses a default prompt.
        
        Returns:
            The LLM's response as a string
        """
        if prompt is None:
            prompt = "Analyze this PDF document and provide a detailed summary of its contents, including any text, tables, and key information."
        
        # Create data URI for the PDF
        data_uri = f"data:application/pdf;base64,{base64_pdf}"
        
        try:
            response = self.client.chat.completions.create(
                model=self.deployment,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": data_uri
                                }
                            }
                        ]
                    }
                ],
                temperature=0
            )
            
            return response.choices[0].message.content
        except Exception as e:
            raise Exception(f"LLM PDF processing failed: {str(e)}")

