# 🔄 Version Comparison Agent

An AI-powered document version comparison tool that compares files across PDF, Excel, and CSV formats, highlighting changes between versions.

## Features

- **PDF Comparison**: Compare two PDF versions with color-coded highlights
  - 🔴 Red: Removed content
  - 🟢 Green: Added content
  - 🟡 Yellow: Modified content
  - Generates two highlighted output PDFs (old and new versions)

- **Excel Comparison**: Compare Excel files with multiple sheet support
  - Compares all sheets with matching names
  - Row-by-row cell comparison
  - Single output file with change column

- **CSV Comparison**: Compare CSV files
  - Row-by-row cell comparison
  - Single output file with change column

## Prerequisites

- Python 3.10+
- Azure OpenAI API access (GPT-4.1 deployment)

## Installation

1. Clone or download this repository

2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   # or
   source venv/bin/activate  # Linux/Mac
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Create a `.env` file in the project root with your Azure OpenAI credentials:
   ```env
   AZURE_OPENAI_API_KEY=your_api_key_here
   AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
   AZURE_OPENAI_DEPLOYMENT=your_deployment_name
   AZURE_OPENAI_API_VERSION=2024-02-15-preview
   ```

## Usage

Run the Streamlit application:

```bash
streamlit run app.py
```

The application will open in your default browser at `http://localhost:8501`.

### PDF Comparison

1. Select the "PDF Comparison" tab
2. Upload the older PDF version
3. Upload the newer PDF version
4. Click "Compare PDF Versions"
5. Download the highlighted PDFs

### Excel/CSV Comparison

1. Select the "Excel / CSV Comparison" tab
2. Choose the file type (Excel or CSV)
3. Upload the older file version
4. Upload the newer file version
5. Optionally toggle AI-powered comparison
6. Click "Compare Files"
7. Download the comparison result file

## File Structure

```
Version Comparison Agent/
├── app.py                    # Main Streamlit application
├── config.py                 # Configuration management
├── llm_client.py             # Azure OpenAI LLM client
├── pdf_comparator.py         # PDF comparison logic
├── excel_csv_comparator.py   # Excel/CSV comparison logic
├── requirements.txt          # Python dependencies
├── .env                      # Environment variables (create this)
└── README.md                 # This file
```

## Configuration

| Environment Variable | Description |
|---------------------|-------------|
| `AZURE_OPENAI_API_KEY` | Your Azure OpenAI API key |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_DEPLOYMENT` | Deployment name (e.g., gpt-4) |
| `AZURE_OPENAI_API_VERSION` | API version (default: 2024-02-15-preview) |

## Limitations

- Maximum file size: 20 MB per file
- PDF comparison works best with text-based PDFs (not scanned documents)
- Excel comparison only processes sheets with matching names in both files
- Large files may take longer to process due to LLM API calls

## Troubleshooting

**"Missing configuration" error:**
- Ensure your `.env` file exists and contains all required variables
- Check that the values are correctly formatted

**PDF highlighting not working:**
- The LLM may not find exact text matches for highlighting
- Complex PDF layouts may affect text extraction accuracy

**Slow performance:**
- Disable "Use AI for intelligent comparison" for faster processing
- Large files are processed in batches which may take time

## License

MIT License

