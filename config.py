"""
Configuration management for the Version Comparison Agent.
"""
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


class Config:
    """Application configuration."""
    
    # Azure OpenAI Settings
    AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
    AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
    AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")
    AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
    
    # File size limit (20 MB)
    MAX_FILE_SIZE_MB = 20
    MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
    
    # Supported file types
    SUPPORTED_PDF_TYPES = [".pdf"]
    SUPPORTED_EXCEL_TYPES = [".xlsx", ".xls"]
    SUPPORTED_CSV_TYPES = [".csv"]
    
    @classmethod
    def validate(cls) -> tuple[bool, str]:
        """Validate that all required configuration is present."""
        missing = []
        
        if not cls.AZURE_OPENAI_API_KEY:
            missing.append("AZURE_OPENAI_API_KEY")
        if not cls.AZURE_OPENAI_ENDPOINT:
            missing.append("AZURE_OPENAI_ENDPOINT")
        if not cls.AZURE_OPENAI_DEPLOYMENT:
            missing.append("AZURE_OPENAI_DEPLOYMENT")
        
        if missing:
            return False, f"Missing configuration: {', '.join(missing)}"
        
        return True, "Configuration valid"

