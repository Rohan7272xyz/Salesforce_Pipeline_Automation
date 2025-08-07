import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class Config:
    """Centralized configuration for the pipeline automation system."""
    
    # Project root directory
    PROJECT_ROOT = Path(__file__).parent.absolute()
    
    # Email Configuration
    EMAIL_USER = os.getenv("EMAIL_USER", "magpipelinemanager@gmail.com")
    EMAIL_PASS = os.getenv("EMAIL_PASS")
    IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.gmail.com")
    IMAP_PORT = int(os.getenv("IMAP_PORT", 993))
    SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
    
    # Authorized Users
    AUTHORIZED_EMAILS = ["Rohan.Anand@mag.us"]
    
    # Contact Information
    YOUR_EMAIL = "Rohan.Anand@mag.us"
    JOE_EMAIL = "Rohan.Anand@mag.us"
    
    # File Paths (relative to project root)
    TEMPLATE_FILENAME = "C5SDEC_Pipeline_Overview_v3_070325.xlsx"
    TEMPLATE_PATH = PROJECT_ROOT / TEMPLATE_FILENAME
    
    # Directory Paths
    INPUT_DIR = PROJECT_ROOT / "input"
    OUTPUT_DIR = PROJECT_ROOT / "output"
    BACKUP_DIR = PROJECT_ROOT / "template_backups"
    TEMP_DIR = PROJECT_ROOT / "temp_templates"
    LOGS_DIR = PROJECT_ROOT / "logs"
    WEB_DIR = PROJECT_ROOT / "web"
    
    # Processing Configuration
    TEMPLATE_SHEET_NAME = 'Pipeline'
    DATA_START_ROW = 5
    CHECK_INTERVAL_SECONDS = 5
    DEFAULT_OUTPUT_NAME = "C5S&DEC_Pipeline_FINAL_SORTED.xlsx"
    
    # File Extensions
    EXCEL_EXTENSIONS = ['.xlsx', '.xls']
    
    @classmethod
    def ensure_directories(cls):
        """Create all necessary directories if they don't exist."""
        directories = [
            cls.INPUT_DIR,
            cls.OUTPUT_DIR, 
            cls.BACKUP_DIR,
            cls.TEMP_DIR,
            cls.LOGS_DIR
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
    
    @classmethod
    def validate_config(cls):
        """Validate that all required configuration is present."""
        errors = []
        
        # Check required environment variables
        if not cls.EMAIL_PASS:
            errors.append("EMAIL_PASS environment variable not set")
        
        # Check template file exists
        if not cls.TEMPLATE_PATH.exists():
            errors.append(f"Template file not found: {cls.TEMPLATE_PATH}")
        
        # Check web directory exists (for eel GUI)
        if not cls.WEB_DIR.exists():
            errors.append(f"Web directory not found: {cls.WEB_DIR}")
        
        if errors:
            raise ValueError("Configuration errors:\n" + "\n".join(f"- {error}" for error in errors))
        
        return True
    
    @classmethod
    def get_downloads_path(cls):
        """Get the user's Downloads directory."""
        return Path.home() / "Downloads"

# Create directories on import
Config.ensure_directories()