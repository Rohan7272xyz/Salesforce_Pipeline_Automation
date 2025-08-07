import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
import sys

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config import Config

def generate_gantt_chart(input_filepath: str) -> str:
    """
    Process the input Excel file using app.py and return the path to the final output file.
    
    Args:
        input_filepath (str): Path to the input Excel file
        
    Returns:
        str: Path to the processed output file
        
    Raises:
        FileNotFoundError: If the expected output file is not created
        subprocess.CalledProcessError: If app.py fails to run
    """
    try:
        # Validate input file exists
        input_path = Path(input_filepath)
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_filepath}")
        
        print(f"📤 Processing file: {input_path.name}")
        
        # Get the path to app.py relative to this script
        app_py_path = project_root / "app.py"
        if not app_py_path.exists():
            raise FileNotFoundError(f"app.py not found at: {app_py_path}")
        
        # Run app.py with the input file
        print("🔄 Running app.py...")
        result = subprocess.run(
            ["python", str(app_py_path), str(input_filepath)], 
            check=True,
            capture_output=True,
            text=True,
            cwd=str(project_root)
        )
        
        print("✅ app.py completed successfully")
        if result.stdout:
            print(f"Output: {result.stdout}")
        
        # Check for the expected output file in Downloads
        template_output_path = Config.get_downloads_path() / Config.DEFAULT_OUTPUT_NAME
        
        if not template_output_path.exists():
            raise FileNotFoundError(f"Expected output file not found at: {template_output_path}")

        # Create timestamped final output file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_output_path = Config.OUTPUT_DIR / f"Pipeline_GanttChart_{timestamp}.xlsx"
        
        # Ensure output directory exists
        Config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        # Move the file from Downloads to our output directory
        shutil.move(str(template_output_path), str(final_output_path))
        print(f"✅ Moved final output to: {final_output_path}")
        
        return str(final_output_path)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ app.py failed with return code {e.returncode}")
        print(f"Error output: {e.stderr}")
        raise
    except Exception as e:
        print(f"❌ Error in generate_gantt_chart: {e}")
        raise

def cleanup_temp_files():
    """Clean up temporary files older than 7 days."""
    try:
        import time
        current_time = time.time()
        seven_days_ago = current_time - (7 * 24 * 60 * 60)
        
        # Clean up input directory
        for file_path in Config.INPUT_DIR.glob("*"):
            if file_path.is_file() and file_path.stat().st_mtime < seven_days_ago:
                file_path.unlink()
                print(f"🗑️ Cleaned up old temp file: {file_path.name}")
                
        # Clean up output directory (keep recent files)
        for file_path in Config.OUTPUT_DIR.glob("Pipeline_GanttChart_*"):
            if file_path.is_file() and file_path.stat().st_mtime < seven_days_ago:
                file_path.unlink()
                print(f"🗑️ Cleaned up old output file: {file_path.name}")
                
    except Exception as e:
        print(f"⚠️ Warning: Cleanup failed: {e}")

if __name__ == "__main__":
    try:
        Config.validate_config()
        print("✅ Configuration validated")
    except ValueError as e:
        print(f"❌ Configuration error: {e}")
        sys.exit(1)
    
    if len(sys.argv) < 2:
        print("Usage: python process.py <input_file_path>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    try:
        output_path = generate_gantt_chart(input_file)
        print(f"✅ Processing completed. Output file: {output_path}")
        
        # Clean up old files
        cleanup_temp_files()
        
    except Exception as e:
        print(f"❌ Processing failed: {e}")
        sys.exit(1)