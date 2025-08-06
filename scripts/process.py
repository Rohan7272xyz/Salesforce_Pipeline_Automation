import os
import shutil
import subprocess
from datetime import datetime

OUTPUT_DIR = "output"
DEFAULT_OUTPUT_NAME = "C5S&DEC_Pipeline_FINAL_SORTED.xlsx"
USER_DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
TEMPLATE_OUTPUT_PATH = os.path.join(USER_DOWNLOADS, DEFAULT_OUTPUT_NAME)

os.makedirs(OUTPUT_DIR, exist_ok=True)

def generate_gantt_chart(input_filepath: str) -> str:
    print("📤 Running app.py...")
    subprocess.run(["python", "app.py", input_filepath], check=True)

    if not os.path.exists(TEMPLATE_OUTPUT_PATH):
        raise FileNotFoundError(f"Expected output file not found at: {TEMPLATE_OUTPUT_PATH}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_output_path = os.path.join(OUTPUT_DIR, f"Pipeline_GanttChart_{timestamp}.xlsx")

    shutil.move(TEMPLATE_OUTPUT_PATH, final_output_path)
    print(f"✅ Moved final output to: {final_output_path}")
    return final_output_path

if __name__ == "__main__":
    test_file = "input/sample_pipeline.xlsx"
    path = generate_gantt_chart(test_file)
    print(f"Use this path in send_email.py: {path}")
