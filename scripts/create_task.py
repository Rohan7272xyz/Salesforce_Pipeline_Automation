import os
import subprocess

# Define paths
task_name = "EmailPipelineAutoRun"
script_dir = r"C:\Users\rohan\Personal Projects\Email_Excel_Python_Alg"
bat_file = os.path.join(script_dir, "run_pipeline.bat")

# Task command
command = [
    "schtasks",
    "/Create",
    "/F",  # Force overwrite if task exists
    "/TN", task_name,
    "/TR", f'"{bat_file}"',
    "/SC", "ONLOGON",
    "/RL", "HIGHEST",
    "/RU", os.getlogin(),  # current user
]

# Run the command
try:
    result = subprocess.run(command, check=True, capture_output=True, text=True)
    print(f"✅ Task '{task_name}' created successfully!")
    print(result.stdout)
except subprocess.CalledProcessError as e:
    print(f"❌ Failed to create task:\n{e.stderr}")
