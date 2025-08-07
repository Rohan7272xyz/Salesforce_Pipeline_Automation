@echo off
REM Fixed Pipeline Runner - Uses relative paths and error handling

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Change to the project directory
cd /d "%SCRIPT_DIR%"

REM Log file in the logs directory
set LOG_FILE=%SCRIPT_DIR%logs\pipeline_output.txt

REM Create logs directory if it doesn't exist
if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"

REM Write startup info to log
echo [%DATE% %TIME%] Pipeline automation starting... >> "%LOG_FILE%"
echo Script directory: %SCRIPT_DIR% >> "%LOG_FILE%"
echo Working directory: %CD% >> "%LOG_FILE%"

REM Try to find Python executable
set PYTHON_EXE=python

REM Check if python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [%DATE% %TIME%] ERROR: Python not found in PATH >> "%LOG_FILE%"
    echo ERROR: Python not found. Make sure Python is installed and in your PATH.
    pause
    exit /b 1
)

REM Log Python version
echo [%DATE% %TIME%] Using Python: >> "%LOG_FILE%"
python --version >> "%LOG_FILE%" 2>&1

REM Run the pipeline
echo [%DATE% %TIME%] Starting pipeline/run_pipeline.py >> "%LOG_FILE%"
python "pipeline\run_pipeline.py" >> "%LOG_FILE%" 2>&1

REM Check if the script ran successfully
if %errorlevel% neq 0 (
    echo [%DATE% %TIME%] ERROR: Pipeline script failed with error code %errorlevel% >> "%LOG_FILE%"
    echo ERROR: Pipeline script failed. Check the log file: %LOG_FILE%
    pause
    exit /b %errorlevel%
)

echo [%DATE% %TIME%] Pipeline completed successfully >> "%LOG_FILE%"
echo Pipeline completed successfully!

REM Keep window open for debugging (remove in production)
REM pause