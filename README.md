# Salesforce Pipeline Automation System

An intelligent email-based automation system that processes Salesforce pipeline reports and provides formatted Excel outputs with automatic template management capabilities.

## Overview

This system monitors a Gmail account for incoming Salesforce pipeline reports, automatically processes them through custom formatting and analysis algorithms, and returns professionally formatted Excel files to the sender. It includes advanced features like automatic error handling, template management, and resilient column structure adaptation.

## Key Features

### Automated Pipeline Processing
- Monitors Gmail for new Excel attachments from authorized senders
- Processes raw Salesforce data through custom algorithms
- Generates formatted pipeline reports with Gantt chart layouts
- Automatically emails processed files back to the sender
- Always CCs the system administrator for transparency

### Intelligent Error Management
- Real-time error detection and logging
- Automatic email alerts to system administrator
- Professional error notifications to stakeholders
- Comprehensive logging with timestamps
- System continues running even after errors

### Dynamic Template Management
- **Subject: "Adjust Columns"** → System sends current template to user for modification
- **Subject: "Here"** → System updates template with user's modified version
- Automatic template backups with timestamps
- Zero-downtime template updates
- Future-proof column structure adaptation

### Email-Based Interface
- No GUI required - everything managed through email
- Professional automated responses
- Clear instructions for users
- Error notifications with actionable guidance

## System Architecture

### Core Files
- `fetch.py` - Email monitoring and attachment downloading
- `app.py` - Excel data processing and formatting engine
- `send_email.py` - Email sending with attachment and CC support
- `process.py` - File management and workflow orchestration
- `run_pipeline.py` - Main pipeline execution loop

### Advanced Features
- `template_manager.py` - Template update workflow management
- `integrated_pipeline.py` - Combined system with all features
- `create_task.py` - Windows task scheduler integration

### Configuration
- `.env` - Environment variables and credentials
- `run_pipeline.bat` - Windows batch execution script

## Setup Instructions

### Prerequisites
```bash
pip install pandas openpyxl python-dotenv imaplib-ssl
```

### Environment Configuration
Create a `.env` file with:
```env
EMAIL_USER=your-gmail-account@gmail.com
EMAIL_PASS=your-app-password
IMAP_SERVER=imap.gmail.com
IMAP_PORT=993
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```

### Template Setup
Ensure the Excel template exists at:
```
C:\Users\rohan\Personal Projects\Email_Excel_Python_Alg\C5SDEC_Pipeline_Overview_v3_070325.xlsx
```

## Usage

### Normal Operation
```bash
python run_pipeline.py          # Original system
python integrated_pipeline.py   # Enhanced system with template management
```

### Template Management Workflow

#### To Update Column Structure:
1. **Boss sends email** with subject: `Adjust Columns`
2. **System responds** with current template file
3. **Boss modifies** template to match new data structure
4. **Boss replies** with subject: `Here` + modified template attached
5. **System updates** template automatically and confirms

#### Normal Pipeline Processing:
1. **Send Excel file** to the monitored Gmail account
2. **System processes** data automatically
3. **Receive formatted report** back via email
4. **Administrator gets CC** for all transactions

## Processed Data Structure

The system handles Salesforce pipeline reports with columns:
- Capture Manager
- Opportunity Name
- Salesforce ID
- Stage
- Positioning
- Contract Ceiling Value
- MAG Value (dynamically added)
- Anticipated RFP Date
- Award Date
- GovWin IQ ID

## Security Features

- **Authorized sender validation** - Only processes emails from approved addresses
- **Automatic template backups** - Every update creates timestamped backup
- **Error containment** - System continues running after individual failures
- **Comprehensive logging** - Full audit trail of all operations
- **Email-only interface** - No exposed web endpoints or APIs

## Directory Structure

```
Email_Excel_Python_Alg/
├── README.md
├── .env
├── *.py                           # Core system files
├── input/                         # Downloaded attachments
├── output/                        # Processed files
├── template_backups/              # Automatic template versions
├── temp_templates/                # Temporary template files
├── web/                          # GUI assets (optional)
├── pipeline_log.txt              # System operation logs
└── C5SDEC_Pipeline_Overview_v3_070325.xlsx  # Active template
```

## Maintenance

### Log Monitoring
Check `pipeline_log.txt` for:
- Successful email processing confirmations
- Error notifications and details
- Template update confirmations

### Template Backups
Located in `template_backups/` with format:
```
template_backup_YYYYMMDD_HHMMSS.xlsx
```

### Error Recovery
- System automatically continues after errors
- Administrator receives immediate email alerts
- Stakeholders get professional error notifications

## Business Value

- **Zero-touch operation** - Fully automated pipeline processing
- **Future-proof design** - Adapts to changing data structures
- **Professional communication** - Automated stakeholder notifications  
- **Audit compliance** - Complete logging and backup systems
- **Business continuity** - System operates independently of administrator presence

## Technical Notes

- **Resilient architecture** - Modular design allows independent testing
- **Backward compatibility** - Original system remains unchanged
- **Email-based configuration** - No technical knowledge required for template updates
- **Windows integration** - Task scheduler support for automatic startup
- **Error isolation** - Template management and pipeline processing are independent

## Support

For technical issues or system modifications, contact the system administrator. The system includes automatic error notifications and professional stakeholder communication for business continuity.

---

**System Status:** Production Ready  
**Last Updated:** August 2025  
**Compatibility:** Windows 10/11, Python 3.7+