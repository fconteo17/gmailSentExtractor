# Gmail Export Tool

A Python tool with a modern dark-themed GUI for exporting sent emails from Gmail to Excel spreadsheets. This tool allows you to manage multiple Gmail accounts and export sent emails within specified date ranges to well-formatted Excel files.

## Features

- Modern dark-themed graphical user interface
- Command-line interface (CLI) mode for automation and scripting
- Multiple Gmail account support with:
  - Checkbox-based account selection
  - Bulk account operations (select all, remove multiple)
  - Visual account management
- Date range selection with calendar widgets
- Well-formatted Excel exports with:
  - Sorted emails by date
  - Auto-adjusted column widths
  - Styled headers
  - Export metadata
- Real-time progress tracking with:
  - Progress bar
  - Email count updates
  - Status messages
- Comprehensive error handling and logging
- Automatic last month date range selection

## Prerequisites

- Python 3.8 or higher
- Google Cloud Project with Gmail API enabled
- OAuth 2.0 Client ID credentials

## Setup

1. Clone the repository:
```bash
git clone https://github.com/fconteo17/gmailSentExtractor.git
cd gmailSentExtractor
```

2. Create and activate a virtual environment:
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Set up Google Cloud Project:
   - Go to [Google Cloud Console](https://console.cloud.google.com)
   - Create a new project or select an existing one
   - Enable the Gmail API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download the credentials and save as `credentials.json` in the `data/config` directory

## Usage

### Graphical Interface Mode

1. Run the application in GUI mode (default):
```bash
python run.py
```

2. Using the GUI:
   - Click "Add Account" to add a new email account
   - Follow the authentication flow in your browser
   - Select one or more accounts using checkboxes
   - Choose date range (defaults to last month)
   - Click "Export Emails" to start the export

3. GUI Features:
   - Use "Select All" to select all accounts
   - Remove multiple accounts at once
   - Monitor progress in real-time
   - Find exported files in the `data/exports` directory

### Command Line Mode

1. Run the application in CLI mode:
```bash
python run.py --cli
```

2. CLI Features:
   - Interactive command-line interface
   - Text-based menu system
   - Progress displayed in terminal
   - Same functionality as GUI mode
   - Suitable for automation and scripting

## Directory Structure

```
gmail-export-tool/
├── data/
│   ├── config/
│   │   └── credentials.json
│   ├── exports/
│   └── tokens/
├── src/
│   ├── __init__.py
│   ├── account_manager.py
│   ├── config.py
│   ├── export.py
│   ├── file_manager.py
│   ├── gmail_service.py
│   ├── gui.py
│   └── main.py
├── README.md
└── requirements.txt
```

## Security

- OAuth 2.0 tokens are stored securely in the `data/tokens` directory
- No email passwords are stored
- Minimal required permissions are requested
- Tokens can be revoked at any time through Google Account settings

## Error Handling

The tool includes comprehensive error handling for:
- Network issues
- Authentication failures
- Invalid inputs
- File system errors
- API rate limits

## License

This project is licensed under the MIT License - see the LICENSE file for details. 