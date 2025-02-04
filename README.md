# Gmail Export Tool

A Python tool for exporting sent emails from Gmail to Excel spreadsheets. This tool allows you to manage multiple Gmail accounts and export sent emails within specified date ranges to well-formatted Excel files.

## Features

- Multiple Gmail account support
- Date range filtering for exports
- Well-formatted Excel exports with:
  - Sorted emails by date
  - Auto-adjusted column widths
  - Styled headers
  - Export metadata
- Progress tracking for long operations
- Comprehensive error handling and logging

## Prerequisites

- Python 3.8 or higher
- Google Cloud Project with Gmail API enabled
- OAuth 2.0 Client ID credentials

## Setup

1. Clone the repository:
```bash
git clone <repository-url>
cd gmail-export-tool
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

1. Run the application:
```bash
python run.py
```

2. First-time setup:
   - Choose option 1 to add a new email account
   - Follow the authentication flow in your browser
   - Grant the necessary permissions

3. Export emails:
   - Choose option 2 to use an existing account
   - Select the account to use
   - Enter the date range for the export
   - Wait for the export to complete
   - Find the exported file in the `data/exports` directory

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
│   ├── main.py
│   └── ui.py
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