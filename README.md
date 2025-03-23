# Gmail Label Manager

A Google Apps Script project that allows for the management of Gmail labels via a Google Sheet.

## Features

- Bidirectional sync between Gmail labels and a Google Sheet
- Support for nested labels
- Automatic creation of parent labels
- Real-time updates when labels are modified in the sheet
- Menu integration with Google Sheets

## File Structure

This project has been organized into a modular structure:

```tree
gmail-label-manager/
├── src/                  # Source code files
│   ├── config.js         # Configuration settings
│   ├── logging.js        # Logging utilities
│   ├── ui.js             # User interface and menu creation
│   ├── triggers.js       # Script triggers and event handlers
│   ├── labelling.js      # Label management functionality
│   ├── sync.js           # Bidirectional synchronization
│   └── appsscript.json   # Project manifest with Gmail API service definition
└── .clasp-example.json   # Example Clasp configuration template
```

## Setup and Deployment

1. Create a new Google Sheet or open an existing one
2. Rename your sheet to match the SHEET_NAME in the CONFIG (default is 'Labels')
3. Set up your sheet with these columns:
   - Column A: Label ID (hidden, managed by script)
   - Column B: Label Name
4. Install clasp: `npm install -g @google/clasp`
5. Login to Google: `clasp login`
6. Create a new script: `clasp create --type sheets --title "Gmail Label Manager"`
7. Copy `.clasp-example.json` to `.clasp.json` and set the scriptId to your own script ID
8. Push the code: `clasp push`
9. Open the script: `clasp open`

Note: The `.clasp.json` file is in `.gitignore` to prevent accidental commits of personal script IDs. This file ensures files are loaded in the correct order during deployment.

## License

Feel free to modify and distribute this code as needed.
