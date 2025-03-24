# Gmail Label Manager

A Google Apps Script project that allows for the management of Gmail labels via a Google Sheet.

## Features

- Real-time updates to Gmail when labels are modified in the sheet
- Bidirectional sync between Gmail labels and a Google Sheet using a merge sync
- Ability to auto-sync labels upon opening the spreadsheet
- Support for nested labels
- Automatic creation of parent labels
- Menu integration with Google Sheets

## Menu Options

After the script has been enabled, the following options will be available in the 'Gmail Labels' menu:

- **Trigger On Spreadsheet Change**: If enabled, when the spreadsheet is modified, it will then trigger a creation, deletion, or update of the corresponding label in Gmail.
- **Auto Sync On Startup**: If enabled, when the spreadsheet is first opened, it will bi-directionally sync all labels between Gmail and the sheet.
- **Sync All Labels**: This will do a one-time bi-directional sync of all labels between Gmail and the spreadsheet.

## Syncing Strategy

For both **Auto Sync On Startup** and **Sync All Labels**, the script will do a merge sync between Gmail labels and the spreadsheet. This means that:

1. Any labels in the spreadsheet that do not exist in Gmail will be created in Gmail.
2. Any labels in Gmail that do not exist in the spreadsheet will be added to the spreadsheet.
3. If there is a mismatch between the label names in the spreadsheet and Gmail, the script will not delete or modify the labels in either place.

After labels are in sync, and **Trigger On Spreadsheet Change** is enabled, it is best to handle the modification of all labels via the spreadsheet.

Note that if a label has any emails attached to it within Gmail, the script will not allow the label to be deleted from within the spreadsheet.

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
10. Visit your newly created Google Sheet (refresh page if it was already open)
11. Click the 'Gmail Labels' menu item and select 'Click to Enable'
12. Grant the script the required permissions when prompted
13. Click the 'Gmail Labels' menu item and select from the available options

Note: The `.clasp.json` file is in `.gitignore` to prevent accidental commits of personal script IDs. This file ensures files are loaded in the correct order during deployment.

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

## License

This project is licensed under the [MIT License](LICENSE).
