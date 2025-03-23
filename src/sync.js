/**
 * sync.js - Gmail Label Manager
 *
 * Contains bidirectional synchronization functionality for Gmail labels.
 */

/**
 * Updates the bidirectional sync function to properly handle nested labels too
 */
function syncAllLabels() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  let lastRow = sheet.getLastRow();

  // Get existing label map from Gmail
  const labelMap = getLabelMap();

  // Step 1: Create a map of all labels already in the spreadsheet
  const sheetLabelMap = {};
  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const labelName = sheet.getRange(row, CONFIG.NAME_COLUMN).getValue();
    if (labelName) {
      sheetLabelMap[labelName] = row;
    }
  }

  // Step 2: Process existing sheet labels
  for (const labelName in sheetLabelMap) {
    const row = sheetLabelMap[labelName];

    try {
      // Check if label exists in Gmail
      const labelId = labelMap[labelName];

      if (!labelId) {
        // This is a new label to create
        // For nested labels, ensure all parent labels exist first
        if (labelName.includes('/')) {
          const parts = labelName.split('/');
          let parentPath = '';

          // Create each level of the hierarchy if needed
          for (let i = 0; i < parts.length - 1; i++) {
            if (parentPath) {
              parentPath += '/';
            }
            parentPath += parts[i];

            // Create the parent label if it doesn't exist
            const parentId = labelMap[parentPath];
            if (!parentId) {
              GmailApp.createLabel(parentPath);
              logDebug(`Created parent label "${parentPath}"`);
            }
          }
        }

        // Create the label (now that all parents exist if needed)
        GmailApp.createLabel(labelName);

        // Get the new ID
        const newLabelId = getLabelId(labelName);
        if (newLabelId) {
          sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(newLabelId);
          logDebug(`Created label "${labelName}" during sync with ID: ${newLabelId}`);
        }
      } else {
        // Label exists, update the ID
        sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(labelId);
        logDebug(`Updated ID for existing label "${labelName}": ${labelId}`);
      }
    } catch (error) {
      logError(`Error syncing label "${labelName}": ${error.message}`);
    }
  }

  // Step 3: Find Gmail labels not in the spreadsheet
  const response = Gmail.Users.Labels.list('me');
  if (response && response.labels) {
    for (const label of response.labels) {
      // Skip system labels (they start with "CATEGORY_" or have reserved names)
      if (label.type === 'system' ||
        label.name.startsWith('CATEGORY_') ||
        ['INBOX', 'SENT', 'DRAFT', 'TRASH', 'SPAM'].includes(label.name)) {
        continue;
      }

      // Skip labels already in the spreadsheet
      if (sheetLabelMap[label.name]) {
        continue;
      }

      // Add this label to the spreadsheet
      lastRow++;
      sheet.getRange(lastRow, CONFIG.LABEL_ID_COLUMN).setValue(label.id);
      sheet.getRange(lastRow, CONFIG.NAME_COLUMN).setValue(label.name);
      logDebug(`Added Gmail label "${label.name}" to spreadsheet`);
    }
  }

  SpreadsheetApp.getActive().toast('All labels synced bidirectionally between Gmail and spreadsheet');
}
