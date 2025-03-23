/**
 * Gmail Label Manager
 * 
 * This script synchronizes Gmail labels with a Google Sheet.
 * It handles creating, updating, and deleting labels based on changes to the sheet.
 * 
 * IMPORTANT SETUP:
 * 1. Rename your sheet to match the SHEET_NAME in the CONFIG (default is 'Labels')
 * 2. Ensure your columns match the exact order specified in CONFIG:
 *    - Column A: Label ID (hidden, managed by script)
 *    - Column B: Label Name
 * 3. Add any additional columns you like within the Google Sheet, such as 'Description'
 *    or 'Additional Instruction' for managing your labels in Gmail
 * 4. Enable the Gmail Advanced Service:
 *    - In the Apps Script editor, click on "Services" (+ button)
 *    - Select "Gmail API" from the list and click "Add"
 *    - Save the script
 */

// Configuration
const CONFIG = {
  SHEET_NAME: 'Labels',      // The name of your sheet - MUST match your actual sheet name
  HEADER_ROW: 1,             // The row containing headers
  LABEL_ID_COLUMN: 1,        // Hidden column A for storing label IDs
  NAME_COLUMN: 2,            // Column B (1-indexed)
  DEBUG_MODE: true           // Set to false in production
};

/**
 * Creates the menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmail Labels')
    .addItem('Sync All Labels', 'syncAllLabels')
    .addItem('Install Edit Trigger', 'installEditTrigger')
    .addToUi();
  
  // Set up the label ID column
  setupLabelIdColumn();
}

/**
 * Sets up the hidden column for label IDs
 */
function setupLabelIdColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  
  // Create header for ID column if needed
  sheet.getRange(CONFIG.HEADER_ROW, CONFIG.LABEL_ID_COLUMN).setValue('Label ID');
  
  // Try to hide the column
  try {
    sheet.hideColumns(CONFIG.LABEL_ID_COLUMN);
    logDebug("Label ID column hidden successfully");
  } catch (error) {
    logError(`Failed to hide Label ID column: ${error.message}`);
  }
}

/**
 * Installs the edit trigger to monitor changes to the sheet
 */
function installEditTrigger() {
  // Delete any existing edit triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create new edit trigger
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  SpreadsheetApp.getActive().toast('Edit trigger installed successfully');
}

/**
 * Handler for edit events in the spreadsheet
 */
function onEditTrigger(e) {
  try {
    // Get sheet, row, and column of edit
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Only process edits in our target sheet
    if (sheetName !== CONFIG.SHEET_NAME) return;
    
    const row = e.range.getRow();
    const column = e.range.getColumn();
    
    // Skip header row
    if (row <= CONFIG.HEADER_ROW) return;
    
    // Only process changes to the name column
    if (column === CONFIG.NAME_COLUMN) {
      logDebug(`Change detected in Name column at row ${row}`);
      
      const newLabelName = e.value || "";
      const oldLabelName = e.oldValue || "";
      
      handleLabelChange(sheet, row, oldLabelName, newLabelName);
    }
  } catch (error) {
    logError(`Error in onEditTrigger: ${error.message}`);
  }
}

/**
 * Gets a mapping of all Gmail label names to their IDs
 * @return {Object} Map of label name -> label ID
 */
function getLabelMap() {
  try {
    // Use the Gmail Advanced Service to get all labels
    const response = Gmail.Users.Labels.list('me');
    
    if (!response || !response.labels) {
      logError("No labels found in the Gmail account");
      return {};
    }
    
    // Create a map of label name -> label ID
    const labelMap = {};
    for (const label of response.labels) {
      labelMap[label.name] = label.id;
    }
    
    logDebug(`Found ${Object.keys(labelMap).length} Gmail labels`);
    return labelMap;
  } catch (error) {
    logError(`Error getting label map: ${error.message}`);
    return {};
  }
}

/**
 * Gets the ID of a Gmail label by name
 * @param {string} labelName The name of the label
 * @return {string|null} The label ID or null if not found
 */
function getLabelId(labelName) {
  const labelMap = getLabelMap();
  return labelMap[labelName] || null;
}

/**
 * Processes label changes (create, update, delete)
 */
function handleLabelChange(sheet, row, oldLabelName, newLabelName) {
  logDebug(`Processing label change: "${oldLabelName}" -> "${newLabelName}"`);
  
  // Determine the action based on the change
  if (oldLabelName === "" && newLabelName !== "") {
    // Create a new label
    createLabel(sheet, row, newLabelName);
  } else if (oldLabelName !== "" && newLabelName === "") {
    // Delete a label
    deleteLabel(sheet, row, oldLabelName);
  } else if (oldLabelName !== "" && newLabelName !== "" && oldLabelName !== newLabelName) {
    // Update a label
    updateLabel(sheet, row, oldLabelName, newLabelName);
  }
}

/**
 * Creates a new Gmail label, supporting nested labels and adding parent labels to spreadsheet
 */
function createLabel(sheet, row, labelName) {
  logDebug(`Creating label: "${labelName}"`);
  
  try {
    // Check if label already exists
    const existingId = getLabelId(labelName);
    if (existingId) {
      logDebug(`Label "${labelName}" already exists with ID: ${existingId}`);
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(existingId);
      // Toast notification for existing label
      SpreadsheetApp.getActive().toast(`Label "${labelName}" already exists in Gmail.`, 'Info', 3);
      return;
    }
    
    // Check if this is a nested label and add parent labels to spreadsheet
    addParentLabelsToSheet(sheet, labelName);
    
    // Create the label (now that all parents exist)
    GmailApp.createLabel(labelName);
    
    // Get the ID of the newly created label
    const newLabelId = getLabelId(labelName);
    
    if (newLabelId) {
      // Store the label ID in the hidden column
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(newLabelId);
      logDebug(`Label "${labelName}" created successfully with ID: ${newLabelId}`);
      
      // Toast notification for new label
      SpreadsheetApp.getActive().toast(`New label "${labelName}" created in Gmail.`, 'Success', 3);
    } else {
      logError(`Failed to retrieve ID for newly created label "${labelName}"`);
      SpreadsheetApp.getActive().toast(`Created label but couldn't retrieve its ID`, 'Warning', 5);
    }
  } catch (error) {
    logError(`Error creating label "${labelName}": ${error.message}`);
    SpreadsheetApp.getActive().toast(`Error creating label "${labelName}": ${error.message}`, 'Error', 10);
  }
}

/**
 * Updates an existing Gmail label, with support for nested labels
 */
function updateLabel(sheet, row, oldLabelName, newLabelName) {
  logDebug(`Updating label: "${oldLabelName}" -> "${newLabelName}"`);
  
  try {
    // Get the ID of the old label
    const oldLabelId = getLabelId(oldLabelName);
    
    // If old label not found, just create the new one
    if (!oldLabelId) {
      logDebug(`Old label "${oldLabelName}" not found, creating new label "${newLabelName}" instead`);
      createLabel(sheet, row, newLabelName);
      return;
    }
    
    // Check if this is a nested label and add parent labels to spreadsheet
    addParentLabelsToSheet(sheet, newLabelName);
    
    // Get all threads with the old label
    const threads = GmailApp.search(`label:${oldLabelName}`);
    logDebug(`Found ${threads.length} threads with label "${oldLabelName}"`);
    
    // Create the new label
    GmailApp.createLabel(newLabelName);
    
    // Get the ID of the new label
    const newLabelId = getLabelId(newLabelName);
    if (!newLabelId) {
      logError(`Failed to get ID for new label "${newLabelName}"`);
      SpreadsheetApp.getActive().toast(`Error updating label: couldn't create new label`, 'Error', 10);
      return;
    }
    
    // Get the actual label objects
    const oldLabel = GmailApp.getUserLabelByName(oldLabelName);
    const newLabel = GmailApp.getUserLabelByName(newLabelName);
    
    if (!oldLabel || !newLabel) {
      logError(`Failed to get label objects for "${oldLabelName}" or "${newLabelName}"`);
      return;
    }
    
    // Process threads in batches
    const BATCH_SIZE = 100;
    for (let i = 0; i < threads.length; i += BATCH_SIZE) {
      const batch = threads.slice(i, i + BATCH_SIZE);
      
      // Add the new label to threads
      if (batch.length > 0) {
        newLabel.addToThreads(batch);
      }
    }
    
    // Remove the old label from threads in batches
    for (let i = 0; i < threads.length; i += BATCH_SIZE) {
      const batch = threads.slice(i, i + BATCH_SIZE);
      
      if (batch.length > 0) {
        oldLabel.removeFromThreads(batch);
      }
    }
    
    // Delete the old label
    oldLabel.deleteLabel();
    
    // Update the label ID in the spreadsheet
    sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(newLabelId);
    
    logDebug(`Label updated successfully from "${oldLabelName}" to "${newLabelName}"`);
    
    // Toast notification for label rename
    SpreadsheetApp.getActive().toast(`The label "${oldLabelName}" has been renamed to "${newLabelName}" within Gmail.`, 'Success', 5);
  } catch (error) {
    logError(`Error updating label from "${oldLabelName}" to "${newLabelName}": ${error.message}`);
    SpreadsheetApp.getActive().toast(`Error updating label: ${error.message}`, 'Error', 10);
  }
}

/**
 * Deletes a Gmail label
 */
function deleteLabel(sheet, row, labelName) {
  logDebug(`Attempting to delete label: "${labelName}"`);
  
  try {
    // Get the label ID
    const labelId = getLabelId(labelName);
    
    // If label doesn't exist, just clear the row and return
    if (!labelId) {
      logDebug(`Label "${labelName}" not found, nothing to delete`);
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).clearContent();
      return;
    }
    
    // Get the label object
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      logDebug(`Label object for "${labelName}" not found`);
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).clearContent();
      return;
    }
    
    // Check if any threads use this label
    const threads = GmailApp.search(`label:${labelName}`);
    
    if (threads.length > 0) {
      // There are threads with this label, send notification
      const message = `Cannot delete label "${labelName}" as it still has ${threads.length} threads using it.`;
      logDebug(message);
      
      // Show warning in spreadsheet
      SpreadsheetApp.getActive().toast(message, 'Warning', 10);
      
      // Restore the label name in the spreadsheet
      sheet.getRange(row, CONFIG.NAME_COLUMN).setValue(labelName);
    } else {
      // No threads with this label, proceed with deletion
      label.deleteLabel();
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).clearContent();
      logDebug(`Label "${labelName}" deleted successfully`);
      
      // Toast notification for label deletion
      SpreadsheetApp.getActive().toast(`The label "${labelName}" has been deleted from Gmail.`, 'Info', 5);
    }
  } catch (error) {
    logError(`Error deleting label "${labelName}": ${error.message}`);
    SpreadsheetApp.getActive().toast(`Error deleting label: ${error.message}`, 'Error', 10);
  }
}

/**
 * Adds parent labels to the spreadsheet for a nested label
 */
function addParentLabelsToSheet(sheet, labelName) {
  // Only process if this is a nested label
  if (!labelName.includes('/')) {
    return;
  }
  
  // Create a map of existing sheet labels to check for parents
  const sheetLabelMap = {};
  const lastRow = sheet.getLastRow();
  for (let r = CONFIG.HEADER_ROW + 1; r <= lastRow; r++) {
    const currentLabel = sheet.getRange(r, CONFIG.NAME_COLUMN).getValue();
    if (currentLabel) {
      sheetLabelMap[currentLabel] = r;
    }
  }
  
  // Process the nested structure
  const parts = labelName.split('/');
  let parentPath = '';
  let parentsAdded = 0;
  
  // Create each level of the hierarchy if needed
  for (let i = 0; i < parts.length - 1; i++) {
    if (parentPath) {
      parentPath += '/';
    }
    parentPath += parts[i];
    
    // Create the parent label in Gmail if it doesn't exist
    const parentId = getLabelId(parentPath);
    if (!parentId) {
      GmailApp.createLabel(parentPath);
      logDebug(`Created parent label "${parentPath}" in Gmail`);
      SpreadsheetApp.getActive().toast(`New parent label "${parentPath}" created in Gmail.`, 'Info', 3);
    }
    
    // Add the parent label to the spreadsheet if it doesn't exist
    if (!sheetLabelMap[parentPath]) {
      // Append to the end of the sheet
      const newRow = lastRow + 1 + parentsAdded;
      
      // Add the parent label to the new row
      sheet.getRange(newRow, CONFIG.NAME_COLUMN).setValue(parentPath);
      
      // Get the parent label ID and add to hidden column
      const parentLabelId = getLabelId(parentPath);
      if (parentLabelId) {
        sheet.getRange(newRow, CONFIG.LABEL_ID_COLUMN).setValue(parentLabelId);
      }
      
      logDebug(`Added parent label "${parentPath}" to spreadsheet at row ${newRow}`);
      
      // Update the map to include the new parent
      sheetLabelMap[parentPath] = newRow;
      parentsAdded++;
    }
  }
}

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

/**
 * Logs debug messages if debug mode is enabled
 */
function logDebug(message) {
  if (CONFIG.DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`);
  }
}

/**
 * Logs error messages
 */
function logError(message) {
  console.error(`[ERROR] ${message}`);
}