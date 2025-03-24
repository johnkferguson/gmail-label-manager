/**
 * triggers.js - Gmail Label Manager
 *
 * Contains functions related to installable triggers, event handling, and settings.
 */

/**
 * Creates the installable trigger for full permissions
 */
function createInstallableTrigger() {
  try {
    // Remove any existing triggers for the same function first
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onOpenWithFullPermissions') {
        ScriptApp.deleteTrigger(trigger);
      }
    }

    // Create a new installable trigger
    ScriptApp.newTrigger('onOpenWithFullPermissions')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();

    console.log("Installable trigger created successfully");
    return true;
  } catch (error) {
    console.error("Error creating installable trigger: " + error.message);
    throw error;
  }
}

/**
 * Toggles the edit trigger on/off
 */
function toggleEditTrigger() {
  // Check current status
  const triggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;

  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
      // Trigger exists, so delete it
      ScriptApp.deleteTrigger(trigger);
      triggerExists = true;
    }
  }

  // If trigger didn't exist, create it
  if (!triggerExists) {
    ScriptApp.newTrigger('onEditTrigger')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    SpreadsheetApp.getActive().toast('Row change trigger enabled');
  } else {
    SpreadsheetApp.getActive().toast('Row change trigger disabled');
  }

  // Refresh the menu
  onOpenWithFullPermissions();
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
 * Gets the auto sync setting status
 * @return {boolean} Whether auto sync is enabled
 */
function getAutoSyncEnabled() {
  return PropertiesService.getUserProperties().getProperty('autoSyncEnabled') === 'true';
}

/**
 * Toggles the auto sync setting on/off
 */
function toggleAutoSync() {
  const enabled = !getAutoSyncEnabled();
  PropertiesService.getUserProperties().setProperty('autoSyncEnabled', enabled.toString());
  SpreadsheetApp.getActive().toast(`Auto Sync ${enabled ? 'ENABLED' : 'DISABLED'}`);
  onOpenWithFullPermissions(); // Refresh menu
}
