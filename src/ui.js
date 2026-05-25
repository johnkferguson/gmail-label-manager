/**
 * ui.js - Gmail Label Manager
 *
 * Contains menu creation and UI-related functions for the application.
 */

/**
 * SIMPLE TRIGGER (limited permissions)
 * This runs automatically when the spreadsheet is opened,
 * but with limited authorization capabilities
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Set up the label ID column - this works with limited permissions
    setupLabelIdColumn();

    // Create a simple menu that shows options available with limited permissions
    ui.createMenu("Gmail Labels").addItem("Click to Enable", "promptForPermissions").addToUi();
  } catch (error) {
    console.error(`Error in onOpen: ${error.message}`);
  }
}

/**
 * INSTALLABLE TRIGGER (full permissions)
 * This runs only after the user has authorized the script
 * and has full authorization capabilities
 */
function onOpenWithFullPermissions() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Check if edit trigger exists - requires full permissions
    let editTriggerEnabled = false;
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (
        trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
        trigger.getHandlerFunction() === HANDLER_ON_EDIT
      ) {
        editTriggerEnabled = true;
        break;
      }
    }

    // Create a full menu with all options
    const menu = ui
      .createMenu("Gmail Labels")
      .addItem(
        `${editTriggerEnabled ? "[ON] " : "[OFF] "}Trigger On Spreadsheet Change`,
        "toggleEditTrigger"
      )
      .addItem(`${getAutoSyncEnabled() ? "[ON] " : "[OFF] "}Auto Sync On Startup`, "toggleAutoSync")
      .addItem("Sync All Labels", "syncAllLabels");

    menu.addToUi();

    // Run auto sync if enabled
    if (getAutoSyncEnabled()) {
      try {
        const results = syncAllLabels();
        showSyncResults(results);
      } catch (error) {
        logError(`Auto sync failed: ${error.message}`);
        SpreadsheetApp.getActive().toast("Auto sync failed - check logs", "Error", 6);
      }
    }

    logDebug("Full permissions menu created successfully");
  } catch (error) {
    console.error(`Error in onOpenWithFullPermissions: ${error.message}`);
  }
}

/**
 * Prompts the user to enable advanced features and creates the installable trigger
 */
function promptForPermissions() {
  const ui = SpreadsheetApp.getUi();

  // Show a dialog with instructions
  const response = ui.alert(
    "Additional Permissions Required",
    "To use all features of this script, you need to grant additional permissions.\n\n" +
      "Would you like to enable these permissions now?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    try {
      // Create an installable trigger that will run with full permissions
      createInstallableTrigger();

      // Show success message
      ui.alert(
        "Success",
        "Permissions have been requested. Please refresh the spreadsheet to see all options.",
        ui.ButtonSet.OK
      );

      // Force run the full permissions function if we got here
      try {
        onOpenWithFullPermissions();
      } catch (e) {
        // This might fail if permissions aren't fully granted yet, which is expected
        console.log(`Could not run full permissions function yet: ${e.message}`);
      }
    } catch (e) {
      // If direct enabling fails, show error
      ui.alert(
        "Error",
        "Unable to enable permissions: " +
          e.message +
          "\n\n" +
          "Please try refreshing the page and trying again.",
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Sets up the hidden column for label IDs
 */
function setupLabelIdColumn() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      // Sheet doesn't exist yet, don't proceed
      return;
    }

    // Create header for ID column if needed
    sheet.getRange(CONFIG.HEADER_ROW, CONFIG.LABEL_ID_COLUMN).setValue("Label ID");

    // Try to hide the column
    try {
      sheet.hideColumns(CONFIG.LABEL_ID_COLUMN);
      logDebug("Label ID column hidden successfully");
    } catch (error) {
      logError(`Failed to hide Label ID column: ${error.message}`);
    }
  } catch (error) {
    // Skip setting up the column if there's an error
    logError(`Error in setupLabelIdColumn: ${error.message}`);
  }
}

/**
 * Render a single multi-line toast summarizing a sync run.
 *
 * Previously this fired up to three sequential toasts with Utilities.sleep
 * between them, hoping the sleeps would serialize the UI. They didn't —
 * Sheets only shows one toast at a time and the second one just replaced
 * the first. Net effect: a 1-second wall-clock delay and the user only saw
 * the last message.
 *
 * One toast, formatted, no sleeps.
 *
 * @param {{ createdInGmail?: string[], addedToSheet?: string[] }} results
 */
function showSyncResults(results) {
  const ss = SpreadsheetApp.getActive();
  const created = results.createdInGmail ?? [];
  const added = results.addedToSheet ?? [];
  const total = created.length + added.length;

  if (total === 0) {
    ss.toast("Auto Sync complete: No changes needed", "Auto Sync Complete", 3);
    return;
  }

  const lines = [`${total} label(s) synchronized.`];
  if (created.length > 0) {
    lines.push(`Created in Gmail: ${created.map((n) => `"${n}"`).join(", ")}`);
  }
  if (added.length > 0) {
    lines.push(`Added to sheet: ${added.map((n) => `"${n}"`).join(", ")}`);
  }
  ss.toast(lines.join("\n"), "Auto Sync Complete", 8);
}
