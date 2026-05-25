/**
 * sync.js — Gmail Label Manager
 *
 * Bidirectional sync between Gmail labels and the configured sheet.
 *
 * Pulls the Gmail label list exactly once at the top and threads the
 * resulting map through every helper that needs to look up or mutate a label.
 * createGmailLabel updates the map in place, so newly-created labels are
 * visible to subsequent steps without a second list fetch.
 */

/**
 * @typedef {object} SyncResults
 * @property {string[]} createdInGmail
 * @property {string[]} addedToSheet
 * @property {string[]} updatedIds
 */

/**
 * Run a one-shot bidirectional merge between Gmail and the sheet.
 *
 * @returns {SyncResults}
 */
function syncAllLabels() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    const message = `Sheet "${CONFIG.SHEET_NAME}" not found — check CONFIG.SHEET_NAME.`;
    logError(message);
    SpreadsheetApp.getActive().toast(message, "Error", 10);
    return { createdInGmail: [], addedToSheet: [], updatedIds: [] };
  }
  let lastRow = sheet.getLastRow();

  /** @type {SyncResults} */
  const results = { createdInGmail: [], addedToSheet: [], updatedIds: [] };

  // One Gmail API round-trip; mutated locally as we create.
  const labelMap = fetchLabelMap();

  // Snapshot of the labels currently in the sheet (name → row).
  /** @type {Record<string, number>} */
  const sheetLabelMap = {};
  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const labelName = sheet.getRange(row, CONFIG.NAME_COLUMN).getValue();
    if (labelName) sheetLabelMap[labelName] = row;
  }

  // Step 1: sheet → Gmail. Create any sheet labels missing from Gmail; for
  // labels that exist in both, refresh the cached ID in the hidden column.
  for (const labelName of Object.keys(sheetLabelMap)) {
    const row = sheetLabelMap[labelName];

    try {
      const labelId = labelMap[labelName];

      if (!labelId) {
        // Create parents first if this is a nested label.
        if (labelName.includes("/")) {
          const parts = labelName.split("/");
          let parentPath = "";
          for (let i = 0; i < parts.length - 1; i++) {
            if (parentPath) parentPath += "/";
            parentPath += parts[i];
            if (!labelMap[parentPath]) {
              createGmailLabel(parentPath, labelMap);
              logDebug(`Created parent label "${parentPath}"`);
            }
          }
        }

        const newLabelId = createGmailLabel(labelName, labelMap);
        if (newLabelId) {
          sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(newLabelId);
          logDebug(`Created label "${labelName}" during sync with ID: ${newLabelId}`);
          results.createdInGmail.push(labelName);
        }
      } else {
        sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(labelId);
        logDebug(`Updated ID for existing label "${labelName}": ${labelId}`);
        results.updatedIds.push(labelName);
      }
    } catch (error) {
      logError(`Error syncing label "${labelName}": ${error.message}`);
    }
  }

  // Step 2: Gmail → sheet. Append any user-managed Gmail label not already
  // in the sheet. labelMap was filtered to exclude system labels at fetch
  // time, so no further filtering needed here.
  for (const [name, id] of Object.entries(labelMap)) {
    if (sheetLabelMap[name]) continue;
    lastRow++;
    sheet.getRange(lastRow, CONFIG.LABEL_ID_COLUMN).setValue(id);
    sheet.getRange(lastRow, CONFIG.NAME_COLUMN).setValue(name);
    logDebug(`Added Gmail label "${name}" to spreadsheet`);
    results.addedToSheet.push(name);
  }

  // Suppress the generic toast when auto-sync is the caller — auto-sync
  // renders its own detailed summary via showSyncResults.
  if (
    !getAutoSyncEnabled() ||
    !ScriptApp.getProjectTriggers().some(
      (trigger) =>
        trigger.getEventType() === ScriptApp.EventType.ON_OPEN &&
        trigger.getHandlerFunction() === "onOpenWithFullPermissions"
    )
  ) {
    SpreadsheetApp.getActive().toast("All labels synced between Spreadsheet and Gmail.");
  }

  return results;
}
