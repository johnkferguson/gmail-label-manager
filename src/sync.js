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

  // Document-scoped lock — prevents auto-sync (on open) and a manual "Sync
  // All Labels" click from running concurrently, which could race on row
  // appends in Step 2. tryLock(0) means contended calls no-op; the user can
  // re-run from the menu once the other sync finishes.
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(0)) {
    logDebug("syncAllLabels skipped — another sync is already in flight");
    SpreadsheetApp.getActive().toast("Another sync is in progress — try again shortly.", "Info", 5);
    return { createdInGmail: [], addedToSheet: [], updatedIds: [] };
  }

  try {
    return syncAllLabelsImpl(sheet);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Inner implementation — the locking + early-exit shell above wraps this.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {SyncResults}
 */
function syncAllLabelsImpl(sheet) {
  let lastRow = sheet.getLastRow();

  /** @type {SyncResults} */
  const results = { createdInGmail: [], addedToSheet: [], updatedIds: [] };

  // One Gmail API round-trip; mutated locally as we create.
  const labelMap = fetchLabelMap();

  // Snapshot of the labels currently in the sheet (name → row). Single batch
  // read of the name column.
  /** @type {Record<string, number>} */
  const sheetLabelMap = {};
  const firstDataRow = CONFIG.HEADER_ROW + 1;
  const dataRowCount = Math.max(0, lastRow - CONFIG.HEADER_ROW);
  const existingNames =
    dataRowCount > 0
      ? sheet.getRange(firstDataRow, CONFIG.NAME_COLUMN, dataRowCount, 1).getValues()
      : [];
  for (let i = 0; i < existingNames.length; i++) {
    const labelName = existingNames[i][0];
    if (labelName) sheetLabelMap[labelName] = firstDataRow + i;
  }

  // Step 1: sheet → Gmail. Stage the desired ID for each existing row in a
  // local array, then write the whole ID column back in a single batch at
  // the end. Avoids one setValue API call per row.
  /** @type {(string | "")[]} */
  const idColumn = dataRowCount > 0 ? new Array(dataRowCount).fill("") : [];
  let anyStep1Writes = false;

  for (const labelName of Object.keys(sheetLabelMap)) {
    const row = sheetLabelMap[labelName];
    const idIndex = row - firstDataRow;

    try {
      const labelId = labelMap[labelName];

      if (!labelId) {
        // Create parents first if this is a nested label. Sync's Step 2 picks
        // up the parent rows we'd otherwise need to append explicitly.
        ensureParentLabelsInGmail(labelName, labelMap);
        const newLabelId = createGmailLabel(labelName, labelMap);
        if (newLabelId) {
          idColumn[idIndex] = newLabelId;
          anyStep1Writes = true;
          logDebug(`Created label "${labelName}" during sync with ID: ${newLabelId}`);
          results.createdInGmail.push(labelName);
        }
      } else {
        idColumn[idIndex] = labelId;
        anyStep1Writes = true;
        logDebug(`Updated ID for existing label "${labelName}": ${labelId}`);
        results.updatedIds.push(labelName);
      }
    } catch (error) {
      logError(`Error syncing label "${labelName}": ${error.message}`);
    }
  }

  if (anyStep1Writes) {
    sheet
      .getRange(firstDataRow, CONFIG.LABEL_ID_COLUMN, idColumn.length, 1)
      .setValues(idColumn.map((id) => [id]));
  }

  // Step 2: Gmail → sheet. Append any user-managed Gmail label not already
  // in the sheet, in a single batch write at the end. labelMap was filtered
  // to exclude system labels at fetch time, so no further filtering needed.
  /** @type {[string, string][]} */
  const appendRows = [];
  for (const [name, id] of Object.entries(labelMap)) {
    if (sheetLabelMap[name]) continue;
    appendRows.push([id, name]);
    results.addedToSheet.push(name);
    logDebug(`Queued Gmail label "${name}" for append`);
  }
  if (appendRows.length > 0) {
    sheet.getRange(lastRow + 1, CONFIG.LABEL_ID_COLUMN, appendRows.length, 2).setValues(appendRows);
    lastRow += appendRows.length;
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
