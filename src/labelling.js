/**
 * labelling.js — Gmail Label Manager
 *
 * Label CRUD against Gmail, driven by sheet edits.
 *
 * Performance note: every public CRUD function accepts an optional pre-fetched
 * `labelMap` (name → id). Callers operating on multiple labels in a single
 * run (e.g. syncAllLabels) fetch the map once at the top and pass it through;
 * single-label callers (e.g. the edit trigger) can omit it and a fresh one is
 * fetched on demand. New labels created during a run are stitched into the
 * passed-in map so subsequent calls in the same run see them.
 *
 * Where possible we use the Gmail Advanced Service's `Gmail.Users.Labels.create`
 * directly — it returns the new label resource (with ID), eliminating the
 * "create then re-fetch the whole list to find the new ID" round-trip that
 * the earlier version did.
 */

/**
 * @typedef {Record<string, string>} LabelMap - Gmail label name → ID.
 */

/**
 * Fetch a fresh name→id map of user-managed Gmail labels. One Gmail API call.
 *
 * System labels (INBOX, SENT, CATEGORY_*, etc.) are filtered out — they're
 * never user-edited and never belong in the sheet.
 *
 * @returns {LabelMap}
 */
function fetchLabelMap() {
  try {
    // Gmail Advanced Service is enabled in the manifest; cast quiets tsc.
    const response = /** @type {any} */ (Gmail).Users.Labels.list("me");
    if (!response?.labels) {
      logError("No labels found in the Gmail account");
      return {};
    }
    /** @type {LabelMap} */
    const map = {};
    for (const label of response.labels) {
      if (label.type === "system") continue;
      if (label.name && label.id) map[label.name] = label.id;
    }
    logDebug(`Fetched ${Object.keys(map).length} user-managed Gmail labels`);
    return map;
  } catch (error) {
    logError(`Error getting label map: ${error.message}`);
    return {};
  }
}

/**
 * Backward-compat alias. Prefer fetchLabelMap + reuse for multi-label flows.
 * @returns {LabelMap}
 */
function getLabelMap() {
  return fetchLabelMap();
}

/**
 * Backward-compat single-label lookup. Each call does a full label-list fetch
 * — for batch operations, fetchLabelMap() once and index into the result.
 *
 * @param {string} labelName
 * @returns {string | null}
 */
function getLabelId(labelName) {
  return fetchLabelMap()[labelName] || null;
}

/**
 * Create a Gmail label via the Advanced Service and return its ID. If
 * `labelMap` is provided, the new entry is stitched in so subsequent
 * lookups in the same run see it.
 *
 * @param {string} labelName
 * @param {LabelMap} [labelMap]
 * @returns {string | null} The label ID, or null on failure.
 */
function createGmailLabel(labelName, labelMap) {
  try {
    if (labelMap?.[labelName]) return labelMap[labelName];
    const created = /** @type {any} */ (Gmail).Users.Labels.create({ name: labelName }, "me");
    if (created?.id) {
      if (labelMap) labelMap[labelName] = created.id;
      return created.id;
    }
    logError(`Gmail.Users.Labels.create returned no ID for "${labelName}"`);
    return null;
  } catch (error) {
    logError(`Error creating Gmail label "${labelName}": ${error.message}`);
    return null;
  }
}

/**
 * Classify an edit and dispatch to create/update/delete. The edit trigger
 * uses this; tests use it directly.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {string} oldLabelName
 * @param {string} newLabelName
 * @param {LabelMap} [labelMap]
 */
function handleLabelChange(sheet, row, oldLabelName, newLabelName, labelMap) {
  logDebug(`Processing label change: "${oldLabelName}" -> "${newLabelName}"`);

  if (oldLabelName === "" && newLabelName !== "") {
    createLabel(sheet, row, newLabelName, labelMap);
  } else if (oldLabelName !== "" && newLabelName === "") {
    deleteLabel(sheet, row, oldLabelName, labelMap);
  } else if (oldLabelName !== "" && newLabelName !== "" && oldLabelName !== newLabelName) {
    updateLabel(sheet, row, oldLabelName, newLabelName, labelMap);
  }
}

/**
 * Create a new Gmail label, creating any missing parents along the way.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {string} labelName
 * @param {LabelMap} [labelMap]
 */
function createLabel(sheet, row, labelName, labelMap) {
  logDebug(`Creating label: "${labelName}"`);
  const map = labelMap ?? fetchLabelMap();

  try {
    const existingId = map[labelName];
    if (existingId) {
      logDebug(`Label "${labelName}" already exists with ID: ${existingId}`);
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(existingId);
      SpreadsheetApp.getActive().toast(`Label "${labelName}" already exists in Gmail.`, "Info", 3);
      return;
    }

    addParentLabelsToSheet(sheet, labelName, map);

    const newLabelId = createGmailLabel(labelName, map);
    if (newLabelId) {
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(newLabelId);
      logDebug(`Label "${labelName}" created successfully with ID: ${newLabelId}`);
      SpreadsheetApp.getActive().toast(`New label "${labelName}" created in Gmail.`, "Success", 3);
    } else {
      logError(`Failed to retrieve ID for newly created label "${labelName}"`);
      SpreadsheetApp.getActive().toast(`Created label but couldn't retrieve its ID`, "Warning", 5);
    }
  } catch (error) {
    logError(`Error creating label "${labelName}": ${error.message}`);
    SpreadsheetApp.getActive().toast(
      `Error creating label "${labelName}": ${error.message}`,
      "Error",
      10
    );
  }
}

/**
 * Rename a Gmail label by creating the new one, retagging threads, removing
 * the old, and deleting it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {string} oldLabelName
 * @param {string} newLabelName
 * @param {LabelMap} [labelMap]
 */
function updateLabel(sheet, row, oldLabelName, newLabelName, labelMap) {
  logDebug(`Updating label: "${oldLabelName}" -> "${newLabelName}"`);
  const map = labelMap ?? fetchLabelMap();

  try {
    const oldLabelId = map[oldLabelName];
    if (!oldLabelId) {
      logDebug(
        `Old label "${oldLabelName}" not found, creating new label "${newLabelName}" instead`
      );
      createLabel(sheet, row, newLabelName, map);
      return;
    }

    addParentLabelsToSheet(sheet, newLabelName, map);

    const threads = GmailApp.search(`label:${oldLabelName}`);
    logDebug(`Found ${threads.length} threads with label "${oldLabelName}"`);

    const newLabelId = createGmailLabel(newLabelName, map);
    if (!newLabelId) {
      logError(`Failed to create new label "${newLabelName}"`);
      SpreadsheetApp.getActive().toast(
        `Error updating label: couldn't create new label`,
        "Error",
        10
      );
      return;
    }

    const oldLabel = GmailApp.getUserLabelByName(oldLabelName);
    const newLabel = GmailApp.getUserLabelByName(newLabelName);
    if (!oldLabel || !newLabel) {
      logError(`Failed to get label objects for "${oldLabelName}" or "${newLabelName}"`);
      return;
    }

    const BATCH_SIZE = 100;
    for (let i = 0; i < threads.length; i += BATCH_SIZE) {
      const batch = threads.slice(i, i + BATCH_SIZE);
      if (batch.length > 0) newLabel.addToThreads(batch);
    }
    for (let i = 0; i < threads.length; i += BATCH_SIZE) {
      const batch = threads.slice(i, i + BATCH_SIZE);
      if (batch.length > 0) oldLabel.removeFromThreads(batch);
    }

    oldLabel.deleteLabel();
    delete map[oldLabelName];

    sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).setValue(newLabelId);
    logDebug(`Label updated successfully from "${oldLabelName}" to "${newLabelName}"`);
    SpreadsheetApp.getActive().toast(
      `The label "${oldLabelName}" has been renamed to "${newLabelName}" within Gmail.`,
      "Success",
      5
    );
  } catch (error) {
    logError(`Error updating label from "${oldLabelName}" to "${newLabelName}": ${error.message}`);
    SpreadsheetApp.getActive().toast(`Error updating label: ${error.message}`, "Error", 10);
  }
}

/**
 * Delete a Gmail label — but only if no threads still reference it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {string} labelName
 * @param {LabelMap} [labelMap]
 */
function deleteLabel(sheet, row, labelName, labelMap) {
  logDebug(`Attempting to delete label: "${labelName}"`);
  const map = labelMap ?? fetchLabelMap();

  try {
    const labelId = map[labelName];
    if (!labelId) {
      logDebug(`Label "${labelName}" not found, nothing to delete`);
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).clearContent();
      return;
    }

    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      logDebug(`Label object for "${labelName}" not found`);
      sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).clearContent();
      return;
    }

    const threads = GmailApp.search(`label:${labelName}`);
    if (threads.length > 0) {
      const message = `Cannot delete label "${labelName}" as it still has ${threads.length} threads using it.`;
      logDebug(message);
      SpreadsheetApp.getActive().toast(message, "Warning", 10);
      sheet.getRange(row, CONFIG.NAME_COLUMN).setValue(labelName);
      return;
    }

    label.deleteLabel();
    delete map[labelName];
    sheet.getRange(row, CONFIG.LABEL_ID_COLUMN).clearContent();
    logDebug(`Label "${labelName}" deleted successfully`);
    SpreadsheetApp.getActive().toast(
      `The label "${labelName}" has been deleted from Gmail.`,
      "Info",
      5
    );
  } catch (error) {
    logError(`Error deleting label "${labelName}": ${error.message}`);
    SpreadsheetApp.getActive().toast(`Error deleting label: ${error.message}`, "Error", 10);
  }
}

/**
 * For a nested label like "Work/Projects/Foo", return the list of parent
 * path segments in order: ["Work", "Work/Projects"]. Empty array for
 * non-nested labels.
 *
 * Pure — used by both the edit path (which also appends parents to the
 * sheet) and the sync path (which doesn't, since sync's Step 2 picks them
 * up after they're created in Gmail).
 *
 * @param {string} labelName
 * @returns {string[]}
 */
function parentLabelPaths(labelName) {
  if (!labelName.includes("/")) return [];
  const parts = labelName.split("/");
  /** @type {string[]} */
  const paths = [];
  let parentPath = "";
  for (let i = 0; i < parts.length - 1; i++) {
    if (parentPath) parentPath += "/";
    parentPath += parts[i];
    paths.push(parentPath);
  }
  return paths;
}

/**
 * Ensure each parent path of a nested label exists in Gmail. Mutates
 * `labelMap` to include any newly created parents so subsequent lookups in
 * the same run see them.
 *
 * @param {string} labelName
 * @param {LabelMap} labelMap
 */
function ensureParentLabelsInGmail(labelName, labelMap) {
  for (const parentPath of parentLabelPaths(labelName)) {
    if (!labelMap[parentPath]) {
      const newId = createGmailLabel(parentPath, labelMap);
      if (newId) logDebug(`Created parent label "${parentPath}" in Gmail`);
    }
  }
}

/**
 * Same as ensureParentLabelsInGmail, but also appends any parent that's
 * missing from the sheet. Used by the edit-trigger path; the sync path
 * doesn't call this because its Step 2 picks up the newly-created parents
 * naturally.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} labelName
 * @param {LabelMap} [labelMap]
 */
function addParentLabelsToSheet(sheet, labelName, labelMap) {
  if (!labelName.includes("/")) return;
  const map = labelMap ?? fetchLabelMap();

  // Snapshot the sheet's existing label names so we know which parents need
  // to be appended.
  /** @type {Record<string, number>} */
  const sheetLabelMap = {};
  const lastRow = sheet.getLastRow();
  const firstDataRow = CONFIG.HEADER_ROW + 1;
  const dataRowCount = Math.max(0, lastRow - CONFIG.HEADER_ROW);
  const existingNames =
    dataRowCount > 0
      ? sheet.getRange(firstDataRow, CONFIG.NAME_COLUMN, dataRowCount, 1).getValues()
      : [];
  for (let i = 0; i < existingNames.length; i++) {
    const currentLabel = existingNames[i][0];
    if (currentLabel) sheetLabelMap[currentLabel] = firstDataRow + i;
  }

  // Stage append rows so we write them in one batch at the end.
  /** @type {[string, string][]} */
  const appendRows = [];

  for (const parentPath of parentLabelPaths(labelName)) {
    if (!map[parentPath]) {
      const newId = createGmailLabel(parentPath, map);
      if (newId) {
        logDebug(`Created parent label "${parentPath}" in Gmail`);
        SpreadsheetApp.getActive().toast(
          `New parent label "${parentPath}" created in Gmail.`,
          "Info",
          3
        );
      }
    }

    if (!sheetLabelMap[parentPath]) {
      const parentId = map[parentPath] ?? "";
      appendRows.push([parentId, parentPath]);
      logDebug(`Queueing parent label "${parentPath}" for append`);
      // Track placeholder row so a later parent in this same call doesn't
      // duplicate it.
      sheetLabelMap[parentPath] = lastRow + appendRows.length;
    }
  }

  if (appendRows.length > 0) {
    sheet.getRange(lastRow + 1, CONFIG.LABEL_ID_COLUMN, appendRows.length, 2).setValues(appendRows);
  }
}
