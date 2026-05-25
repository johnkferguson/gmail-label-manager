/**
 * triggers.js — Gmail Label Manager
 *
 * Installable trigger management and event handlers.
 *
 * Trigger handler names are referenced as strings by ScriptApp.newTrigger and
 * matched against trigger.getHandlerFunction(); centralizing them as
 * constants prevents a rename from silently orphaning a registered trigger.
 */

const HANDLER_OPEN_FULL = "onOpenWithFullPermissions";
const HANDLER_ON_EDIT = "onEditTrigger";

/**
 * Installs the onOpen handler that runs with full permissions. Idempotent —
 * removes any prior trigger with the same handler before creating.
 */
function createInstallableTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === HANDLER_OPEN_FULL) {
        ScriptApp.deleteTrigger(trigger);
      }
    }

    ScriptApp.newTrigger(HANDLER_OPEN_FULL)
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();

    console.log("Installable trigger created successfully");
    return true;
  } catch (error) {
    console.error(`Error creating installable trigger: ${error.message}`);
    throw error;
  }
}

/**
 * Toggle the on-edit trigger on/off.
 *
 * Important: match BOTH event type AND handler name. The earlier version
 * matched only event type and would clobber any other ON_EDIT trigger that
 * happened to exist on the spreadsheet (e.g. a future feature or another
 * script bound to the same sheet).
 */
function toggleEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let triggerExisted = false;

  for (const trigger of triggers) {
    if (
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
      trigger.getHandlerFunction() === HANDLER_ON_EDIT
    ) {
      ScriptApp.deleteTrigger(trigger);
      triggerExisted = true;
    }
  }

  if (!triggerExisted) {
    ScriptApp.newTrigger(HANDLER_ON_EDIT)
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    SpreadsheetApp.getActive().toast("Row change trigger enabled");
  } else {
    SpreadsheetApp.getActive().toast("Row change trigger disabled");
  }

  // Refresh the menu so the [ON]/[OFF] prefix reflects the new state.
  onOpenWithFullPermissions();
}

/**
 * Edit event handler. Filters to our target sheet + name column only.
 *
 * Wrapped in a document-scoped lock: two near-simultaneous edits (paste,
 * autofill, undo+redo) can fire onEdit in parallel. Without serialization
 * they race on getLabelId / setValue and can double-create or write a stale
 * ID back. tryLock(0) means a contended edit no-ops immediately; the user's
 * change is still in the sheet, so they can re-trigger the sync via the
 * "Sync All Labels" menu if needed.
 */
function onEditTrigger(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  const column = e.range.getColumn();
  if (row <= CONFIG.HEADER_ROW) return;
  if (column !== CONFIG.NAME_COLUMN) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(0)) {
    logDebug("onEditTrigger skipped — another edit is in flight");
    return;
  }

  try {
    logDebug(`Change detected in Name column at row ${row}`);
    const newLabelName = e.value || "";
    const oldLabelName = e.oldValue || "";

    handleLabelChange(sheet, row, oldLabelName, newLabelName);
  } catch (error) {
    logError(`Error in onEditTrigger: ${error.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * @returns {boolean}
 */
function getAutoSyncEnabled() {
  return PropertiesService.getUserProperties().getProperty("autoSyncEnabled") === "true";
}

function toggleAutoSync() {
  const enabled = !getAutoSyncEnabled();
  PropertiesService.getUserProperties().setProperty("autoSyncEnabled", enabled.toString());
  SpreadsheetApp.getActive().toast(`Auto Sync ${enabled ? "ENABLED" : "DISABLED"}`);
  onOpenWithFullPermissions();
}
