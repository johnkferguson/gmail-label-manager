import { describe, expect, it } from "vitest";
import { loadScripts } from "./loader.js";
import { makeGmailAdvanced, makeGmailApp } from "./mocks/gmail.js";
import { makeLockService } from "./mocks/lock-service.js";
import { makePropertiesService } from "./mocks/properties.js";
import { makeSheet, makeSpreadsheet, makeSpreadsheetApp } from "./mocks/spreadsheet.js";

const baseConfig = () => ({
  SHEET_NAME: "Labels",
  HEADER_ROW: 1,
  LABEL_ID_COLUMN: 1,
  NAME_COLUMN: 2,
  DEBUG_MODE: false,
});

function makeScriptApp() {
  return {
    EventType: { ON_EDIT: "ON_EDIT", ON_OPEN: "ON_OPEN" },
    getProjectTriggers: () => [],
  };
}

/**
 * @param {{
 *   initialLabels?: Array<{ name: string, id: string, type?: string }>,
 *   sheetData?: unknown[][],
 *   contendedLock?: boolean,
 * }} [opts]
 */
function load(opts = {}) {
  const initialLabels = opts.initialLabels ?? [];
  const sheetData = opts.sheetData ?? [["Label ID", "Label Name"]];
  const contendedLock = opts.contendedLock ?? false;
  const sheet = makeSheet(sheetData);
  const spreadsheet = makeSpreadsheet({
    activeSheet: sheet,
    namedSheets: { Labels: sheet },
  });
  const spreadsheetApp = makeSpreadsheetApp(spreadsheet);
  const gmailAdvanced = makeGmailAdvanced({ labels: initialLabels });
  const gmailApp = makeGmailApp();
  const props = makePropertiesService();
  const lockService = makeLockService({ contended: contendedLock });
  const scriptApp = makeScriptApp();

  const ctx = loadScripts(
    ["src/config.js", "src/logging.js", "src/labelling.js", "src/triggers.js", "src/sync.js"],
    {
      Gmail: gmailAdvanced,
      GmailApp: gmailApp,
      SpreadsheetApp: spreadsheetApp,
      PropertiesService: props,
      LockService: lockService,
      ScriptApp: scriptApp,
      CONFIG: baseConfig(),
    },
    ["CONFIG"]
  );
  return { ctx, sheet, spreadsheet, gmailAdvanced, gmailApp, lockService };
}

describe("syncAllLabels — sheet → Gmail", () => {
  it("creates labels in Gmail that are in the sheet but not yet in Gmail", () => {
    const { ctx, gmailAdvanced } = load({
      sheetData: [
        ["Label ID", "Label Name"],
        ["", "NewLabel"],
      ],
    });
    const result = ctx.syncAllLabels();
    expect(result.createdInGmail).toEqual(["NewLabel"]);
    expect(
      gmailAdvanced._calls.find((c) => c.method === "create" && c.name === "NewLabel")
    ).toBeTruthy();
  });

  it("creates parents in Gmail for nested sheet labels", () => {
    const { ctx, gmailAdvanced } = load({
      sheetData: [
        ["Label ID", "Label Name"],
        ["", "Work/Projects/Foo"],
      ],
    });
    ctx.syncAllLabels();
    const createNames = gmailAdvanced._calls
      .filter((c) => c.method === "create")
      .map((c) => c.name);
    expect(createNames).toContain("Work");
    expect(createNames).toContain("Work/Projects");
    expect(createNames).toContain("Work/Projects/Foo");
  });

  it("backfills the hidden ID column for labels that already exist in both", () => {
    const { ctx, sheet } = load({
      initialLabels: [{ name: "ExistingLabel", id: "L42", type: "user" }],
      sheetData: [
        ["Label ID", "Label Name"],
        ["", "ExistingLabel"],
      ],
    });
    const result = ctx.syncAllLabels();
    expect(result.updatedIds).toEqual(["ExistingLabel"]);
    expect(sheet._data[1][0]).toBe("L42");
  });
});

describe("syncAllLabels — Gmail → sheet", () => {
  it("appends user labels present in Gmail but missing from the sheet", () => {
    const { ctx, sheet } = load({
      initialLabels: [
        { name: "OnlyInGmail", id: "L7", type: "user" },
        { name: "AlsoOnlyInGmail", id: "L8", type: "user" },
      ],
      sheetData: [["Label ID", "Label Name"]],
    });
    const result = ctx.syncAllLabels();
    expect(result.addedToSheet).toEqual(["OnlyInGmail", "AlsoOnlyInGmail"]);
    expect(sheet._data[1]).toEqual(["L7", "OnlyInGmail"]);
    expect(sheet._data[2]).toEqual(["L8", "AlsoOnlyInGmail"]);
  });

  it("does NOT append system labels", () => {
    const { ctx, sheet } = load({
      initialLabels: [
        { name: "INBOX", id: "INBOX", type: "system" },
        { name: "MyLabel", id: "L1", type: "user" },
      ],
    });
    ctx.syncAllLabels();
    expect(sheet._data.length).toBe(2); // header + 1 row
    expect(sheet._data[1]).toEqual(["L1", "MyLabel"]);
  });
});

describe("syncAllLabels — batching", () => {
  it("performs the row backfill in a single setValues call", () => {
    const { ctx, sheet } = load({
      initialLabels: [
        { name: "A", id: "L1", type: "user" },
        { name: "B", id: "L2", type: "user" },
        { name: "C", id: "L3", type: "user" },
      ],
      sheetData: [
        ["Label ID", "Label Name"],
        ["", "A"],
        ["", "B"],
        ["", "C"],
      ],
    });
    ctx.syncAllLabels();
    const idSetValues = sheet._calls.filter(
      (c) => c.method === "setValues" && c.col === 1 && c.w === 1
    );
    expect(idSetValues.length).toBe(1);
  });

  it("performs the appends in a single setValues call", () => {
    const { ctx, sheet } = load({
      initialLabels: [
        { name: "A", id: "L1", type: "user" },
        { name: "B", id: "L2", type: "user" },
        { name: "C", id: "L3", type: "user" },
      ],
    });
    ctx.syncAllLabels();
    const appendCalls = sheet._calls.filter((c) => c.method === "setValues" && c.w === 2);
    expect(appendCalls.length).toBe(1);
  });

  it("fetches the Gmail label list exactly once per run", () => {
    const { ctx, gmailAdvanced } = load({
      initialLabels: [{ name: "A", id: "L1", type: "user" }],
      sheetData: [
        ["Label ID", "Label Name"],
        ["", "NewLabel1"],
        ["", "NewLabel2"],
        ["", "NewLabel3"],
      ],
    });
    ctx.syncAllLabels();
    const listCalls = gmailAdvanced._calls.filter((c) => c.method === "list");
    expect(listCalls.length).toBe(1);
  });
});

describe("syncAllLabels — locking", () => {
  it("no-ops with a toast when the lock is contended", () => {
    const { ctx, spreadsheet } = load({ contendedLock: true });
    const result = ctx.syncAllLabels();
    expect(result).toEqual({ createdInGmail: [], addedToSheet: [], updatedIds: [] });
    expect(spreadsheet._toasts.some((t) => t.text.includes("Another sync"))).toBe(true);
  });
});
