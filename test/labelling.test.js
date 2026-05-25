import { beforeEach, describe, expect, it } from "vitest";
import { loadScripts } from "./loader.js";
import { makeGmailAdvanced, makeGmailApp } from "./mocks/gmail.js";
import { makeSheet, makeSpreadsheet, makeSpreadsheetApp } from "./mocks/spreadsheet.js";

const baseConfig = () => ({
  SHEET_NAME: "Labels",
  HEADER_ROW: 1,
  LABEL_ID_COLUMN: 1,
  NAME_COLUMN: 2,
  DEBUG_MODE: false,
});

/**
 * @param {{
 *   initialLabels?: Array<{ name: string, id: string, type?: string }>,
 *   sheetData?: unknown[][],
 *   existingLabels?: Record<string, object>,
 *   threadsByLabel?: Record<string, object[]>,
 * }} [opts]
 */
function load(opts = {}) {
  const initialLabels = opts.initialLabels ?? [];
  const sheetData = opts.sheetData ?? [["Label ID", "Label Name"]];
  const existingLabels = opts.existingLabels ?? {};
  const threadsByLabel = opts.threadsByLabel ?? {};
  const sheet = makeSheet(sheetData);
  const spreadsheet = makeSpreadsheet({
    activeSheet: sheet,
    namedSheets: { Labels: sheet },
  });
  const spreadsheetApp = makeSpreadsheetApp(spreadsheet);
  const gmailAdvanced = makeGmailAdvanced({ labels: initialLabels });
  const gmailApp = makeGmailApp({ existingLabels, threadsByLabel });
  const ctx = loadScripts(
    ["src/config.js", "src/logging.js", "src/labelling.js"],
    {
      Gmail: gmailAdvanced,
      GmailApp: gmailApp,
      SpreadsheetApp: spreadsheetApp,
      CONFIG: baseConfig(),
    },
    ["CONFIG"]
  );
  return { ctx, sheet, spreadsheet, gmailAdvanced, gmailApp };
}

describe("parentLabelPaths", () => {
  const { ctx } = load();
  it("returns [] for a non-nested label", () => {
    expect(ctx.parentLabelPaths("Work")).toEqual([]);
  });
  it("returns one parent for a single nesting", () => {
    expect(ctx.parentLabelPaths("Work/Projects")).toEqual(["Work"]);
  });
  it("returns every prefix for deeper nesting", () => {
    expect(ctx.parentLabelPaths("Work/Projects/Foo/Bar")).toEqual([
      "Work",
      "Work/Projects",
      "Work/Projects/Foo",
    ]);
  });
  it("handles a trailing slash by returning the prefix and the empty-name parent", () => {
    expect(ctx.parentLabelPaths("Work/")).toEqual(["Work"]);
  });
});

describe("fetchLabelMap", () => {
  it("returns name→id for user labels", () => {
    const { ctx } = load({
      initialLabels: [
        { name: "Work", id: "L1", type: "user" },
        { name: "Personal", id: "L2", type: "user" },
      ],
    });
    expect(ctx.fetchLabelMap()).toEqual({ Work: "L1", Personal: "L2" });
  });

  it("filters out system labels", () => {
    const { ctx } = load({
      initialLabels: [
        { name: "Work", id: "L1", type: "user" },
        { name: "INBOX", id: "INBOX", type: "system" },
        { name: "CATEGORY_UPDATES", id: "CAT", type: "system" },
      ],
    });
    expect(ctx.fetchLabelMap()).toEqual({ Work: "L1" });
  });

  it("returns {} on API error", () => {
    const { ctx } = load();
    // Replace Gmail with a throwing one
    ctx.Gmail = {
      Users: {
        Labels: {
          list: () => {
            throw new Error("boom");
          },
        },
      },
    };
    expect(ctx.fetchLabelMap()).toEqual({});
  });
});

describe("createGmailLabel", () => {
  let ctx;
  let gmailAdvanced;
  beforeEach(() => {
    ({ ctx, gmailAdvanced } = load());
  });

  it("returns the existing id if labelMap already has the label", () => {
    const map = { Work: "L1" };
    const id = ctx.createGmailLabel("Work", map);
    expect(id).toBe("L1");
    // Should not have called create
    expect(gmailAdvanced._calls.some((c) => c.method === "create")).toBe(false);
  });

  it("calls Gmail.Users.Labels.create and returns the new id when missing", () => {
    const map = {};
    const id = ctx.createGmailLabel("NewLabel", map);
    expect(id).toMatch(/^L\d+$/);
    expect(map.NewLabel).toBe(id);
    expect(gmailAdvanced._calls.find((c) => c.method === "create")).toEqual({
      method: "create",
      name: "NewLabel",
    });
  });

  it("works without a labelMap (no mutation)", () => {
    const id = ctx.createGmailLabel("AnotherLabel");
    expect(id).toMatch(/^L\d+$/);
  });
});

describe("handleLabelChange", () => {
  it('dispatches "" → name to createLabel', () => {
    const { ctx, sheet } = load({
      sheetData: [
        ["Label ID", "Label Name"],
        ["", "NewLabel"],
      ],
    });
    ctx.handleLabelChange(sheet, 2, "", "NewLabel");
    // ID column should have been written (col 1 of row 2)
    const writes = sheet._calls.filter(
      (c) => c.method === "setValue" && c.row === 2 && c.col === 1
    );
    expect(writes.length).toBeGreaterThan(0);
  });

  it('dispatches name → "" to deleteLabel', () => {
    const { ctx, sheet, gmailApp } = load({
      sheetData: [
        ["Label ID", "Label Name"],
        ["L1", "Work"],
      ],
      existingLabels: {
        Work: { deleteLabel: () => {}, getName: () => "Work" },
      },
      initialLabels: [{ name: "Work", id: "L1", type: "user" }],
    });
    ctx.handleLabelChange(sheet, 2, "Work", "");
    expect(gmailApp._calls.some((c) => c.method === "search" && c.arg === "label:Work")).toBe(true);
  });

  it("dispatches name1 → name2 to updateLabel", () => {
    const { ctx, sheet, gmailApp } = load({
      sheetData: [
        ["Label ID", "Label Name"],
        ["L1", "OldName"],
      ],
      existingLabels: {
        OldName: {
          deleteLabel: () => {},
          addToThreads: () => {},
          removeFromThreads: () => {},
        },
        NewName: {
          deleteLabel: () => {},
          addToThreads: () => {},
          removeFromThreads: () => {},
        },
      },
      initialLabels: [{ name: "OldName", id: "L1", type: "user" }],
    });
    ctx.handleLabelChange(sheet, 2, "OldName", "NewName");
    expect(gmailApp._calls.some((c) => c.method === "search" && c.arg === "label:OldName")).toBe(
      true
    );
  });

  it("no-ops when both names are the same", () => {
    const { ctx, sheet } = load();
    ctx.handleLabelChange(sheet, 2, "Same", "Same");
    expect(sheet._calls.filter((c) => c.method === "setValue").length).toBe(0);
  });
});
