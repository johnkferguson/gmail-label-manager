/**
 * Minimal mocks for SpreadsheetApp + Sheet + Range.
 *
 * Backed by a 2D array (rows × cols, 1-indexed via the GAS API). Captures
 * range read/write traffic so tests can assert on what hit the sheet.
 */

/**
 * @param {unknown[][]} initialData - 2D array of rows. Index 0 = row 1 (header).
 */
export function makeSheet(initialData = []) {
  /** @type {unknown[][]} */
  const data = initialData.map((row) => row.slice());
  /** @type {Array<{ method: string, row: number, col: number, h?: number, w?: number, value?: unknown, values?: unknown }>} */
  const calls = [];

  function ensure(row, col, h = 1, w = 1) {
    while (data.length < row + h - 1) data.push([]);
    for (let r = 0; r < data.length; r++) {
      while (data[r].length < col + w - 1) data[r].push("");
    }
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {number} [numRows]
   * @param {number} [numCols]
   */
  function getRange(row, col, numRows, numCols) {
    const h = numRows ?? 1;
    const w = numCols ?? 1;
    ensure(row, col, h, w);
    return {
      getValue: () => {
        calls.push({ method: "getValue", row, col });
        return data[row - 1]?.[col - 1] ?? "";
      },
      setValue: (value) => {
        calls.push({ method: "setValue", row, col, value });
        if (!data[row - 1]) data[row - 1] = [];
        data[row - 1][col - 1] = value;
      },
      clearContent: () => {
        calls.push({ method: "clearContent", row, col });
        if (data[row - 1]) data[row - 1][col - 1] = "";
      },
      getValues: () => {
        calls.push({ method: "getValues", row, col, h, w });
        const out = [];
        for (let r = 0; r < h; r++) {
          const rowOut = [];
          for (let c = 0; c < w; c++) {
            rowOut.push(data[row - 1 + r]?.[col - 1 + c] ?? "");
          }
          out.push(rowOut);
        }
        return out;
      },
      setValues: (values) => {
        calls.push({ method: "setValues", row, col, h, w, values });
        for (let r = 0; r < values.length; r++) {
          if (!data[row - 1 + r]) data[row - 1 + r] = [];
          for (let c = 0; c < values[r].length; c++) {
            data[row - 1 + r][col - 1 + c] = values[r][c];
          }
        }
      },
    };
  }

  return {
    getLastRow: () => data.length,
    getName: () => "Labels",
    getRange,
    hideColumns: () => {},
    _data: data,
    _calls: calls,
  };
}

/**
 * @param {{ activeSheet?: ReturnType<typeof makeSheet>, namedSheets?: Record<string, ReturnType<typeof makeSheet>> }} [opts]
 */
export function makeSpreadsheet(opts = {}) {
  const activeSheet = opts.activeSheet;
  const namedSheets = opts.namedSheets ?? {};
  /** @type {Array<{ method: string, text: string, title?: string, timeout?: number }>} */
  const toasts = [];
  const obj = {
    getActiveSheet: () => activeSheet ?? null,
    getSheetByName: (name) => namedSheets[name] ?? activeSheet ?? null,
    getId: () => "spreadsheet-id",
    toast: (text, title, timeout) => {
      toasts.push({ method: "toast", text, title, timeout });
    },
    _toasts: toasts,
  };
  return obj;
}

/**
 * @param {ReturnType<typeof makeSpreadsheet>} spreadsheet
 */
export function makeSpreadsheetApp(spreadsheet) {
  return {
    getActiveSpreadsheet: () => spreadsheet,
    getActive: () => spreadsheet,
  };
}
