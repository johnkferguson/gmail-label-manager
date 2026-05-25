/**
 * config.js — Gmail Label Manager
 *
 * Configuration is split so users can tweak the user-facing settings without
 * touching app structure:
 *
 *   - USER_CONFIG: things you might reasonably change per installation
 *     (sheet name, debug toggle).
 *   - APP_CONFIG : sheet layout. Changing these breaks compatibility with
 *     existing populated sheets, so they're separated from the user-facing
 *     settings.
 *
 * Both merge into a single global CONFIG that the rest of the code reads.
 */

// User-facing settings. Edit these to taste.
const USER_CONFIG = {
  // Name of the sheet (tab) that mirrors your Gmail labels.
  SHEET_NAME: "Labels",

  // Verbose logging to View → Logs. Off by default for production installs;
  // flip to true while debugging.
  DEBUG_MODE: false,
};

// App-internal sheet layout. Only change if you know what you're doing.
const APP_CONFIG = {
  // Row holding the column headers.
  HEADER_ROW: 1,

  // Hidden column A — stores the Gmail label ID for each row. Managed by the
  // script; don't edit by hand.
  LABEL_ID_COLUMN: 1,

  // Column B — the label name. This is the only column users edit directly.
  NAME_COLUMN: 2,
};

const CONFIG = {
  ...USER_CONFIG,
  ...APP_CONFIG,
};
