# Gmail Label Manager

[![CI](https://github.com/johnkferguson/gmail-label-manager/actions/workflows/ci.yml/badge.svg)](https://github.com/johnkferguson/gmail-label-manager/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A Google Apps Script add-on that lets you manage your Gmail labels from a Google Sheet. Edit a label name in the sheet → Gmail renames it. Add a row → Gmail creates it. Two-way sync keeps both sides aligned.

## Features

- **Live edits**: rename / add / remove a Gmail label by editing the sheet.
- **Bidirectional merge sync**: one-shot reconciliation between sheet ↔ Gmail.
- **Auto-sync on open**: optional — runs the merge when you open the spreadsheet.
- **Nested labels**: parent labels are auto-created in both Gmail and the sheet.
- **Sheet menu integration**: enable/disable the live trigger and run sync from a "Gmail Labels" menu.

## How it works

The script is **container-bound** to a single Google Sheet (you create it via `clasp create --type sheets`, which binds it to a sheet you own). Two trigger modes:

- **`onOpen` (simple trigger, limited permissions)** runs unauthenticated when the sheet opens. It can only show a "Click to Enable" menu — Apps Script restricts what unauthorized simple triggers can do.
- **`onOpenWithFullPermissions` (installable trigger)** is created after the user clicks Enable. From then on, the full menu (toggles + sync button) appears on every open, and (if auto-sync is on) a sync runs automatically.

The live propagation uses a separate installable **edit trigger** (`onEditTrigger`). When you change a label name in column B, it diffs old vs new and dispatches to create / rename / delete in Gmail.

## Privacy & security

- **No third-party servers.** The code runs as your own Apps Script project, bound to your own sheet, under your own Google account. Your labels and your script — there's nowhere else for it to go.
- **You install it yourself** via `clasp push`. Read every line first; it's a single repo.
- **No outbound network calls.** Only first-party Google APIs (Gmail Advanced Service, GmailApp, SpreadsheetApp, PropertiesService, ScriptApp, LockService).
- **Minimal scopes.** The OAuth consent screen will ask for exactly these:
  | Scope | Used for |
  |---|---|
  | `gmail.modify` | Read your Gmail labels and create / rename / delete them |
  | `spreadsheets.currentonly` | Read / write **only** the spreadsheet this script is bound to |
  | `script.scriptapp` | Register the installable triggers (onOpen + onEdit) |

  The `spreadsheets.currentonly` scope is the narrowest available — the script can only touch the one sheet it's bound to, not your other spreadsheets.

To revoke: <https://myaccount.google.com/permissions> → find the script project → Remove access.

## Setup

You'll need [Bun](https://bun.sh), a Google account, and the Apps Script API enabled at <https://script.google.com/home/usersettings>.

```bash
bun install
bunx clasp login
bunx clasp create --type sheets --title "Gmail Label Manager"
cp .clasp-example.json .clasp.json    # then paste in the scriptId from the create output
bun run push
```

The `--type sheets` flag creates a new Google Sheet AND a bound script — both newly created, owned by you.

In the new sheet:

1. Rename the first sheet (tab) to **`Labels`** (matches `CONFIG.SHEET_NAME`).
2. Refresh the page so the simple trigger fires and the menu appears.
3. **Gmail Labels → Click to Enable**. Grant the requested scopes.
4. Refresh the page once more — the full menu appears.
5. **Sync All Labels** to pull your existing Gmail labels into the sheet.

## Menu options

After enabling, the **Gmail Labels** menu has:

| Option | Behavior |
|---|---|
| `[ON]/[OFF] Trigger On Spreadsheet Change` | Toggle live edit-trigger. When ON, editing column B in the Labels sheet immediately creates / renames / deletes the corresponding Gmail label. |
| `[ON]/[OFF] Auto Sync On Startup` | Toggle the bidirectional merge that runs every time the sheet opens. |
| `Sync All Labels` | Run the merge sync once. |

## Sheet format

| Column | Purpose |
|---|---|
| A — **`Label ID`** | Hidden, managed by the script. Stores the Gmail label ID for each row. Don't edit. |
| B — **`Label Name`** | The label name. This is the only column you edit. |

Use a slash (`/`) for nested labels — e.g. `Work/Projects/Foo`. Missing parent labels (`Work`, `Work/Projects`) are auto-created in Gmail and appended to the sheet.

## Sync semantics

The merge sync (Auto Sync On Startup or Sync All Labels) reconciles both sides without deletion:

1. Sheet labels missing from Gmail → **created** in Gmail. Nested parents created too.
2. Gmail labels missing from the sheet → **appended** to the sheet (rows added at the bottom).
3. Labels present in both → the hidden ID column is refreshed to the current Gmail ID.

No label is ever deleted by the merge — to delete one, **clear the name cell** in column B with the edit trigger enabled. If the label still has emails attached, the script refuses to delete it and restores the cell.

## Configuration

`src/config.js` is split into:

- **`USER_CONFIG`** — user-facing settings you might change per installation:
  | Field | Default | Description |
  |---|---|---|
  | `SHEET_NAME` | `"Labels"` | Name of the sheet tab to sync. Must match the actual tab name. |
  | `DEBUG_MODE` | `false` | Verbose `[DEBUG]` lines to View → Logs. Flip to `true` while debugging. |
- **`APP_CONFIG`** — internal sheet layout (`HEADER_ROW`, `LABEL_ID_COLUMN`, `NAME_COLUMN`). Don't change after you've populated a sheet.

## Development

```bash
bun install
bun run lint        # biome check
bun run lint:fix    # biome check --write
bun run format
bun run test        # vitest run
bun run typecheck   # tsc --noEmit (JSDoc-based, no compile)
bun run push        # clasp push
bun run pull        # clasp pull
```

The source under `src/` is Apps Script JS — top-level `function` and `const` declarations are global at runtime, so no `import` / `export`. Tests load each source file into a Node `vm` context with mocked GAS globals (`Gmail`, `GmailApp`, `SpreadsheetApp`, etc.) via `test/loader.js`.

### Project layout

```text
src/
  appsscript.json   — manifest (scopes, runtime, Gmail Advanced Service)
  config.js         — USER_CONFIG + APP_CONFIG → CONFIG
  logging.js        — logDebug / logError
  labelling.js      — fetchLabelMap, createGmailLabel, parentLabelPaths,
                      handleLabelChange, createLabel, updateLabel,
                      deleteLabel, ensureParentLabelsInGmail,
                      addParentLabelsToSheet
  sync.js           — syncAllLabels (one-shot bidirectional merge)
  triggers.js       — onEditTrigger, install/toggle for the edit trigger,
                      auto-sync state
  ui.js             — onOpen, onOpenWithFullPermissions, menu, setup,
                      showSyncResults

test/
  loader.js         — vm-based source loader
  mocks/            — Gmail, SpreadsheetApp, PropertiesService, LockService
  *.test.js         — vitest suites
```

## Troubleshooting

- **Permission errors / "Click to Enable" never updates** — re-authorize from the menu. The simple `onOpen` trigger has no scope; the installable `onOpenWithFullPermissions` is what does the work.
- **`Sheet "Labels" not found`** — your tab name doesn't match `CONFIG.SHEET_NAME`. Either rename the tab or update the config.
- **Label rename leaves both labels in Gmail** — the rename runs add-new → remove-old → delete-old. If the script hits Apps Script's 6-minute runtime cap mid-rename, you can end up with both labels. Re-run "Sync All Labels" to reconcile.
- **Can't delete a label** — the script refuses to delete a Gmail label that still has emails on it. Move/archive the emails first, then clear the cell again.
- **Verbose logs missing** — set `USER_CONFIG.DEBUG_MODE = true` in `src/config.js` and `bun run push`.

## License

[MIT](LICENSE).
