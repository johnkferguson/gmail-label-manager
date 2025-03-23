/**
 * config.js - Gmail Label Manager
 *
 * Contains configuration settings and constants for the application.
 */

// Configuration
const CONFIG = {
  SHEET_NAME: 'Labels',      // The name of your sheet - MUST match your actual sheet name
  HEADER_ROW: 1,             // The row containing headers
  LABEL_ID_COLUMN: 1,        // Hidden column A for storing label IDs
  NAME_COLUMN: 2,            // Column B (1-indexed)
  DEBUG_MODE: true           // Set to false in production
};
