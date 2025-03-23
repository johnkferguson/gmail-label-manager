/**
 * logging.js - Gmail Label Manager
 *
 * Contains utility functions for logging and debugging.
 */

/**
 * Logs debug messages if debug mode is enabled
 */
function logDebug(message) {
  if (CONFIG.DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`);
  }
}

/**
 * Logs error messages
 */
function logError(message) {
  console.error(`[ERROR] ${message}`);
}
