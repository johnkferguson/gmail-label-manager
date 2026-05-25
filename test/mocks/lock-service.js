/**
 * Mock LockService. Default behavior: tryLock always succeeds.
 * Tests that want contention pass { contended: true }.
 */

/** @param {{ contended?: boolean }} [opts] */
export function makeLockService(opts = {}) {
  const contended = opts.contended ?? false;
  /** @type {Array<{ method: string }>} */
  const calls = [];
  const lock = {
    tryLock: (_ms) => {
      calls.push({ method: "tryLock" });
      return !contended;
    },
    releaseLock: () => {
      calls.push({ method: "releaseLock" });
    },
  };
  return {
    getDocumentLock: () => lock,
    getUserLock: () => lock,
    _calls: calls,
  };
}
