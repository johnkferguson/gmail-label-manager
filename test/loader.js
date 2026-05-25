/**
 * vm-based loader for the Apps Script source files. Same pattern as the
 * sibling gmail-conversation-labels / gmail-notes-from-sheet repos.
 *
 * Top-level `function` declarations land on the context automatically; top-
 * level `const`/`let` declarations are script-scoped, so pass their names in
 * `exposeNames` to expose them via a generated capture suffix.
 */

import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

/**
 * @param {string[]} relativePaths
 * @param {Record<string, unknown>} [mocks]
 * @param {string[]} [exposeNames]
 * @returns {vm.Context & Record<string, unknown>}
 */
export function loadScripts(relativePaths, mocks = {}, exposeNames = []) {
  const context = vm.createContext({
    console,
    // Share Node intrinsics so `instanceof Date` etc. work across realms.
    Date,
    Number,
    String,
    Object,
    Array,
    BigInt,
    ...mocks,
  });

  const parts = relativePaths.map((relPath) => readFileSync(resolve(repoRoot, relPath), "utf-8"));

  const captureSuffix =
    exposeNames.length > 0
      ? `\n${exposeNames
          .map((name) => `try { globalThis[${JSON.stringify(name)}] = ${name}; } catch (_) {}`)
          .join("\n")}`
      : "";

  const combined = `${parts.join("\n")}${captureSuffix}`;
  vm.runInContext(combined, context, { filename: relativePaths.join(",") });
  return context;
}
