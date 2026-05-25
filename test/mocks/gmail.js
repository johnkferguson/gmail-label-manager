/**
 * Mocks for GmailApp and the Gmail Advanced Service used by labelling.js.
 */

/**
 * @param {{ labels?: Array<{ name: string, id: string, type?: string }> }} [opts]
 */
export function makeGmailAdvanced(opts = {}) {
  /** @type {Array<{ name: string, id: string, type?: string }>} */
  const labels = (opts.labels ?? []).slice();
  let nextId = 1000;
  /** @type {Array<{ method: string, name?: string }>} */
  const calls = [];
  return {
    Users: {
      Labels: {
        list: (_user) => {
          calls.push({ method: "list" });
          return { labels: labels.slice() };
        },
        create: (resource, _user) => {
          calls.push({ method: "create", name: resource.name });
          const created = { name: resource.name, id: `L${nextId++}`, type: "user" };
          labels.push(created);
          return created;
        },
      },
    },
    _labels: labels,
    _calls: calls,
  };
}

/**
 * @param {{ threadsByLabel?: Record<string, object[]>, existingLabels?: Record<string, object> }} [opts]
 */
export function makeGmailApp(opts = {}) {
  const threadsByLabel = opts.threadsByLabel ?? {};
  const existingLabels = opts.existingLabels ?? {};
  /** @type {Array<{ method: string, arg?: string }>} */
  const calls = [];
  return {
    search: (query) => {
      calls.push({ method: "search", arg: query });
      const match = query.match(/^label:(.+)$/);
      const name = match ? match[1] : "";
      return threadsByLabel[name] ?? [];
    },
    getUserLabelByName: (name) => {
      calls.push({ method: "getUserLabelByName", arg: name });
      return existingLabels[name] ?? null;
    },
    createLabel: (name) => {
      calls.push({ method: "createLabel", arg: name });
      return { getName: () => name };
    },
    _calls: calls,
  };
}
