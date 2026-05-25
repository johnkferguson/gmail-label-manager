function makeStore() {
  /** @type {Map<string, string>} */
  const store = new Map();
  return {
    getProperty: (key) => (store.has(key) ? store.get(key) : null),
    setProperty: (key, value) => {
      store.set(key, String(value));
    },
    getProperties: () => Object.fromEntries(store),
    deleteProperty: (key) => {
      store.delete(key);
    },
    deleteAllProperties: () => {
      store.clear();
    },
    _store: store,
  };
}

export function makePropertiesService() {
  const script = makeStore();
  const user = makeStore();
  return {
    getScriptProperties: () => script,
    getUserProperties: () => user,
    _script: script,
    _user: user,
  };
}
