// A transaction/request double: native browser cloning, permissions and disk durability remain browser checks.
export function testIndexedDb() {
  const databases = new Map<string, Map<string, Map<IDBValidKey, unknown>>>();
  const failure = { open: false, read: false, write: false, clear: false };
  const calls: Array<{ database: string; store: string; action: string }> = [];
  const factory = {
    open(name: string) {
      const request = {} as IDBOpenDBRequest;
      queueMicrotask(() => {
        if (failure.open) { Object.assign(request, { error: new DOMException("Storage blocked", "SecurityError") }); request.onerror?.(new Event("error")); return; }
        const fresh = !databases.has(name);
        if (fresh) databases.set(name, new Map());
        const stores = databases.get(name)!;
        const database = {
          createObjectStore(store: string) { stores.set(store, new Map()); },
          close() {},
          transaction(store: string, _mode: IDBTransactionMode) {
            const transaction = { error: null } as IDBTransaction;
            const operation = (action: "get" | "put" | "clear", value?: unknown, key?: IDBValidKey) => {
              const result = {} as IDBRequest;
              calls.push({ database: name, store, action });
              queueMicrotask(() => {
                if (failure[action === "get" ? "read" : action === "put" ? "write" : "clear"]) {
                  Object.assign(transaction, { error: new DOMException("Storage failed", "DataCloneError") });
                  transaction.onabort?.(new Event("abort")); return;
                }
                const records = stores.get(store)!;
                if (!records) throw new Error(`Unknown store ${store}`);
                if (action === "put") records.set(key!, value);
                if (action === "clear") records.clear();
                Object.assign(result, { result: action === "get" ? records.get(key!) : key });
                result.onsuccess?.(new Event("success"));
                transaction.oncomplete?.(new Event("complete"));
              });
              return result;
            };
            Object.assign(transaction, { objectStore: () => ({ get: (key: IDBValidKey) => operation("get", undefined, key), put: (value: unknown, key: IDBValidKey) => operation("put", value, key), clear: () => operation("clear") }) });
            return transaction;
          }
        } as unknown as IDBDatabase;
        Object.assign(request, { result: database });
        if (fresh) request.onupgradeneeded?.(new Event("upgradeneeded") as IDBVersionChangeEvent);
        request.onsuccess?.(new Event("success"));
      });
      return request;
    }
  } as unknown as IDBFactory;
  return { factory, databases, failure, calls };
}
