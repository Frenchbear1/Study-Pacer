// Everything lives on this device in IndexedDB: the original files, what was worked out from them on import,
// the reading position, and the session in progress (so a reload picks up exactly where you were).
const DB = 'study-pacer';
const VERSION = 1;
let dbp;

function open() {
  dbp ??= new Promise((res, rej) => {
    const r = indexedDB.open(DB, VERSION);
    r.onupgradeneeded = () => {
      const db = r.result;
      if (!db.objectStoreNames.contains('docs')) db.createObjectStore('docs', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('files')) db.createObjectStore('files');
      if (!db.objectStoreNames.contains('content')) db.createObjectStore('content');
      if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbp;
}

async function tx(name, mode, fn) {
  const db = await open();
  return new Promise((res, rej) => {
    const t = db.transaction(name, mode);
    const s = t.objectStore(name);
    let out;
    Promise.resolve(fn(s)).then((v) => { out = v; });
    t.oncomplete = () => res(out);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error);
  });
}
const req = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

export const store = {
  allDocs: () => tx('docs', 'readonly', (s) => req(s.getAll())),
  getDoc: (id) => tx('docs', 'readonly', (s) => req(s.get(id))),
  putDoc: (d) => tx('docs', 'readwrite', (s) => req(s.put(d))),
  async deleteDoc(id) {
    const db = await open();
    await new Promise((res, rej) => {
      const t = db.transaction(['docs', 'files', 'content'], 'readwrite');
      t.objectStore('docs').delete(id);
      t.objectStore('files').delete(id);
      t.objectStore('content').delete(id);
      t.oncomplete = res; t.onerror = () => rej(t.error);
    });
  },
  getFile: (id) => tx('files', 'readonly', (s) => req(s.get(id))),
  putFile: (id, blob) => tx('files', 'readwrite', (s) => req(s.put(blob, id))),
  getContent: (id) => tx('content', 'readonly', (s) => req(s.get(id))),
  putContent: (id, c) => tx('content', 'readwrite', (s) => req(s.put(c, id))),
  get: (k) => tx('kv', 'readonly', (s) => req(s.get(k))),
  set: (k, v) => tx('kv', 'readwrite', (s) => req(s.put(v, k))),
  del: (k) => tx('kv', 'readwrite', (s) => req(s.delete(k))),
  async estimate() {
    try { const e = await navigator.storage?.estimate?.(); return e || null; } catch { return null; }
  },
  persist: () => navigator.storage?.persist?.().catch(() => false),
};

export const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
