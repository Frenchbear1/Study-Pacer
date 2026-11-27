// The Windows app (Tauri) gives the page a few extra powers: pick a folder, list the documents in it, and read
// files straight from disk. In a normal browser none of this exists and the app works from imported copies.
const T = window.__TAURI__;
export const isDesktop = !!T?.core;
export const invoke = (cmd, args) => T.core.invoke(cmd, args);
export const fileUrl = (path) => T.core.convertFileSrc(path);
export async function readLocal(path) {
  const r = await fetch(fileUrl(path));
  if (!r.ok) throw new Error(r.status === 404 ? 'the file has moved or been deleted' : `couldn't read the file (${r.status})`);
  return r.blob();
}
export const pickFolder = () => invoke('pick_folder');
export const scanFolder = (root) => invoke('scan', { root });
export const reveal = (path) => invoke('reveal', { path }).catch(() => {});
// a short, stable id for a path (the same file always gets the same id, so progress survives rescans)
export function pathId(s) {
  let h1 = 0x811c9dc5, h2 = 0x01000193;
  const t = s.toLowerCase();
  for (let i = 0; i < t.length; i++) { const c = t.charCodeAt(i); h1 = Math.imul(h1 ^ c, 16777619); h2 = Math.imul(h2 ^ c, 2246822519) ^ (h2 >>> 13); }
  return 'L' + (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36);
}
