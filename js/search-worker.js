// Search runs here, away from the page, so typing never stutters. The text index is read from IndexedDB
// once and kept in memory; a query runs over every page of the documents it's given and reports, per document,
// the pages that match with a short snippet around the first hit on each.
import { compile, inWhere } from './search.js';

let dbp;
const open = () => (dbp ??= new Promise((res, rej) => { const r = indexedDB.open('study-pacer'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }));
const texts = new Map(); // docId -> { pages, marks }
async function load(ids) {
  const miss = ids.filter((id) => !texts.has(id));
  if (!miss.length) return;
  const db = await open();
  if (!db.objectStoreNames.contains('texts')) return;
  await new Promise((res) => {
    const t = db.transaction('texts', 'readonly');
    const s = t.objectStore('texts');
    for (const id of miss) { const r = s.get(id); r.onsuccess = () => { if (r.result) texts.set(id, r.result); }; }
    t.oncomplete = res; t.onerror = res;
  });
}

// a line of text around the first hit, with the hits inside it as [start, end, colour]
function snippet(text, ranges, width = 170) {
  const first = ranges[0] || [0, 0];
  let a = Math.max(0, first[0] - Math.round(width * 0.35));
  let b = Math.min(text.length, a + width);
  a = Math.max(0, Math.min(a, b - width));
  // start and end on word boundaries
  if (a > 0) { const sp = text.indexOf(' ', a); if (sp > -1 && sp < first[0]) a = sp + 1; }
  if (b < text.length) { const sp = text.lastIndexOf(' ', b); if (sp > first[1]) b = sp; }
  const lead = a > 0 ? 1 : 0;
  // whitespace swapped one-for-one so the hit positions stay exact
  const out = (lead ? '…' : '') + text.slice(a, b).replace(/\s/g, ' ') + (b < text.length ? '…' : '');
  const hl = ranges.filter((r) => r[0] >= a && r[1] <= b).map((r) => [r[0] - a + lead, r[1] - a + lead, r[2]]);
  return { s: out, hl };
}
// the heading a hit sits under (the last heading before it on the page)
function headingFor(text, marks, at) {
  if (!marks) return '';
  let h = '';
  for (let i = 0; i < marks.length; i += 3) if (marks[i] === 1 && marks[i + 1] <= at) h = text.slice(marks[i + 1], marks[i + 2]);
  return h.trim().slice(0, 80);
}

let latest = 0;
self.onmessage = async ({ data }) => {
  if (data.t === 'drop') { for (const id of data.ids) texts.delete(id); return; }
  if (data.t !== 'q') return;
  const { qid, query, opts, ids } = data;
  latest = qid;
  const m = compile(query, opts);
  if (!m) { self.postMessage({ qid, empty: true }); return; }
  if (m.error) { self.postMessage({ qid, error: m.error }); return; }
  await load(ids);
  if (qid !== latest) return;
  const results = [];
  let hits = 0, pagesHit = 0, searched = 0, missing = 0;
  const t0 = performance.now();
  for (const id of ids) {
    const t = texts.get(id);
    if (!t) { missing++; continue; }
    searched++;
    const pages = [];
    let count = 0;
    for (let n = 0; n < t.pages.length; n++) {
      let r = m.match(t.pages[n]);
      if (r && opts.where && opts.where !== 'any') r = inWhere(r, t.marks?.[n] || null, opts.where);
      if (!r) continue;
      count += r.length;
      // only the first pages of a document carry a snippet; the rest are just numbers
      pages.push(pages.length < 60 ? { n, c: r.length, ...snippet(t.pages[n], r), h: headingFor(t.pages[n], t.marks?.[n], r[0]?.[0] ?? 0) } : { n, c: r.length });
    }
    if (pages.length) { results.push({ id, pages, count }); hits += count; pagesHit += pages.length; }
    // let a newer query take over
    if (performance.now() - t0 > 40) { await new Promise((r) => setTimeout(r)); if (qid !== latest) return; }
  }
  self.postMessage({ qid, results, hits, pagesHit, searched, missing, terms: m.terms });
};
