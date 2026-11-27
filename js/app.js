import { store } from './store.js';
import { importFiles, ACCEPT, labelOf } from './importers.js';
import { Pacer, weightsFor, suggestTime, clock, human, timeOfDay, METHODS, methodCfg, DEFAULTS } from './pacer.js';
import { Viewer, loadDoc } from './viewer.js';
import { renderSlide } from './pptx.js';
import { THEME_DEFAULTS, applyReadingTheme, openThemeBar, themeBarOpen, blurPx, BLUR_DEF } from './themes.js';

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const reduced = () => matchMedia('(prefers-reduced-motion: reduce)').matches;
const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
const ICON = {
  check: '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  dots: '<svg viewBox="0 0 24 24"><circle cx="5.5" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="18.5" cy="12" r="1.6"/></svg>',
  prev: '<svg viewBox="0 0 24 24"><path d="M15 5l-7 7 7 7"/></svg>',
  next: '<svg viewBox="0 0 24 24"><path d="M9 5l7 7-7 7"/></svg>',
  play: '<svg viewBox="0 0 24 24"><path d="M8 5.5v13l10.5-6.5Z"/></svg>',
  pause: '<svg viewBox="0 0 24 24"><rect x="6.5" y="5" width="3.8" height="14" rx="1.2"/><rect x="13.7" y="5" width="3.8" height="14" rx="1.2"/></svg>',
  close: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  grip: '<svg viewBox="0 0 24 24"><path d="M8 7h.01M8 12h.01M8 17h.01M16 7h.01M16 12h.01M16 17h.01" stroke-width="3.2"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M4 7h16M9 7V4.5h6V7M6.5 7l1 13h9l1-13"/></svg>',
  pen: '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9l-4-4L4 16Z"/></svg>',
  book: '<svg viewBox="0 0 24 24"><path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v15H6.5A2.5 2.5 0 0 0 4 20.5Z"/><path d="M4 20.5A2.5 2.5 0 0 1 6.5 18H20v3H6.5"/></svg>',
  timer: '<svg viewBox="0 0 24 24"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2.5M9.5 2.5h5"/></svg>',
  reset: '<svg viewBox="0 0 24 24"><path d="M4 12a8 8 0 1 0 2.4-5.7L4 8.5"/><path d="M4 3.5v5h5"/></svg>',
  chart: '<svg viewBox="0 0 24 24"><path d="M5 20V10M12 20V4M19 20v-7"/></svg>',
  coffee: '<svg viewBox="0 0 24 24"><path d="M4 9h13v5a5 5 0 0 1-5 5H9a5 5 0 0 1-5-5Z"/><path d="M17 10h1.5a2.5 2.5 0 0 1 0 5H17M8 3v3M12 3v3"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>',
  view: '<svg viewBox="0 0 24 24"><rect x="5" y="3.5" width="14" height="17" rx="2"/><path d="M9 8h6M9 12h6M9 16h3"/></svg>',
  zoom: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4 4M8.5 11h5M11 8.5v5"/></svg>',
  gear: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M12 2.5v3M12 18.5v3M2.5 12h3M18.5 12h3M5.3 5.3l2.1 2.1M16.6 16.6l2.1 2.1M5.3 18.7l2.1-2.1M16.6 7.4l2.1-2.1"/></svg>',
  stop: '<svg viewBox="0 0 24 24"><rect x="6" y="6" width="12" height="12" rx="2.5"/></svg>',
  notes: '<svg viewBox="0 0 24 24"><path d="M5 4h14v12l-4 4H5Z"/><path d="M15 20v-4h4M8.5 9h7M8.5 13h4"/></svg>',
  minus: '<svg viewBox="0 0 24 24"><path d="M5 12h14"/></svg>',
};

// ================= preferences =================
const PKEY = 'study-pacer:prefs';
const PREF_DEFAULTS = {
  ...THEME_DEFAULTS, glassT: 50, blurL: BLUR_DEF, glassG: {},
  view: 'paged', fit: 'page', turn: 'slide', tapEdges: true, notes: true,
  fontSize: 18, flowW: 720, flowFont: 'serif', lineH: 1.65,
  autoPause: true, sound: true, haptics: true, wake: true, nudge: true, hudCompact: false,
  sort: 'recent', setup: null, customMethod: null,
};
let prefs;
try { prefs = { ...PREF_DEFAULTS, ...JSON.parse(localStorage.getItem(PKEY) || '{}') }; } catch { prefs = { ...PREF_DEFAULTS }; }
const savePrefs = () => { try { localStorage.setItem(PKEY, JSON.stringify(prefs)); } catch {} };

// ================= glass (ported from Findly) =================
const GLASS = { '--a-top': 76, '--a-bar': 60, '--a-menu': 82, '--a-sheet': 97, '--a-tbar': 94 };
const GLASS_GROUPS = {
  top: { vars: ['--a-top'], blur: '--blur-top' },
  bottom: { vars: ['--a-bar'], blur: '--blur-bottom' },
  menus: { vars: ['--a-menu'], blur: '--blur-pop' },
  sheets: { vars: ['--a-sheet', '--a-tbar'], blur: '--blur-sheet' },
};
// out of the box: a little see-through with the slightest blur; sheets a touch more solid
const GLASS_START = { top: { t: 62, b: 34 }, bottom: { t: 62, b: 38 }, menus: { t: 62, b: 38 }, sheets: { t: 58, b: 45 } };
const glassOf = (g) => { const o = prefs.glassG?.[g]; const s = GLASS_START[g]; return { t: o?.t ?? s?.t ?? prefs.glassT ?? 50, b: o?.b ?? s?.b ?? prefs.blurL ?? BLUR_DEF }; };
function applyGlass() {
  const root = document.documentElement.style;
  for (const [g, { vars, blur }] of Object.entries(GLASS_GROUPS)) {
    const { t: tt, b } = glassOf(g);
    const t = Math.max(0, Math.min(100, tt)) / 100;
    for (const k of vars) {
      const base = GLASS[k];
      const a = t <= 0.5 ? 100 + (base - 100) * (t / 0.5) : base + (6 - base) * ((t - 0.5) / 0.5);
      root.setProperty(k, `${a.toFixed(1)}%`);
    }
    root.setProperty(blur, `saturate(180%) blur(${blurPx(b).toFixed(1)}px)`);
  }
  const sa = parseFloat(root.getPropertyValue('--a-sheet')) || 100;
  root.setProperty('--a-sheet-card', `${(sa + (100 - sa) * 0.3).toFixed(1)}%`);
  root.setProperty('--bpx', `${blurPx(prefs.blurL ?? BLUR_DEF).toFixed(1)}px`);
}
function applyFlowPrefs() {
  const r = document.documentElement.style;
  r.setProperty('--flow-font', prefs.flowFont === 'sans' ? 'var(--ui)' : prefs.flowFont === 'mono' ? 'var(--mono)' : prefs.flowFont === 'plain' ? 'Verdana, "Atkinson Hyperlegible", Tahoma, sans-serif' : 'var(--read)');
  r.setProperty('--flow-lh', String(prefs.lineH || 1.65));
  document.body.classList.toggle('tap-edges', !!prefs.tapEdges);
}
const openThemes = () => { closeMenu(); openThemeBar(prefs, { save: savePrefs, glass: applyGlass, glassOf }); };
applyReadingTheme(prefs);
applyGlass();
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if ((prefs.rtheme || 'auto') === 'auto') applyReadingTheme(prefs); });

// ================= small UI pieces =================
function toast(html, { ms = 3200, actions = [], plain = false, id } = {}) {
  const host = $('#toasts');
  if (id) host.querySelector(`[data-id="${id}"]`)?.remove();
  const t = document.createElement('div');
  t.className = `toast${plain || !actions.length ? ' plain' : ''}`;
  if (id) t.dataset.id = id;
  t.innerHTML = `<span>${html}</span>${actions.map((a, i) => `<button class="btn${a.primary ? ' primary' : ''}" data-i="${i}">${esc(a.label)}</button>`).join('')}`;
  const close = () => { if (!t.isConnected || t.classList.contains('out')) return; t.classList.add('out'); setTimeout(() => t.remove(), 260); };
  t.onclick = (e) => { const b = e.target.closest('[data-i]'); if (b) { actions[Number(b.dataset.i)].run?.(); close(); } };
  host.append(t);
  if (ms) setTimeout(close, ms);
  t.close = close;
  return t;
}
let menuEl = null;
function closeMenu() {
  if (!menuEl) return;
  const m = menuEl; menuEl = null;
  m.classList.add('out'); setTimeout(() => m.remove(), 170);
}
function openMenu(anchor, items) {
  closeMenu();
  const m = document.createElement('div');
  m.className = 'menu glass-menu';
  m.innerHTML = items.map((it, i) => (it === '-' ? '<hr>' : `<button data-i="${i}" class="${it.danger ? 'danger' : ''}">${it.icon ? ICON[it.icon] || it.icon : ''}<span>${esc(it.label)}</span>${it.key ? `<span class="m-key">${esc(it.key)}</span>` : ''}</button>`)).join('');
  document.body.append(m);
  const r = anchor.getBoundingClientRect();
  const mw = m.offsetWidth, mh = m.offsetHeight;
  let left = Math.min(innerWidth - mw - 8, Math.max(8, r.right - mw));
  let top = r.bottom + 6;
  let oy = 'top';
  if (top + mh > innerHeight - 8) { top = Math.max(8, r.top - mh - 6); oy = 'bottom'; }
  m.style.left = `${left}px`; m.style.top = `${top}px`;
  m.style.setProperty('--ox', `${r.left + r.width / 2 - left}px ${oy}`);
  m.onclick = (e) => { const b = e.target.closest('[data-i]'); if (!b) return; const it = items[Number(b.dataset.i)]; closeMenu(); it.run?.(); };
  menuEl = m;
  setTimeout(() => document.addEventListener('pointerdown', function off(e) { if (menuEl === m && !m.contains(e.target)) closeMenu(); document.removeEventListener('pointerdown', off, true); }, true));
}

// Sheets: slide up from the bottom (phone) or float in the middle (wide screens); swipe down on the grab bar
// or from the top of the content to close. Nothing behind scrolls while one is open.
const sheets = [];
function openSheet({ title, body = '', foot = '', wide = false, onClose, cls = '' }) {
  const scrim = document.createElement('div');
  scrim.className = 'scrim';
  const sh = document.createElement('section');
  sh.className = `sheet${wide ? ' wide' : ''} ${cls}`;
  sh.setAttribute('role', 'dialog');
  sh.innerHTML = `<div class="grab"></div><header class="sh-head"><h3>${esc(title)}</h3><button class="icon-btn" data-close aria-label="Close">${ICON.close}</button></header><div class="sh-body">${body}</div>${foot ? `<footer class="sh-foot">${foot}</footer>` : ''}`;
  document.body.append(scrim, sh);
  document.body.classList.add('locked');
  const desk = () => innerWidth >= 760;
  const inF = desk() ? [{ opacity: 0, transform: 'translate(-50%, -48%) scale(0.95)' }, { opacity: 1, transform: 'translate(-50%, -50%)' }] : [{ transform: 'translate(-50%, 100%)' }, { transform: 'translate(-50%, 0)' }];
  sh.animate(inF, { duration: reduced() ? 1 : 460, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.04)' });
  let closed = false;
  const close = (fast) => {
    if (closed) return; closed = true;
    const i = sheets.indexOf(api); if (i >= 0) sheets.splice(i, 1);
    if (!sheets.length) document.body.classList.remove('locked');
    scrim.classList.add('out');
    const cur = getComputedStyle(sh).transform;
    const outF = desk() ? [{ opacity: 1 }, { opacity: 0, transform: 'translate(-50%, -48%) scale(0.96)' }] : [{ transform: cur === 'none' ? 'translate(-50%, 0)' : cur }, { transform: 'translate(-50%, 105%)' }];
    sh.animate(outF, { duration: fast || reduced() ? 1 : 280, easing: 'cubic-bezier(0.4, 0, 1, 1)', fill: 'forwards' }).onfinish = () => { sh.remove(); scrim.remove(); };
    onClose?.();
  };
  scrim.onclick = () => close();
  sh.querySelector('[data-close]').onclick = () => close();
  // swipe to dismiss
  let drag = null;
  const bodyEl = sh.querySelector('.sh-body');
  sh.addEventListener('pointerdown', (e) => {
    if (desk() || e.pointerType === 'mouse') return;
    const onGrab = e.target.closest('.grab, .sh-head');
    if (!onGrab && (bodyEl.scrollTop > 0 || e.target.closest('input, textarea, .q-handle, [data-noswipe]'))) return;
    drag = { y: e.clientY, t: performance.now(), dy: 0, grab: !!onGrab };
  });
  sh.addEventListener('pointermove', (e) => {
    if (!drag) return;
    const dy = e.clientY - drag.y;
    if (!drag.on && dy > 10 && (drag.grab || bodyEl.scrollTop <= 0)) { drag.on = true; sh.setPointerCapture(e.pointerId); }
    if (drag.on) { drag.dy = Math.max(0, dy); sh.style.transform = `translate(-50%, ${drag.dy}px)`; }
  });
  const up = () => {
    if (!drag) return;
    const d = drag; drag = null;
    if (!d.on) return;
    const v = d.dy / (performance.now() - d.t);
    if (d.dy > 120 || v > 0.6) close();
    else { sh.animate([{ transform: `translate(-50%, ${d.dy}px)` }, { transform: 'translate(-50%, 0)' }], { duration: 320, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.2)' }); sh.style.transform = ''; }
  };
  sh.addEventListener('pointerup', up); sh.addEventListener('pointercancel', up);
  const api = { el: sh, body: bodyEl, close, set foot(h) { const f = sh.querySelector('.sh-foot'); if (f) f.innerHTML = h; } };
  sheets.push(api);
  return api;
}
const sw = (k, on, extra = '') => `<label class="switch"><input type="checkbox" data-k="${k}" ${on ? 'checked' : ''} ${extra}><i></i></label>`;
const fillTrack = (r) => r.style.setProperty('--fillpct', `${((r.value - r.min) / (r.max - r.min)) * 100}%`);

// ================= sound and touch =================
let actx;
function chime(kind) {
  if (!prefs.sound) return;
  try {
    actx ??= new (window.AudioContext || window.webkitAudioContext)();
    if (actx.state === 'suspended') actx.resume();
    const seq = { due: [659, 784, 988], end: [784, 659], finish: [523, 659, 784, 1047], soft: [880] }[kind] || [660];
    const t0 = actx.currentTime + 0.02;
    seq.forEach((f, i) => {
      const o = actx.createOscillator(), g = actx.createGain();
      o.type = 'sine'; o.frequency.value = f;
      const t = t0 + i * 0.13;
      g.gain.setValueAtTime(0, t); g.gain.linearRampToValueAtTime(0.16, t + 0.02); g.gain.exponentialRampToValueAtTime(0.0001, t + 0.9);
      o.connect(g).connect(actx.destination); o.start(t); o.stop(t + 1);
    });
  } catch {}
}
const buzz = (p) => { if (prefs.haptics) try { navigator.vibrate?.(p); } catch {} };
let wake = null;
async function keepAwake(on) {
  try {
    if (on && prefs.wake && 'wakeLock' in navigator && document.visibilityState === 'visible') { if (!wake) { wake = await navigator.wakeLock.request('screen'); wake.addEventListener('release', () => { wake = null; }); } }
    else if (!on && wake) { await wake.release(); wake = null; }
  } catch {}
}

// ================= library =================
let docs = [];
const selected = new Set();
const docById = (id) => docs.find((d) => d.id === id);
const unitWord = (d, n = 2) => (d.kind === 'slides' ? (n === 1 ? 'slide' : 'slides') : n === 1 ? 'page' : 'pages');
const readCount = (d) => (d.read || []).length;

async function loadLibrary() {
  docs = await store.allDocs();
  renderLibrary();
}
function sortedDocs() {
  const s = prefs.sort;
  const list = [...docs];
  if (s === 'name') list.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  else if (s === 'progress') list.sort((a, b) => readCount(b) / b.units - readCount(a) / a.units || (b.opened || b.added) - (a.opened || a.added));
  else list.sort((a, b) => (b.opened || b.added) - (a.opened || a.added));
  return list;
}
function estimateFor(d) { return suggestTime(d.words.map((w) => ({ words: w, kind: d.kind })), speed); }
function cardHtml(d, i) {
  const pct = d.units ? Math.round((readCount(d) / d.units) * 100) : 0;
  const wide = (d.aspect || 0.77) > 1.05;
  let art;
  if (d.thumb) art = `<img src="${d.thumb}" alt="" loading="lazy" draggable="false">`;
  else if (d.kind === 'slides') art = `<div class="mini-slide" data-mini="${d.id}"><div style="--ar:${d.aspect || 1.78}"></div></div>`;
  else art = `<div class="mini-flow" data-flow="${d.id}"><b>${esc(d.name)}</b></div>`;
  return `<div class="card${selected.has(d.id) ? ' sel' : ''}" role="listitem" data-id="${d.id}" style="--i:${Math.min(i, 16)}" tabindex="0">
    <div class="card-art${wide ? ' wide' : ''}${d.kind !== 'flow' ? ' tintable' : ''}">${art}
      <span class="card-badge ${d.kind}">${esc(labelOf(d))}</span>
      ${pct ? `<span class="card-prog"><i style="width:${pct}%"></i></span>` : ''}
    </div>
    <button class="card-check" data-check aria-label="Select ${esc(d.name)}">${ICON.check}</button>
    <button class="card-menu" data-cmenu aria-label="More for ${esc(d.name)}">${ICON.dots}</button>
    <div class="card-meta"><b>${esc(d.name)}</b><small>${d.units} ${unitWord(d, d.units)} · ~${human(estimateFor(d))}${pct ? ` · ${pct}%` : ''}</small></div>
  </div>`;
}
function renderLibrary() {
  const grid = $('#grid');
  const list = sortedDocs();
  $('#empty').hidden = list.length > 0;
  $('#libHead').hidden = !list.length;
  grid.innerHTML = list.map(cardHtml).join('');
  grid.classList.toggle('selecting', selected.size > 0);
  $$('#sortSeg button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.sort === prefs.sort)));
  // previews that need the document itself
  for (const el of $$('[data-mini]', grid)) miniSlide(el);
  for (const el of $$('[data-flow]', grid)) miniFlow(el);
  renderSelbar();
  renderHero();
}
async function miniSlide(el) {
  const d = docById(el.dataset.mini);
  try {
    const x = await loadDoc(d);
    const box = el.firstElementChild;
    const s = renderSlide(x.deck, 0, x.urls);
    box.append(s);
    const fit = () => { const w = box.clientWidth; if (w) s.style.transform = `scale(${w / (x.deck.w * (96 / 72))})`; };
    fit(); new ResizeObserver(fit).observe(box);
  } catch {}
}
async function miniFlow(el) {
  const d = docById(el.dataset.flow);
  try {
    const c = await store.getContent(d.id);
    const tmp = document.createElement('div');
    tmp.innerHTML = c.chunks[0] || '';
    tmp.querySelectorAll('h1,h2,h3,img,table,pre').forEach((n) => n.remove());
    el.insertAdjacentHTML('beforeend', [...tmp.querySelectorAll('p, li')].slice(0, 8).map((p) => `<p>${esc(p.textContent.slice(0, 220))}</p>`).join(''));
  } catch {}
}
function renderSelbar() {
  const bar = $('#selbar');
  const was = !bar.hidden;
  const ids = [...selected].filter(docById);
  bar.hidden = !ids.length;
  if (!ids.length) return;
  const ds = ids.map(docById);
  const units = ds.reduce((a, d) => a + d.units, 0);
  const est = ds.reduce((a, d) => a + estimateFor(d), 0);
  $('#selCount').textContent = ids.length === 1 ? ds[0].name : `${ids.length} documents`;
  $('#selInfo').textContent = `${units} pages · about ${human(est)}`;
  if (!was) { bar.classList.remove('in'); void bar.offsetWidth; bar.classList.add('in'); }
}
async function renderHero() {
  const hero = $('#hero');
  const ses = await store.get('session');
  const hist = (await store.get('history')) || [];
  let html = '';
  if (ses && !ses.pacer?.s?.finished) {
    const p = Pacer.restore(ses);
    const sn = p.snapshot();
    const pct = sn.doneN / sn.n;
    const C = 2 * Math.PI * 27;
    const names = [...new Set(ses.units.map((u) => docById(u.doc)?.name).filter(Boolean))];
    html += `<div class="hero"><div class="hero-ring"><svg viewBox="0 0 66 66"><circle class="tr" cx="33" cy="33" r="27"/><circle class="ar" cx="33" cy="33" r="27" stroke-dasharray="${C}" stroke-dashoffset="${C * (1 - pct)}"/></svg><b>${Math.round(pct * 100)}%</b></div>
      <div class="hero-text"><small>Session in progress</small><h3>${esc(names.slice(0, 2).join(' + '))}${names.length > 2 ? ` + ${names.length - 2} more` : ''}</h3><p>${plural(sn.left, 'page')} to go · ${sn.open ? `${human(sn.work)} so far` : sn.remaining > 0 ? `${human(sn.remaining)} left` : 'out of time'}</p></div>
      <div class="hero-act"><button class="btn" data-act="endSaved">End</button><button class="btn primary" data-act="resume">${ICON.play}<span>Resume</span></button></div></div>`;
  }
  if (hist.length) {
    const wk = Date.now() - 7 * 864e5;
    const recent = hist.filter((h) => h.t > wk);
    const today = hist.filter((h) => new Date(h.t).toDateString() === new Date().toDateString());
    const sum = (a, k) => a.reduce((s, h) => s + (h[k] || 0), 0);
    const streak = (() => { let n = 0; const days = new Set(hist.map((h) => new Date(h.t).toDateString())); for (let d = new Date(); days.has(d.toDateString()); d.setDate(d.getDate() - 1)) n++; return n; })();
    html += `<div class="week"><div><b>${human(sum(today, 'work'))}</b><span>today</span></div><div><b>${human(sum(recent, 'work'))}</b><span>this week</span></div><div><b>${sum(recent, 'pages')}</b><span>pages this week</span></div>${speed.wpm ? `<div><b>${speed.wpm} wpm</b><span>your study pace</span></div>` : ''}${streak > 1 ? `<div><b>${streak} days</b><span>streak</span></div>` : ''}</div>`;
  }
  hero.innerHTML = html;
}

// importing
const filePick = $('#filePick');
filePick.accept = ACCEPT;
filePick.onchange = () => { if (filePick.files.length) doImport([...filePick.files]); filePick.value = ''; };
async function doImport(files) {
  store.persist();
  const t = toast(`<span>Importing…</span> <span class="prog"><i></i></span>`, { ms: 0, plain: true, id: 'imp' });
  const bar = t.querySelector('.prog i');
  const lab = t.querySelector('span span') || t.firstElementChild.firstElementChild;
  const { docs: added, errors } = await importFiles(files, (i, p, name) => {
    bar.style.setProperty('--p', `${((i + p) / files.length) * 100}%`);
    lab.textContent = files.length > 1 ? `Reading ${i + 1} of ${files.length}` : `Reading ${name.length > 28 ? name.slice(0, 26) + '…' : name}`;
  });
  t.close();
  docs.push(...added);
  added.forEach((d) => selected.add(d.id));
  renderLibrary();
  if (added.length) toast(`${ICON.check} Added ${added.length === 1 ? `“${esc(added[0].name)}”` : plural(added.length, 'document')}`, { actions: [{ label: 'Pace it', primary: true, run: () => openSetup([...added.map((d) => d.id)]) }], ms: 5000 });
  for (const e of errors) toast(esc(e), { ms: 6000 });
}
// drag and drop, anywhere
let dragDepth = 0;
addEventListener('dragenter', (e) => { if (![...(e.dataTransfer?.types || [])].includes('Files')) return; dragDepth++; $('#dropVeil').hidden = false; });
addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('#dropVeil').hidden = true; } });
addEventListener('dragover', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) e.preventDefault(); });
addEventListener('drop', (e) => {
  if (!e.dataTransfer?.files?.length) return;
  e.preventDefault(); dragDepth = 0; $('#dropVeil').hidden = true;
  if (!$('#reader').hidden) exitReader();
  doImport([...e.dataTransfer.files]);
});
// a share target / launch queue (installed app: "Open with Study Pacer")
if ('launchQueue' in window) window.launchQueue.setConsumer(async (p) => { if (p.files?.length) doImport(await Promise.all(p.files.map((h) => h.getFile()))); });

$('#grid').addEventListener('click', (e) => {
  const card = e.target.closest('.card');
  if (!card) return;
  const id = card.dataset.id;
  if (e.target.closest('[data-cmenu]')) { cardMenu(e.target.closest('[data-cmenu]'), id); return; }
  if (e.target.closest('[data-check]') || e.metaKey || e.ctrlKey || e.shiftKey || (selected.size && matchMedia('(hover: none)').matches)) { toggleSel(id, card); return; }
  openFree([id], card);
});
$('#grid').addEventListener('keydown', (e) => { const c = e.target.closest('.card'); if (c && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); e.key === ' ' ? toggleSel(c.dataset.id, c) : openFree([c.dataset.id], c); } });
// long press selects (phones)
let lp;
$('#grid').addEventListener('pointerdown', (e) => { const c = e.target.closest('.card'); if (!c || e.pointerType === 'mouse') return; lp = setTimeout(() => { buzz(12); toggleSel(c.dataset.id, c); c.addEventListener('click', (ev) => { ev.stopPropagation(); }, { capture: true, once: true }); }, 480); });
['pointerup', 'pointercancel', 'pointermove'].forEach((t) => $('#grid').addEventListener(t, (e) => { if (t !== 'pointermove' || Math.abs(e.movementX) + Math.abs(e.movementY) > 3) clearTimeout(lp); }));
function toggleSel(id, card) {
  selected.has(id) ? selected.delete(id) : selected.add(id);
  card.classList.toggle('sel', selected.has(id));
  $('#grid').classList.toggle('selecting', selected.size > 0);
  if (selected.has(id) && !reduced()) card.querySelector('.card-art').animate([{ transform: 'scale(1)' }, { transform: 'scale(0.95)' }, { transform: 'scale(1)' }], { duration: 320, easing: 'cubic-bezier(.2,.9,.3,1.3)' });
  renderSelbar();
}
function cardMenu(btn, id) {
  const d = docById(id);
  openMenu(btn, [
    { label: 'Read', icon: 'book', run: () => openFree([id]) },
    { label: 'Pace it', icon: 'timer', run: () => openSetup([id]) },
    '-',
    { label: 'Rename', icon: 'pen', run: () => renameDoc(d) },
    ...(readCount(d) ? [{ label: 'Reset progress', icon: 'reset', run: async () => { d.read = []; d.pos = 0; await store.putDoc(d); renderLibrary(); } }] : []),
    { label: 'Delete', icon: 'trash', danger: true, run: () => deleteDocs([id]) },
  ]);
}
function renameDoc(d) {
  const s = openSheet({ title: 'Rename', body: `<div class="group"><div class="row"><input class="field-in text" id="rn" value="${esc(d.name)}" maxlength="140"></div></div>`, foot: '<span class="grow"></span><button class="btn primary" data-save>Save</button>' });
  const inp = $('#rn', s.el); setTimeout(() => { inp.focus(); inp.select(); }, 300);
  const save = async () => { const v = inp.value.trim(); if (v) { d.name = v; await store.putDoc(d); renderLibrary(); } s.close(); };
  s.el.querySelector('[data-save]').onclick = save;
  inp.onkeydown = (e) => { if (e.key === 'Enter') save(); };
}
async function deleteDocs(ids) {
  const gone = ids.map(docById).filter(Boolean);
  for (const id of ids) { await store.deleteDoc(id); selected.delete(id); }
  docs = docs.filter((d) => !ids.includes(d.id));
  // a saved session that used them can't continue
  const ses = await store.get('session');
  if (ses && ses.units.some((u) => ids.includes(u.doc))) await store.del('session');
  renderLibrary();
  toast(`Deleted ${gone.length === 1 ? `“${esc(gone[0].name)}”` : plural(gone.length, 'document')}`);
}

// ================= reading speed & history =================
let speed = { wpm: null, slide: null };
async function loadSpeed() { speed = (await store.get('speed')) || speed; }
async function learnSpeed(p) {
  const st = p.stats();
  let changed = false;
  if (st.wpm && st.wpm > 40 && st.wpm < 900) { speed.wpm = Math.round(speed.wpm ? speed.wpm * 0.6 + st.wpm * 0.4 : st.wpm); changed = true; }
  const sl = st.pages.filter((x) => x.done && docById(x.doc)?.kind === 'slides');
  if (sl.length >= 3) { const a = sl.reduce((s, x) => s + x.sec, 0) / sl.length; speed.slide = Math.round(speed.slide ? speed.slide * 0.6 + a * 0.4 : a); changed = true; }
  if (changed) await store.set('speed', speed);
  const hist = (await store.get('history')) || [];
  if (st.work > 60) {
    hist.push({ t: Date.now(), work: Math.round(st.work), pages: st.pages.filter((x) => x.done).length, words: st.words, wpm: st.wpm, docs: [...new Set(p.units.map((u) => u.doc))].map((id) => docById(id)?.name) });
    await store.set('history', hist.slice(-300));
  }
}

// ================= setup: pace a set of documents =================
function unitsFor(cfgDocs) {
  const units = [];
  for (const { id, from, to } of cfgDocs) {
    const d = docById(id);
    if (!d) continue;
    for (let i = from; i <= to; i++) units.push({ doc: id, i, words: d.words[i] || 0, kind: d.kind });
  }
  return units;
}
function openSetup(ids, { fromReader = null } = {}) {
  closeMenu();
  const last = prefs.setup || {};
  const st = {
    docs: ids.map(docById).filter(Boolean).map((d) => {
      // start where you left off when part of it is read
      const read = new Set(d.read || []);
      let from = 0; while (from < d.units - 1 && read.has(from)) from++;
      if (from >= d.units - 1 && read.size >= d.units) from = 0;
      return { id: d.id, from: fromReader?.doc === d.id ? fromReader.i : from, to: d.units - 1 };
    }),
    mode: last.mode || 'total', total: null, per: last.per || 60, by: null,
    profile: last.profile || 'smart', strength: last.strength ?? 70, rebalance: last.rebalance || 'spread',
    method: last.method || 'pomodoro', m: { ...(last.m || {}) }, interrupt: last.interrupt ?? true, recall: last.recall ?? true, breakDocs: last.breakDocs ?? false,
  };
  const units = () => unitsFor(st.docs);
  const suggest = () => suggestTime(units(), speed);
  st.total = Math.max(5 * 60, Math.round(suggest() / 300) * 300 || 1800);
  const d0 = new Date(Date.now() + st.total * 1000 + 15 * 60000); d0.setMinutes(Math.ceil(d0.getMinutes() / 5) * 5, 0, 0);
  st.by = d0.getTime();
  const method = () => (st.method === 'custom' ? { id: 'custom', name: 'Custom', ...(prefs.customMethod || { focus: 30, short: 5, long: 15, every: 4 }) } : METHODS.find((m) => m.id === st.method) || METHODS[0]);
  const mvals = () => ({ ...method(), ...st.m });

  const s = openSheet({ title: 'Pace it', wide: false, body: '<div id="su"></div>', foot: '<div class="grow" id="suSum"></div><button class="btn primary big" data-start>Start</button>' });
  const root = $('#su', s.el);
  const timeHtml = () => {
    if (st.mode === 'total') return `<div class="big-time"><div class="stepper"><button data-tadj="-300" aria-label="5 minutes less">−</button></div><div class="readout"><b>${human(st.total).replace(' min', 'm').replace(' h', 'h')}</b><small>to read it all</small></div><div class="stepper"><button data-tadj="300" aria-label="5 minutes more">+</button></div></div>
      <div class="chips">${[15, 30, 45, 60, 90, 120, 180].map((m) => `<button class="chip${st.total === m * 60 ? ' on' : ''}" data-tset="${m * 60}">${m < 60 ? `${m}m` : `${m / 60}h`}</button>`).join('')}<button class="chip accent" data-tset="${Math.round(suggest() / 60) * 60}">Suggested ~${human(suggest())}</button></div>`;
    if (st.mode === 'per') return `<div class="big-time"><div class="stepper"><button data-padj="-5">−</button></div><div class="readout"><b>${clock(st.per, { hours: false })}</b><small>average a page</small></div><div class="stepper"><button data-padj="5">+</button></div></div>
      <div class="chips">${[20, 30, 45, 60, 90, 120, 180, 300].map((v) => `<button class="chip${st.per === v ? ' on' : ''}" data-pset="${v}">${v < 60 ? `${v}s` : `${v / 60}m`}</button>`).join('')}</div>`;
    if (st.mode === 'by') {
      const dt = new Date(st.by);
      const hh = String(dt.getHours()).padStart(2, '0'), mm = String(dt.getMinutes()).padStart(2, '0');
      const left = (st.by - Date.now()) / 1000;
      return `<div class="big-time"><div class="readout"><input type="time" id="byTime" value="${hh}:${mm}"><small>${left > 60 ? `${human(left)} from now${left > 86400 ? '' : ''}` : 'pick a time later today'}</small></div></div>
        <div class="chips">${[30, 60, 90, 120].map((m) => `<button class="chip" data-byin="${m}">in ${m < 60 ? `${m}m` : `${m / 60}h`}</button>`).join('')}</div>`;
    }
    return `<p class="method-blurb" style="padding-top:12px;text-align:center">No deadline. Each page gets a target from your usual pace (${speed.wpm ? `${speed.wpm} words a minute` : 'a careful study pace until it learns yours'}), and your focus cycles run as set below.</p>`;
  };
  const planBars = () => {
    const u = units();
    const w = weightsFor(u, st);
    if (!w.length) return '';
    const n = Math.min(120, w.length);
    const bins = Array.from({ length: n }, (_, b) => { const a = Math.floor((b * w.length) / n), z = Math.max(a + 1, Math.floor(((b + 1) * w.length) / n)); let s2 = 0; for (let i = a; i < z; i++) s2 += w[i]; return s2 / (z - a); });
    const max = Math.max(...bins);
    return `<div class="plan-chart">${bins.map((v) => `<i style="--h:${Math.max(6, (v / max) * 100)}%"></i>`).join('')}</div><div class="plan-cap"><span>first page</span><span>time per page, as planned</span><span>last</span></div>`;
  };
  const methodBar = (m) => {
    if (m.flowtime) return '<span class="m-bar"><i style="flex:5"></i><i class="r" style="flex:1"></i><i style="flex:3"></i><i class="r" style="flex:.6"></i></span>';
    if (m.pages) return '<span class="m-bar"><i style="flex:3"></i><i class="r" style="flex:.6"></i><i style="flex:3"></i><i class="r" style="flex:.6"></i></span>';
    if (!m.focus) return '<span class="m-bar"><i style="flex:1"></i></span>';
    const cyc = m.every || 2; let h = '';
    for (let i = 0; i < Math.min(cyc, 4); i++) { h += `<i style="flex:${m.focus}"></i>`; h += `<i class="${m.long && i === cyc - 1 ? 'L' : 'r'}" style="flex:${m.long && i === cyc - 1 ? m.long : m.short}"></i>`; }
    return `<span class="m-bar">${h}</span>`;
  };
  const pattern = (m) => (m.flowtime ? 'break when you need' : m.pages ? `${m.pages} pages / ${m.short}m` : m.focus ? `${m.focus} / ${m.short}${m.long ? ` · ${m.long}m long` : ''}` : 'free');
  const render = () => {
    const u = units();
    const m = mvals();
    root.innerHTML = `
      <div class="group-t">Reading</div>
      <div class="group" id="queue">${st.docs.map((q, k) => { const d = docById(q.id); return `<div class="qrow" data-q="${k}">
        ${st.docs.length > 1 ? `<span class="q-handle" aria-label="Drag to reorder">${ICON.grip}</span>` : ''}
        <span class="q-thumb">${d.thumb ? `<img src="${d.thumb}" alt="">` : esc(labelOf(d))}</span>
        <div class="q-text"><b>${esc(d.name)}</b><div class="q-range">${unitWord(d)} <input class="field-in" type="number" min="1" max="${d.units}" value="${q.from + 1}" data-from="${k}" inputmode="numeric"> to <input class="field-in" type="number" min="1" max="${d.units}" value="${q.to + 1}" data-to="${k}" inputmode="numeric"> of ${d.units}</div></div>
        ${st.docs.length > 1 ? `<button class="icon-btn" data-qdel="${k}" aria-label="Remove">${ICON.close}</button>` : ''}
      </div>`; }).join('')}</div>
      <div class="group-t">Time</div>
      <div class="group"><div class="row" style="padding-bottom:0"><div class="seg wide" role="group">${[['total', 'Total time'], ['per', 'Per page'], ['by', 'Finish by'], ['open', 'Open']].map(([k, l]) => `<button data-mode="${k}" aria-pressed="${st.mode === k}">${l}</button>`).join('')}</div></div>${timeHtml()}</div>
      <div class="group-t">Study method</div>
      <div class="group"><div class="methods">${[...METHODS, { id: 'custom', name: 'Custom', ...(prefs.customMethod || { focus: 30, short: 5, long: 15, every: 4 }) }].map((x) => `<button class="method${st.method === x.id ? ' on' : ''}" data-method="${x.id}"><b>${esc(x.name)}</b><span class="m-pat">${esc(pattern(x))}</span>${methodBar(x)}</button>`).join('')}</div>
        <p class="method-blurb">${esc(method().blurb || 'Your own rhythm. Change any of the numbers below.')}</p>
        ${m.focus && !m.flowtime ? `<div class="row"><div class="rt"><b>Focus</b></div><div class="stepper"><button data-mv="focus:-5">−</button><output>${m.focus} min</output><button data-mv="focus:5">+</button></div></div>
        <div class="row"><div class="rt"><b>Short break</b></div><div class="stepper"><button data-mv="short:-1">−</button><output>${m.short} min</output><button data-mv="short:1">+</button></div></div>
        <div class="row"><div class="rt"><b>Long break</b><small>${m.long ? `after every ${m.every || 4} focus blocks` : 'off'}</small></div><div class="stepper"><button data-mv="long:-5">−</button><output>${m.long ? `${m.long} min` : 'Off'}</output><button data-mv="long:5">+</button></div></div>
        ${m.long ? `<div class="row"><div class="rt"><b>Long break every</b></div><div class="stepper"><button data-mv="every:-1">−</button><output>${m.every || 4} blocks</output><button data-mv="every:1">+</button></div></div>` : ''}
        <div class="row"><div class="rt"><b>Interrupt me when time’s up</b><small>${st.interrupt ? 'Stops you mid-page; you can take the break, ask for more time, finish the page or skip it' : 'Waits for you to turn the page'}</small></div>${sw('interrupt', st.interrupt)}</div>` : ''}
        ${m.pages ? `<div class="row"><div class="rt"><b>Break every</b></div><div class="stepper"><button data-mv="pages:-1">−</button><output>${m.pages} pages</output><button data-mv="pages:1">+</button></div></div><div class="row"><div class="rt"><b>For</b></div><div class="stepper"><button data-mv="short:-1">−</button><output>${m.short} min</output><button data-mv="short:1">+</button></div></div>` : ''}
        ${m.id !== 'none' ? `<div class="row"><div class="rt"><b>Recall check at breaks</b><small>Before you rest, try to recall what you just read without looking. Retrieval practice is one of the most reliable ways to remember.</small></div>${sw('recall', st.recall)}</div>` : ''}
        ${st.docs.length > 1 ? `<div class="row"><div class="rt"><b>Break between documents</b></div>${sw('breakDocs', st.breakDocs)}</div>` : ''}
      </div>
      <div class="group-t">Pacing</div>
      <div class="group">
        <div class="row stack"><div class="seg wide" role="group">${[['smart', 'Smart'], ['even', 'Even'], ['front', 'Front-load'], ['back', 'Back-load']].map(([k, l]) => `<button data-profile="${k}" aria-pressed="${st.profile === k}">${l}</button>`).join('')}</div>
          <small class="muted">${{ smart: 'Pages with more on them get more time; a title slide gets a moment.', even: 'Every page gets the same time.', front: 'More time early, while you’re fresh.', back: 'More time towards the end, for the harder parts.' }[st.profile]}</small></div>
        ${st.profile === 'smart' ? `<div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>How much content matters</b><output>${st.strength}%</output></div><input type="range" class="sl" min="0" max="100" value="${st.strength}" data-strength></div>` : ''}
        ${planBars()}
      </div>
      <div class="group">
        <div class="row stack"><div class="rt"><b>Finish a page early or late</b></div><div class="seg wide" role="group">${[['spread', 'Share it out'], ['carry', 'Next page'], ['fixed', 'Keep plan']].map(([k, l]) => `<button data-reb="${k}" aria-pressed="${st.rebalance === k}">${l}</button>`).join('')}</div>
        <small class="muted">${{ spread: 'Time you save or lose is shared over every page still to go.', carry: 'The difference rolls onto the very next page.', fixed: 'Each page keeps its time; you just see if you’re ahead or behind.' }[st.rebalance]}</small></div>
      </div>`;
    root.querySelectorAll('input[type=range]').forEach(fillTrack);
    summary();
  };
  const summary = () => {
    const u = units();
    const n = u.length;
    let tot = st.mode === 'total' ? st.total : st.mode === 'per' ? st.per * n : st.mode === 'by' ? (st.by - Date.now()) / 1000 : suggest();
    const p = new Pacer(u.length ? u : [{ doc: 'x', i: 0, words: 0 }], cfg());
    const brk = p.plannedBreaks(n, st.mode === 'by' ? tot : tot, st.docs.length);
    if (st.mode === 'by') tot = Math.max(60, tot - brk);
    const end = st.mode === 'by' ? st.by : Date.now() + (tot + brk) * 1000;
    $('#suSum', s.el).innerHTML = n ? `<b>${n} pages · ${n ? clock(tot / n, { hours: false }) : '0'} a page</b><br>${st.mode === 'open' ? 'open-ended' : `done around ${timeOfDay(end)}`}${brk ? ` · ${human(brk)} of breaks` : ''}` : 'Nothing to read';
    s.el.querySelector('[data-start]').disabled = !n || (st.mode === 'by' && st.by - Date.now() < 120000);
  };
  const cfg = () => {
    const mc = methodCfg(mvals(), {});
    return { mode: st.mode, total: st.mode === 'open' ? suggest() : st.total, per: st.per, by: st.mode === 'by' ? st.by : null, profile: st.profile, strength: st.strength, rebalance: st.rebalance, ...mc, interrupt: mc.flowtime ? false : st.interrupt, recall: st.recall, breakDocs: st.breakDocs };
  };
  root.addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    const ds = b.dataset;
    if (ds.mode) st.mode = ds.mode;
    else if (ds.tadj) st.total = Math.max(300, Math.min(16 * 3600, st.total + Number(ds.tadj)));
    else if (ds.tset) st.total = Number(ds.tset);
    else if (ds.padj) st.per = Math.max(5, Math.min(1800, st.per + Number(ds.padj) * (st.per >= 120 ? 6 : st.per >= 60 ? 3 : 1)));
    else if (ds.pset) st.per = Number(ds.pset);
    else if (ds.byin) { st.by = Date.now() + Number(ds.byin) * 60000; }
    else if (ds.method) { st.method = ds.method; st.m = {}; }
    else if (ds.mv) {
      const [k, dv] = ds.mv.split(':');
      const m = mvals();
      const lim = { focus: [5, 180], short: [1, 60], long: [0, 60], every: [2, 8], pages: [1, 100] }[k];
      st.m[k] = Math.max(lim[0], Math.min(lim[1], (m[k] || 0) + Number(dv)));
      if (k === 'long' && st.m.long && !m.every) st.m.every = 4;
      // edits to a preset become your Custom rhythm
      prefs.customMethod = { focus: mvals().focus || 30, short: mvals().short || 5, long: mvals().long || 0, every: mvals().every || 4, blurb: undefined };
      if (st.method !== 'custom' && !method().pages && !method().flowtime) { const keep = { ...mvals() }; st.method = 'custom'; st.m = {}; prefs.customMethod = { focus: keep.focus, short: keep.short, long: keep.long, every: keep.every || 4 }; }
    }
    else if (ds.profile) st.profile = ds.profile;
    else if (ds.reb) st.rebalance = ds.reb;
    else if (ds.qdel) st.docs.splice(Number(ds.qdel), 1);
    else return;
    const y = s.body.scrollTop;
    render();
    s.body.scrollTop = y;
  });
  root.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.strength != null && t.hasAttribute('data-strength')) { st.strength = Number(t.value); fillTrack(t); t.closest('.row').querySelector('output').textContent = `${st.strength}%`; const pc = root.querySelector('.plan-chart'); if (pc) pc.outerHTML = planBars().split('<div class="plan-cap">')[0]; summary(); return; }
    if (t.id === 'byTime') { const [h, m] = t.value.split(':').map(Number); const d = new Date(); d.setHours(h, m, 0, 0); if (d.getTime() < Date.now() + 60000) d.setDate(d.getDate() + 1); st.by = d.getTime(); t.nextElementSibling.textContent = `${human((st.by - Date.now()) / 1000)} from now`; summary(); return; }
    if (t.dataset.k) { st[t.dataset.k] = t.checked; const y = s.body.scrollTop; render(); s.body.scrollTop = y; return; }
  });
  root.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.from != null || t.dataset.to != null) {
      const k = Number(t.dataset.from ?? t.dataset.to);
      const q = st.docs[k], d = docById(q.id);
      const v = Math.max(1, Math.min(d.units, Math.round(Number(t.value) || 1))) - 1;
      if (t.dataset.from != null) q.from = Math.min(v, q.to); else q.to = Math.max(v, q.from);
      const y = s.body.scrollTop; render(); s.body.scrollTop = y;
    }
  });
  // drag to reorder documents
  root.addEventListener('pointerdown', (e) => {
    const h = e.target.closest('.q-handle');
    if (!h) return;
    e.preventDefault();
    const row = h.closest('.qrow');
    const rows = $$('.qrow', root);
    const from = Number(row.dataset.q);
    const hgt = row.offsetHeight;
    let to = from;
    const y0 = e.clientY;
    row.classList.add('dragging');
    h.setPointerCapture(e.pointerId);
    const move = (ev) => {
      const dy = ev.clientY - y0;
      row.style.transform = `translateY(${dy}px)`;
      to = Math.max(0, Math.min(rows.length - 1, from + Math.round(dy / hgt)));
      rows.forEach((r, i) => { if (r === row) return; const shift = from < to && i > from && i <= to ? -hgt : from > to && i < from && i >= to ? hgt : 0; r.style.transition = 'transform .25s cubic-bezier(.2,.9,.3,1)'; r.style.transform = shift ? `translateY(${shift}px)` : ''; });
    };
    const up = () => {
      h.removeEventListener('pointermove', move); h.removeEventListener('pointerup', up); h.removeEventListener('pointercancel', up);
      const [it] = st.docs.splice(from, 1); st.docs.splice(to, 0, it);
      render();
    };
    h.addEventListener('pointermove', move); h.addEventListener('pointerup', up); h.addEventListener('pointercancel', up);
  });
  s.el.querySelector('[data-start]').onclick = () => {
    const u = units();
    if (!u.length) return;
    prefs.setup = { mode: st.mode, per: st.per, profile: st.profile, strength: st.strength, rebalance: st.rebalance, method: st.method, m: st.m, interrupt: st.interrupt, recall: st.recall, breakDocs: st.breakDocs };
    savePrefs();
    s.close();
    const pacer = new Pacer(u, cfg());
    startSession(u, pacer);
  };
  render();
}

// ================= the reader =================
const R = { viewer: null, units: [], pacer: null, k: 0, flags: new Set(), docs: new Map(), raf: 0, session: null, recall: [], lastText: 0, awayPaused: false, veil: null };
const reader = $('#reader');

Pacer.restore = (ses) => { const p = new Pacer(ses.units, ses.pacer.cfg, ses.pacer.s); return p; };

async function openFree(ids, fromCard) {
  const ds = ids.map(docById).filter(Boolean);
  if (!ds.length) return;
  const units = [];
  for (const d of ds) for (let i = 0; i < d.units; i++) units.push({ doc: d.id, i, words: d.words[i] || 0, kind: d.kind });
  const start = ds.length === 1 ? Math.min(ds[0].pos || 0, ds[0].units - 1) : 0;
  await openReader(units, null, start, fromCard);
}
async function startSession(units, pacer) {
  await store.del('session');
  R.recall = [];
  const s = { id: Date.now(), units, pacer: pacer.toJSON(), flags: [], recall: [], created: Date.now() };
  R.session = s;
  await openReader(units, pacer, 0);
  pacer.start();
  saveSession();
  toast(`${ICON.timer} ${pacer.cfg.mode === 'open' ? 'Pacing started' : `${human(pacer.snapshot().remaining)} for ${plural(units.length, 'page')}`}${pacer.cfg.method !== 'none' ? ` · ${esc(METHODS.find((m) => m.id === pacer.cfg.method)?.name || 'Custom')}` : ''}`, { ms: 2600 });
}
async function resumeSession() {
  const ses = await store.get('session');
  if (!ses) return;
  if (ses.units.some((u) => !docById(u.doc))) { toast('Some documents in that session were deleted'); await store.del('session'); renderHero(); return; }
  const p = Pacer.restore(ses);
  R.session = ses;
  R.recall = ses.recall || [];
  await openReader(ses.units, p, p.s.cur);
  (ses.flags || []).forEach((k) => { R.flags.add(k); R.viewer.setFlag(k, true); });
  p.resume();
  if (p.s.brk) showBreak();
}
async function saveSession() {
  if (!R.pacer || !R.session) return;
  R.session.pacer = R.pacer.toJSON();
  R.session.flags = [...R.flags];
  R.session.recall = R.recall;
  try { await store.set('session', R.session); } catch {}
}

async function openReader(units, pacer, start, fromEl) {
  closeMenu();
  R.units = units; R.pacer = pacer; R.flags = new Set();
  R.docs = new Map(units.map((u) => [u.doc, docById(u.doc)]));
  R.k = start;
  $('#library').inert = true;
  reader.hidden = false;
  reader.classList.remove('immersive');
  reader.classList.toggle('paged-mode', prefs.view !== 'scroll');
  // zoom in from the card that was tapped
  if (!reduced()) {
    const r = fromEl?.getBoundingClientRect?.();
    if (r) {
      reader.style.transformOrigin = `${r.left + r.width / 2}px ${r.top + r.height / 2}px`;
      reader.animate([{ transform: `scale(${Math.max(0.2, r.width / innerWidth)})`, opacity: 0, borderRadius: '30px' }, { transform: 'none', opacity: 1, borderRadius: '0' }], { duration: 480, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
    } else reader.animate([{ opacity: 0, transform: 'scale(1.03)' }, { opacity: 1, transform: 'none' }], { duration: 380, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
  }
  R.viewer?.destroy();
  const v = new Viewer($('#stage'), {
    docs: R.docs, units, prefs,
    onPage: (k, dir) => moved(k, dir, true),
    onTap: (e) => tapStage(e),
    onSwipe: (dir) => { const j = R.k + dir; if (j < 0 || j >= units.length) { if (dir > 0 && R.pacer) nextPage(); return false; } go(j); return true; },
  });
  v.onZoom = (z) => showZoom(z);
  R.viewer = v;
  await v.primeNotes();
  v.show(start, 0);
  renderTop();
  renderHud(true);
  cancelAnimationFrame(R.raf);
  loop();
  keepAwake(!!pacer);
  $('#stage').focus({ preventScroll: true });
  history.pushState({ reader: 1 }, '');
}
function exitReader(fromPop) {
  if (reader.hidden || R.exiting) return;
  R.exiting = true;
  cancelAnimationFrame(R.raf);
  if (R.pacer && !R.pacer.s.finished) { R.pacer.tick(); saveSession(); }
  closeVeil(true);
  $('.nudge')?.remove();
  const out = reader.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'scale(0.97)' }], { duration: reduced() ? 1 : 240, easing: 'ease-in' });
  out.onfinish = () => { reader.hidden = true; R.exiting = false; R.viewer?.destroy(); R.viewer = null; $('#stage').innerHTML = ''; $('#hud').innerHTML = ''; };
  $('#library').inert = false;
  keepAwake(false);
  R.pacer = null;
  loadLibrary();
  if (!fromPop && history.state?.reader) history.back();
}
addEventListener('popstate', () => { if (sheets.length) { sheets[sheets.length - 1].close(); history.pushState({ reader: 1 }, ''); return; } if (!reader.hidden) exitReader(true); });

function go(k) {
  if (k < 0 || k >= R.units.length || k === R.k) return;
  const dir = k > R.k ? 1 : -1;
  R.viewer.show(k, Math.abs(k - R.k) > 1 ? 0 : dir);
  moved(k, dir, false);
}
function moved(k, dir, fromScroll) {
  const from = R.k;
  R.k = k;
  if (R.pacer) {
    const r = R.pacer.go(k);
    if (r?.break) { chime('due'); buzz([30, 60, 30]); showBreak(); }
  }
  // remember what's been read and where you are
  const u = R.units[k], d = R.docs.get(u.doc);
  const fu = R.units[from];
  if (fu && dir > 0) markRead(fu);
  if (d) { d.pos = u.i; d.opened = Date.now(); clearTimeout(R._dp); R._dp = setTimeout(() => store.putDoc(d), 600); }
  renderTop();
  renderHud();
  saveSession();
}
function markRead(u) {
  const d = R.docs.get(u.doc);
  if (!d) return;
  d.read ||= [];
  if (!d.read.includes(u.i)) { d.read.push(u.i); store.putDoc(d); }
}
function nextPage() {
  if (R.viewer.mode === 'paged' && R.k < R.units.length - 1) { go(R.k + 1); return; }
  if (R.k < R.units.length - 1) { go(R.k + 1); return; }
  // the very end
  markRead(R.units[R.k]);
  if (!R.pacer) { toast(`${ICON.check} That’s the end`, { actions: [{ label: 'Library', run: () => exitReader() }] }); return; }
  const r = R.pacer.next();
  if (r.finished) finishSession();
  else if (r.unread != null) {
    const n = R.pacer.left();
    toast(`${plural(n, 'page')} still unread`, { ms: 6000, actions: [{ label: 'Go there', primary: true, run: () => go(r.unread) }, { label: 'Finish', run: () => { R.pacer.finish(); finishSession(); } }] });
  }
}
const prevPage = () => go(R.k - 1);
function tapStage(e) {
  if (R.veil) return;
  const x = e.clientX / innerWidth;
  if (prefs.tapEdges && R.viewer.mode === 'paged' && e.pointerType !== 'mouse') {
    if (x < 0.22) { prevPage(); return; }
    if (x > 0.78) { nextPage(); return; }
  }
  if (e.pointerType === 'mouse') return;
  reader.classList.toggle('immersive');
}
function renderTop() {
  const u = R.units[R.k], d = R.docs.get(u.doc);
  $('#rdName').textContent = d.name;
  const docIds = [...R.docs.keys()];
  const head = d.heads?.[u.i];
  $('#rdSub').textContent = `${d.kind === 'slides' ? 'Slide' : 'Page'} ${u.i + 1} of ${d.units}${docIds.length > 1 ? ` · document ${docIds.indexOf(u.doc) + 1} of ${docIds.length}` : ''}${head && head !== d.name ? ` · ${head}` : ''}`;
  const fb = $('#flagBtn');
  fb.classList.toggle('on', R.flags.has(R.k));
  // progress, a segment per document
  const prog = $('#rdProg');
  const segs = docIds.map((id) => {
    const ks = R.units.map((x, j) => (x.doc === id ? j : -1)).filter((j) => j >= 0);
    const done = R.pacer ? ks.filter((j) => R.pacer.s.done[j]).length : ks.filter((j) => j < R.k).length;
    return { id, n: ks.length, p: (done / ks.length) * 100 };
  });
  if (prog.childElementCount !== segs.length) prog.innerHTML = segs.map((s) => `<span class="seg-doc" style="--n:${s.n}"><i></i></span>`).join('');
  segs.forEach((s, i) => { const el = prog.children[i]; el.style.setProperty('--p', `${s.p}%`); el.classList.toggle('cur', s.id === u.doc); });
}

// ---- the pacer bar ----
const RING_R = 23, RING_C = 2 * Math.PI * RING_R, FO_R = 28.5, FO_C = 2 * Math.PI * FO_R;
function renderHud(fresh) {
  const hud = $('#hud');
  hud.classList.toggle('compact', !!prefs.hudCompact);
  if (!R.pacer) {
    const u = R.units[R.k], d = R.docs.get(u.doc);
    const left = R.units.length - R.k - 1;
    hud.classList.add('free-hud');
    hud.innerHTML = `<div class="hud-stats"><div class="hs hs-main"><b>${R.k + 1} / ${R.units.length}</b><small>${left ? `${plural(left, d.kind === 'slides' ? 'slide' : 'page')} to go` : 'last page'}</small></div></div>
      <div class="hud-ctrl"><button class="icon-btn prev" data-act="prev" aria-label="Previous">${ICON.prev}</button><button class="btn primary" data-act="paceThis">${ICON.timer}<span>Pace it</span></button><button class="icon-btn" data-act="next" aria-label="Next">${ICON.next}</button></div>`;
    if (fresh) { hud.classList.remove('in'); void hud.offsetWidth; hud.classList.add('in'); }
    return;
  }
  hud.classList.remove('free-hud');
  if (fresh || !hud.querySelector('.ring')) {
    hud.innerHTML = `<div class="ring" data-act="ring" title="Pause or resume (P)"><svg viewBox="0 0 64 64"><circle class="tr" cx="32" cy="32" r="${RING_R}"/><circle class="ar" cx="32" cy="32" r="${RING_R}" stroke-dasharray="${RING_C}"/><circle class="fo" cx="32" cy="32" r="${FO_R}" stroke-dasharray="${FO_C}" stroke-dashoffset="${FO_C}"/></svg><b></b></div>
      <div class="hud-stats"><button class="hs hs-main" data-act="stats"><b></b><small></small></button><button class="hs hs-two" data-act="overview"><b></b><small></small></button><button class="hs hs-third" data-act="stats"><b></b><small></small></button></div>
      <div class="hud-ctrl"><button class="icon-btn prev" data-act="prev" aria-label="Previous page">${ICON.prev}</button><button class="icon-btn play" data-act="pause" aria-label="Pause"></button><button class="icon-btn" data-act="next" aria-label="Next page">${ICON.next}</button></div>`;
    hud.classList.remove('in'); void hud.offsetWidth; hud.classList.add('in');
    R._playState = null;
  }
  paintHud(true);
}
function paintHud(text) {
  const p = R.pacer;
  if (!p) return;
  const sn = p.snapshot();
  const hud = $('#hud');
  const ring = hud.querySelector('.ring');
  if (!ring) return;
  // page ring: drains as the page's time runs out
  const ar = ring.querySelector('.ar');
  let frac, label, cls = '';
  if (sn.review) { frac = 1; label = clock(sn.pageSpent, { hours: false }); cls = 'review'; }
  else if (sn.pageLeft >= 0) { frac = sn.pageAlloc ? sn.pageLeft / sn.pageAlloc : 0; label = clock(sn.pageLeft, { hours: false }); cls = frac < 0.2 ? 'warn' : ''; }
  else { frac = Math.min(1, -sn.pageLeft / Math.max(10, sn.pageAlloc)); label = `+${clock(-sn.pageLeft, { hours: false })}`; cls = 'over'; }
  ar.style.strokeDashoffset = String(RING_C * (1 - Math.max(0, Math.min(1, frac))));
  ring.className = `ring ${cls}${sn.paused ? ' paused' : ''}`;
  // outer arc: the focus block
  const fo = ring.querySelector('.fo');
  if (sn.focusLeft != null) { const tot = p.cfg.breakEvery * 60 + (p.s.snooze || 0); fo.style.strokeDashoffset = String(FO_C * (1 - Math.max(0, Math.min(1, sn.focusLeft / tot)))); fo.style.opacity = 1; }
  else fo.style.opacity = 0;
  if (!text) return;
  ring.querySelector('b').textContent = sn.paused ? 'II' : label;
  const [a, b, c] = hud.querySelectorAll('.hs');
  const dots = p.cfg.longEvery ? `<span class="focus-dots">${Array.from({ length: p.cfg.longEvery }, (_, i) => `<i class="${i < sn.cycles % p.cfg.longEvery ? 'on' : ''}"></i>`).join('')}</span>` : '';
  if (sn.open) { a.querySelector('b').textContent = clock(sn.work); a.querySelector('small').innerHTML = `${sn.focusLeft != null ? `focus ${clock(Math.max(0, sn.focusLeft), { hours: false })}` : 'reading time'}${dots}`; }
  else {
    a.querySelector('b').innerHTML = sn.remaining >= 0 ? clock(sn.remaining) : `<span class="bad">${clock(sn.remaining)}</span>`;
    a.querySelector('small').innerHTML = sn.focusLeft != null ? `focus ${clock(Math.max(0, sn.focusLeft), { hours: false })}${dots}` : `${sn.remaining >= 0 ? 'left' : 'over'}${dots}`;
  }
  b.querySelector('b').textContent = sn.left ? `${sn.left} to go` : 'All read';
  b.querySelector('small').textContent = `${sn.review ? 'reviewing · ' : ''}${sn.docs.length > 1 ? `${plural(sn.docsLeft, 'doc')} · ` : ''}${sn.doneN} of ${sn.n} done`;
  const ahead = Math.round(sn.delta / 60);
  c.querySelector('b').textContent = timeOfDay(sn.predicted);
  c.querySelector('small').innerHTML = sn.trust < 0.15 || sn.open ? 'expected finish' : Math.abs(ahead) < 1 ? '<span class="good">right on pace</span>' : ahead > 0 ? `<span class="good">${ahead} min ahead</span>` : `<span class="${ahead < -5 ? 'bad' : 'warn'}">${-ahead} min behind</span>`;
  const play = hud.querySelector('.play');
  const st = sn.paused ? 'play' : 'pause';
  if (R._playState !== st) { play.innerHTML = ICON[st]; play.setAttribute('aria-label', sn.paused ? 'Resume' : 'Pause'); R._playState = st; }
}
function loop() {
  const tickOnce = () => {
    const p = R.pacer;
    if (p && !p.s.finished) {
      p.tick();
      const now = performance.now();
      const text = now - R.lastText > 220;
      if (text) R.lastText = now;
      paintHud(text);
      const sn = text ? p.snapshot() : null;
      if (sn) {
        if (sn.due && !R.veil) showDue();
        if (sn.brk && !R.veil) showBreak();
        if (R.veil?.kind === 'break') paintBreak(sn);
        // a gentle heads-up a minute before a focus block ends
        if (prefs.nudge && sn.focusLeft != null && p.cfg.interrupt && sn.focusLeft < 60 && sn.focusLeft > 0 && !R._nudged && !sn.paused) { R._nudged = true; nudge(`Break in a minute · wrap up your thought`); chime('soft'); }
        if (sn.focusLeft != null && sn.focusLeft > 60) R._nudged = false;
        if (sn.over && !p.s.overNotified) { p.s.overNotified = true; outOfTime(sn); }
        if (now - (R._saved || 0) > 4000) { R._saved = now; saveSession(); }
      }
    }
    R.raf = requestAnimationFrame(tickOnce);
  };
  R.raf = requestAnimationFrame(tickOnce);
}
function nudge(text) {
  $('.nudge')?.remove();
  const n = document.createElement('div');
  n.className = 'nudge glass-menu';
  n.textContent = text;
  reader.append(n);
  setTimeout(() => { n.classList.add('out'); setTimeout(() => n.remove(), 320); }, 5200);
}
function outOfTime(sn) {
  chime('end'); buzz([40, 80, 40]);
  toast(`⏰ Time’s up · ${plural(sn.left, 'page')} to go`, { ms: 12000, actions: [{ label: '+5 min', run: () => addTime(300) }, { label: '+15', run: () => addTime(900) }, { label: 'Keep going', primary: true }] });
}
function addTime(sec) { R.pacer?.addTime(sec); toast(`${ICON.plus} Added ${human(sec)}`, { ms: 1800 }); saveSession(); paintHud(true); }
function togglePause() {
  const p = R.pacer;
  if (!p) return;
  if (p.s.paused) { p.resume(); R.awayPaused = false; } else p.pause();
  buzz(8);
  const ring = $('#hud .ring');
  ring?.animate([{ transform: 'scale(1)' }, { transform: 'scale(0.86)' }, { transform: 'scale(1)' }], { duration: 360, easing: 'cubic-bezier(.2,.9,.3,1.4)' });
  paintHud(true);
  saveSession();
}

// ---- interruptions: the focus block is done ----
function confetti(colors) {
  if (reduced()) return;
  const c = document.createElement('div');
  c.className = 'confetti';
  const cols = colors || ['#e2552c', '#f5b400', '#2f9e44', '#2f54eb', '#d6457b', '#0e9aa7'];
  c.innerHTML = Array.from({ length: 46 }, () => {
    const a = Math.random() * Math.PI * 2, r = 160 + Math.random() * 260;
    return `<i style="--c:${cols[Math.floor(Math.random() * cols.length)]};--x:${Math.cos(a) * r}px;--y:${Math.sin(a) * r - 80}px;--r:${Math.random() * 720 - 360}deg;--d:${0.75 + Math.random() * 0.25}s"></i>`;
  }).join('');
  document.body.append(c);
  setTimeout(() => c.remove(), 2800);
}
function openVeil(kind, html) {
  closeVeil(true);
  const v = document.createElement('div');
  v.className = 'veil';
  v.innerHTML = `<div class="veil-card">${html}</div>`;
  document.body.append(v);
  R.veil = { kind, el: v };
  closeMenu();
  return v;
}
function closeVeil(fast) {
  const v = R.veil;
  if (!v) return;
  R.veil = null;
  if (fast) { v.el.remove(); return; }
  v.el.classList.add('out');
  setTimeout(() => v.el.remove(), 400);
}
function showDue() {
  const p = R.pacer, sn = p.snapshot();
  const due = sn.due;
  const pages = p.s.log.filter((e) => e.type === 'page' && e.t > (p.s.lastBreakAt || p.s.started || 0)).length;
  const focus = Math.round((p.cfg.breakEvery * 60 + (p.s.snooze || 0)) / 60);
  chime('due'); buzz([30, 70, 30, 70, 60]);
  const v = openVeil('due', `
    <div class="burst done" style="--bc:#e2552c"><span class="ripple"></span><span class="ripple"></span><span class="ripple"></span>
      <svg class="dial" viewBox="0 0 160 160"><circle class="tr" cx="80" cy="80" r="70"/><circle class="ar" cx="80" cy="80" r="70"/></svg>
      <span class="core">${ICON.check}</span></div>
    <h2>${due.long ? 'Set complete. Long break!' : 'Focus block done!'}</h2>
    <p>${focus} minutes of focus${pages ? `, ${plural(pages, 'page')} read` : ''}. ${due.long ? `That’s ${p.cfg.longEvery} blocks in a row. You’ve earned a proper rest.` : 'Step away for a moment; it helps it stick.'}</p>
    <div class="veil-actions">
      <button class="btn primary full" data-v="take">${ICON.coffee}<span>Take ${human(due.len)} break</span></button>
      <button class="btn" data-v="s5">+5 minutes</button><button class="btn" data-v="s10">+10 minutes</button>
      <button class="btn" data-v="page">Finish this page</button><button class="btn" data-v="skip">Skip break</button>
    </div>
    <div class="veil-more"><button class="chip" data-v="s2">+2</button><button class="chip" data-v="s15">+15</button><button class="chip" data-v="s25">+25 min</button></div>`);
  confetti();
  v.onclick = (e) => {
    const b = e.target.closest('[data-v]');
    if (!b) return;
    const a = b.dataset.v;
    if (a === 'take') { p.acceptDue(); closeVeil(true); showBreak(); }
    else if (a === 'page') { p.deferDue(); closeVeil(); toast('Break when you turn the page', { ms: 2200 }); }
    else if (a === 'skip') { p.skipDue(); closeVeil(); toast('Skipped. Next focus block started', { ms: 2200 }); }
    else if (a[0] === 's') { const m = Number(a.slice(1)); p.snoozeDue(m * 60); closeVeil(); toast(`${m} more minutes`, { ms: 1800 }); }
    saveSession();
  };
}

// ---- the break itself ----
const TIPS = ['Stand up and stretch your back and shoulders.', 'Look at something 20 feet away for 20 seconds to rest your eyes.', 'Drink some water.', 'Take a short walk, even just around the room.', 'Close your eyes and take a few slow breaths.', 'Roll your neck and loosen your jaw.', 'Step away from screens; your brain keeps consolidating while you rest.'];
function showBreak() {
  const p = R.pacer, sn = p.snapshot();
  if (!sn.brk) return;
  const since = (sn.brk.since || []).map((k) => R.units[k]).filter(Boolean);
  const heads = [...new Set(since.map((u) => R.docs.get(u.doc)?.heads?.[u.i]).filter((h) => h && h.length > 2))].slice(-6);
  const recall = p.cfg.recall && since.length && sn.brk.why !== 'manual';
  const C = 2 * Math.PI * 104;
  const v = openVeil('break', `
    <div class="breathe"><span class="orb"></span><span class="orb o2"></span><svg viewBox="0 0 220 220"><circle class="tr" cx="110" cy="110" r="104"/><circle class="ar" cx="110" cy="110" r="104" stroke-dasharray="${C}"/></svg>
      <div class="bt"><b id="bkLeft">${clock(sn.brk.left, { hours: false })}</b><span id="bkBreath">Breathe in</span></div></div>
    <h2 id="bkTitle">${sn.brk.long ? 'Long break' : sn.brk.why === 'doc' ? 'One document down' : 'Break time'}</h2>
    <p class="tip">${esc(TIPS[Math.floor(Math.random() * TIPS.length)])}</p>
    ${recall ? `<div class="recall"><small>Recall check</small><p>Without looking back, what were the key ideas from the last ${plural(since.length, 'page')}?</p>${heads.length ? `<ul>${heads.map((h) => `<li>${esc(h)}</li>`).join('')}</ul>` : ''}<textarea id="bkNote" placeholder="Jot down what you remember (saved with this session)"></textarea></div>` : ''}
    <div class="veil-actions">
      <button class="btn" data-b="m1">+1 minute</button><button class="btn" data-b="m5">+5 minutes</button>
      <button class="btn primary full" data-b="end" id="bkEnd">I’m ready, back to it</button>
    </div>`);
  R.veil.C = C;
  R.veil.startT = performance.now();
  v.onclick = (e) => {
    const b = e.target.closest('[data-b]');
    if (!b) return;
    if (b.dataset.b === 'm1') p.extendBreak(60);
    else if (b.dataset.b === 'm5') p.extendBreak(300);
    else if (b.dataset.b === 'end') endBreakNow();
    paintBreak(p.snapshot());
  };
}
function paintBreak(sn) {
  const v = R.veil?.el;
  if (!v || !sn.brk) { if (R.veil?.kind === 'break' && !sn.brk) closeVeil(); return; }
  const left = sn.brk.left;
  $('#bkLeft', v).textContent = left > 0 ? clock(left, { hours: false }) : '0:00';
  v.querySelector('.breathe .ar').style.strokeDashoffset = String(R.veil.C * (1 - Math.max(0, left) / sn.brk.len));
  // 10 s breathing: 4 in, 1 hold, 5 out (matches the orb)
  const ph = ((performance.now() - R.veil.startT) / 1000) % 10;
  const word = ph < 4 ? 'Breathe in' : ph < 5 ? 'Hold' : 'Breathe out';
  const bw = $('#bkBreath', v); if (bw.textContent !== word && left > 0) bw.textContent = word;
  if (left <= 0 && !R.veil.ready) {
    R.veil.ready = true;
    chime('end'); buzz([40, 60, 40]);
    $('#bkTitle', v).textContent = 'Break’s over';
    bw.textContent = 'Ready when you are';
    const end = $('#bkEnd', v); end.animate([{ transform: 'scale(1)' }, { transform: 'scale(1.05)' }, { transform: 'scale(1)' }], { duration: 700, iterations: 3 });
  }
}
function endBreakNow() {
  const p = R.pacer;
  const note = $('#bkNote')?.value.trim();
  if (note) R.recall.push({ t: Date.now(), note, pages: p.s.brk?.since || [] });
  p.endBreak();
  closeVeil();
  toast('Back to it. Focus block started', { ms: 1800 });
  saveSession();
}

// ---- finished ----
async function finishSession() {
  const p = R.pacer;
  p.finish();
  const st = p.stats();
  chime('finish'); buzz([30, 50, 30, 50, 90]);
  await learnSpeed(p);
  await store.del('session');
  const sn = p.snapshot();
  const delta = Math.round((p.cfg.mode === 'by' ? (p.cfg.by - Date.now()) : (p.s.budget + p.s.added - p.s.work) * 1000) / 60000);
  const v = openVeil('finish', `
    <div class="burst done fin"><span class="ripple"></span><span class="ripple"></span><span class="ripple"></span>
      <svg class="dial" viewBox="0 0 160 160"><circle class="tr" cx="80" cy="80" r="70"/><circle class="ar" cx="80" cy="80" r="70"/></svg><span class="core">${ICON.check}</span></div>
    <h2>All done!</h2>
    <p>${plural(sn.n, 'page')} in ${human(st.work)}${st.breaks ? ` with ${plural(st.breaks, 'break')}` : ''}.${p.cfg.mode !== 'open' ? ` ${delta >= 1 ? `${delta} min ahead of plan.` : delta <= -1 ? `${-delta} min over plan.` : 'Right on time.'}` : ''}${st.wpm ? ` You read at about ${st.wpm} words a minute.` : ''}</p>
    <div class="veil-actions"><button class="btn" data-f="stats">${ICON.chart}<span>Session stats</span></button><button class="btn primary" data-f="lib">Library</button></div>`);
  confetti(['#2f9e44', '#5dd17c', '#f5b400', '#2f54eb', '#e2552c']);
  v.onclick = (e) => {
    const b = e.target.closest('[data-f]');
    if (!b) return;
    if (b.dataset.f === 'lib') { closeVeil(); exitReader(); }
    else { closeVeil(); openStats(p); }
  };
}
async function endSessionAsk() {
  closeMenu();
  const p = R.pacer;
  if (!p) return;
  const s = openSheet({ title: 'End this session?', body: `<p class="muted" style="margin:0 6px 8px">${plural(p.left(), 'page')} still to go. Your time so far is kept in your history.</p>`, foot: '<span class="grow"></span><button class="btn" data-x="keep">Keep going</button><button class="btn primary" data-x="end">End session</button>' });
  s.el.querySelector('[data-x="keep"]').onclick = () => s.close();
  s.el.querySelector('[data-x="end"]').onclick = async () => { s.close(); p.finish(); await learnSpeed(p); await store.del('session'); R.session = null; openStats(p, true); };
}

// ---- overview of every page ----
function openOverview() {
  closeMenu();
  const p = R.pacer;
  const groups = [];
  R.units.forEach((u, k) => { let g = groups[groups.length - 1]; if (!g || g.doc !== u.doc) { g = { doc: u.doc, ks: [] }; groups.push(g); } g.ks.push(k); });
  const body = groups.map((g) => {
    const d = R.docs.get(g.doc);
    const done = p ? g.ks.filter((k) => p.s.done[k]).length : 0;
    return `<div class="ov-doc"><b>${esc(d.name)}</b><small>${p ? `${done} of ${g.ks.length} read` : `${g.ks.length} ${unitWord(d)}`}</small></div>
      <div class="ov-grid">${g.ks.map((k) => {
        const u = R.units[k];
        const spent = p ? p.s.spent[k] : 0;
        return `<button class="ov-cell${k === R.k ? ' cur' : ''}${p?.s.done[k] ? ' done' : ''}${R.flags.has(k) ? ' flag' : ''}" data-k="${k}"><span class="ov-art" style="--ar:${d.kind === 'flow' ? 0.77 : u.ar || d.aspect || 0.77}"></span><small>${u.i + 1}${spent > 1 ? ` · <b>${clock(spent, { hours: false })}</b>` : ''}</small></button>`;
      }).join('')}</div>`;
  }).join('');
  const flagged = [...R.flags].sort((a, b) => a - b);
  const s = openSheet({ title: 'Pages', wide: true, body: `${flagged.length ? `<div class="group-t">Flagged</div><div class="chips" style="justify-content:flex-start;padding:0 2px 10px">${flagged.map((k) => `<button class="chip accent" data-k="${k}">★ ${esc(R.docs.get(R.units[k].doc).name.slice(0, 24))} · ${R.units[k].i + 1}</button>`).join('')}</div>` : ''}${body}` });
  s.body.onclick = (e) => { const c = e.target.closest('[data-k]'); if (c) { s.close(); go(Number(c.dataset.k)); } };
  const io = new IntersectionObserver((ents) => {
    for (const en of ents) {
      if (!en.isIntersecting) continue;
      io.unobserve(en.target);
      const k = Number(en.target.closest('[data-k]').dataset.k);
      R.viewer?.thumb(k, 110).then(({ el }) => { en.target.append(el); el.animate?.([{ opacity: 0 }, { opacity: 1 }], { duration: 200 }); }).catch(() => {});
    }
  }, { root: s.body, rootMargin: '200px' });
  $$('.ov-art', s.body).forEach((a) => io.observe(a));
  requestAnimationFrame(() => s.body.querySelector('.ov-cell.cur')?.scrollIntoView({ block: 'center' }));
}

// ---- stats ----
function openStats(p = R.pacer, ended = false) {
  closeMenu();
  if (!p) return;
  const st = p.stats();
  const pages = st.pages.filter((x) => x.done || x.sec > 2);
  const max = Math.max(1, ...pages.map((x) => Math.max(x.sec, x.planned || 0)));
  const recall = (R.session?.recall || R.recall || []);
  const name = (j) => { const u = p.units[j]; const d = docById(u.doc); return `${esc(d?.name?.slice(0, 30) || '')} · ${u.i + 1}`; };
  const body = `
    <div class="stat-grid">
      <div class="stat"><b>${human(st.work)}</b><span>reading time</span></div>
      <div class="stat"><b>${pages.filter((x) => x.done).length}/${p.units.length}</b><span>pages done</span></div>
      <div class="stat"><b>${st.avg ? clock(st.avg, { hours: false }) : '–'}</b><span>average a page</span></div>
      <div class="stat"><b>${st.wpm ? st.wpm : '–'}</b><span>words a minute</span></div>
      <div class="stat"><b>${st.breaks}</b><span>${st.breaks === 1 ? 'break' : 'breaks'} · ${human(st.breakTotal)}</span></div>
      <div class="stat"><b>${p.s.cycles || 0}</b><span>focus blocks</span></div>
    </div>
    ${pages.length ? `<div class="group"><div class="bars">${pages.map((x, k) => `<span class="b" title="${name(x.j).replace(/"/g, '')}: ${clock(x.sec, { hours: false })}${x.planned ? ` of ${clock(x.planned, { hours: false })}` : ''}"><i class="${x.planned && x.sec > x.planned * 1.05 ? 'over' : ''}" style="--h:${(x.sec / max) * 100}%;--k:${Math.min(k, 80)}"></i>${x.planned ? `<u style="--p:${(x.planned / max) * 100}%"></u>` : ''}</span>`).join('')}</div>
      <div class="bars-cap"><span><i></i>time on page</span><span><i class="o"></i>over its time</span><span><i class="p"></i>planned</span></div></div>` : ''}
    ${st.fastest ? `<div class="group"><div class="row"><div class="rt"><b>Quickest</b><small>${name(st.fastest.j)}</small></div><output>${clock(st.fastest.sec, { hours: false })}</output></div><div class="row"><div class="rt"><b>Longest</b><small>${name(st.slowest.j)}</small></div><output>${clock(st.slowest.sec, { hours: false })}</output></div></div>` : ''}
    ${recall.length ? `<div class="group-t">What you recalled</div><div class="group">${recall.map((r) => `<div class="row"><div class="rt"><small>${timeOfDay(r.t)}</small><b style="font-weight:500;white-space:pre-wrap">${esc(r.note)}</b></div></div>`).join('')}</div>` : ''}
    <div class="group-t">Log</div>
    <div class="group logl">${p.s.log.filter((e) => ['page', 'break', 'add', 'snooze', 'skip', 'pause', 'resume'].includes(e.type)).slice(-80).reverse().map((e) => `<div class="row"><span class="muted" style="width:70px">${timeOfDay(e.t)}</span><span class="rt">${e.type === 'page' ? `Read ${name(e.i)}` : e.type === 'break' ? `${e.why === 'manual' ? 'Took' : 'Started'} a ${human(e.len)} break` : e.type === 'add' ? `Added ${human(e.sec)}` : e.type === 'snooze' ? `Asked for ${human(e.sec)} more` : e.type === 'skip' ? 'Skipped a break' : e.type === 'pause' ? 'Paused' : 'Resumed'}</span>${e.type === 'page' ? `<output>${clock(e.sec, { hours: false })}</output>` : ''}</div>`).join('') || '<div class="row muted">Nothing yet</div>'}</div>`;
  const s = openSheet({ title: ended ? 'Session summary' : 'This session', wide: true, body, foot: ended ? '<span class="grow"></span><button class="btn primary" data-x="lib">Back to library</button>' : `<span class="grow"></span><button class="btn danger" data-x="end">End session</button>`, onClose: () => { if (ended && !reader.hidden) exitReader(); } });
  s.el.querySelector('[data-x]')?.addEventListener('click', (e) => { const x = e.currentTarget.dataset.x; s.close(); if (x === 'end') endSessionAsk(); });
}

// ---- more ----
function moreMenu(btn) {
  const p = R.pacer;
  openMenu(btn, [
    ...(p ? [
      { label: p.s.paused ? 'Resume' : 'Pause', icon: p.s.paused ? 'play' : 'pause', key: 'P', run: togglePause },
      { label: 'Take a break now', icon: 'coffee', key: 'B', run: () => { p.takeBreakNow(); showBreak(); } },
      { label: 'Add 5 minutes', icon: 'plus', run: () => addTime(300) },
      { label: 'Session stats', icon: 'chart', key: 'S', run: () => openStats() },
      '-',
    ] : [{ label: 'Pace these pages', icon: 'timer', run: () => paceThis() }, '-']),
    { label: prefs.view === 'scroll' ? 'Page by page' : 'Continuous scroll', icon: 'view', key: 'V', run: toggleView },
    { label: prefs.fit === 'width' ? 'Fit whole page' : 'Fit to width', icon: 'zoom', run: () => { prefs.fit = prefs.fit === 'width' ? 'page' : 'width'; savePrefs(); R.viewer.relayout(); } },
    { label: prefs.hudCompact ? 'Full pacer bar' : 'Compact pacer bar', icon: 'timer', run: () => { prefs.hudCompact = !prefs.hudCompact; savePrefs(); renderHud(); } },
    ...([...R.docs.values()].some((d) => d.kind === 'slides') ? [{ label: prefs.notes ? 'Hide speaker notes' : 'Show speaker notes', icon: 'notes', run: () => { prefs.notes = !prefs.notes; savePrefs(); R.viewer.lru.clear(); const k = R.k; R.viewer.k = -1; R.viewer.show(k, 0); } }] : []),
    { label: 'Settings', icon: 'gear', run: openSettings },
    ...(p ? ['-', { label: 'End session', icon: 'stop', danger: true, run: endSessionAsk }] : []),
  ]);
}
function toggleView() { prefs.view = prefs.view === 'scroll' ? 'paged' : 'scroll'; savePrefs(); reader.classList.toggle('paged-mode', prefs.view !== 'scroll'); R.viewer.setMode(); toast(prefs.view === 'scroll' ? 'Continuous scroll' : 'Page by page', { ms: 1400 }); }
function paceThis() {
  const ids = [...R.docs.keys()];
  const u = R.units[R.k];
  exitReader();
  setTimeout(() => openSetup(ids, { fromReader: u }), 260);
}
function showZoom(z) {
  const pill = $('#zoomPill');
  pill.hidden = false;
  pill.innerHTML = `<button class="icon-btn" data-z="-1" aria-label="Zoom out">${ICON.minus}</button><output>${Math.round(z * 100)}%</output><button class="icon-btn" data-z="1" aria-label="Zoom in">${ICON.plus}</button><button class="icon-btn" data-z="0" aria-label="Reset zoom">${ICON.reset}</button>`;
  clearTimeout(R._zh); R._zh = setTimeout(() => { pill.hidden = true; }, 3500);
}
$('#zoomPill').onclick = (e) => { const b = e.target.closest('[data-z]'); if (!b) return; const d = Number(b.dataset.z); showZoom(R.viewer.setZoom(d === 0 ? 1 : R.viewer.zoom * (d > 0 ? 1.2 : 1 / 1.2))); };
function toggleFlag() {
  const on = !R.flags.has(R.k);
  on ? R.flags.add(R.k) : R.flags.delete(R.k);
  R.viewer.setFlag(R.k, on);
  const b = $('#flagBtn');
  b.classList.toggle('on', on);
  b.animate([{ transform: 'scale(1) rotate(0)' }, { transform: 'scale(1.35) rotate(-14deg)' }, { transform: 'scale(1)' }], { duration: 420, easing: 'cubic-bezier(.2,.9,.3,1.4)' });
  buzz(8);
  saveSession();
}

// ================= settings =================
function openSettings() {
  closeMenu();
  const seg = (k, opts) => `<div class="seg" role="group">${opts.map(([v, l]) => `<button data-set="${k}" data-v="${v}" aria-pressed="${String(prefs[k]) === String(v)}">${l}</button>`).join('')}</div>`;
  const render = async () => {
    const est = await store.estimate();
    return `
    <div class="group-t">Reading</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Layout</b></div>${seg('view', [['paged', 'Pages'], ['scroll', 'Scroll']])}</div>
      <div class="row"><div class="rt"><b>Fit</b></div>${seg('fit', [['page', 'Whole page'], ['width', 'Width']])}</div>
      <div class="row"><div class="rt"><b>Page turn</b></div>${seg('turn', [['slide', 'Slide'], ['stack', 'Stack'], ['fade', 'Fade'], ['none', 'None']])}</div>
      <div class="row"><div class="rt"><b>Tap the edges to turn pages</b><small>Tap the middle to hide the bars</small></div>${sw('tapEdges', prefs.tapEdges)}</div>
      <div class="row"><div class="rt"><b>Speaker notes under slides</b></div>${sw('notes', prefs.notes)}</div>
    </div>
    <div class="group-t">Text documents</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Font</b></div>${seg('flowFont', [['serif', 'Serif'], ['sans', 'Sans'], ['plain', 'Wide'], ['mono', 'Mono']])}</div>
      <div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>Size</b><output>${prefs.fontSize}px</output></div><input type="range" class="sl" min="13" max="28" value="${prefs.fontSize}" data-r="fontSize"></div>
      <div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>Line spacing</b><output>${prefs.lineH}</output></div><input type="range" class="sl" min="1.2" max="2.2" step="0.05" value="${prefs.lineH}" data-r="lineH"></div>
      <div class="row stack"><div class="rt" style="flex-direction:row;justify-content:space-between"><b>Page width</b><output>${prefs.flowW}px</output></div><input type="range" class="sl" min="480" max="1100" step="20" value="${prefs.flowW}" data-r="flowW"></div>
    </div>
    <div class="group-t">Pacer</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Pause when I leave</b><small>Switching apps or tabs pauses the clock</small></div>${sw('autoPause', prefs.autoPause)}</div>
      <div class="row"><div class="rt"><b>Heads-up before a break</b><small>A quiet note a minute before the focus block ends</small></div>${sw('nudge', prefs.nudge)}</div>
      <div class="row"><div class="rt"><b>Sounds</b></div>${sw('sound', prefs.sound)}</div>
      <div class="row"><div class="rt"><b>Vibration</b></div>${sw('haptics', prefs.haptics)}</div>
      <div class="row"><div class="rt"><b>Keep the screen on</b><small>While a session is running</small></div>${sw('wake', prefs.wake)}</div>
      <div class="row"><div class="rt"><b>Compact pacer bar</b></div>${sw('hudCompact', prefs.hudCompact)}</div>
    </div>
    <div class="group-t">Look</div>
    <div class="group"><div class="row"><div class="rt"><b>Theme, paper and glass</b><small>Colours, page textures, see-through and blur</small></div><button class="btn" data-x="themes">Open</button></div></div>
    <div class="group-t">Your data</div>
    <div class="group">
      <div class="row"><div class="rt"><b>Stored on this device</b><small>${docs.length} documents${est?.usage ? ` · ${(est.usage / 1048576).toFixed(est.usage > 1e8 ? 0 : 1)} MB used` : ''}</small></div></div>
      ${speed.wpm ? `<div class="row"><div class="rt"><b>Your study pace</b><small>${speed.wpm} words a minute${speed.slide ? ` · ${speed.slide}s a slide` : ''}, learnt from your sessions</small></div><button class="btn" data-x="forget">Forget</button></div>` : ''}
      <div class="row"><div class="rt"><b>Remove everything</b><small>Documents, sessions and history</small></div><button class="btn danger" data-x="wipe">Remove</button></div>
    </div>
    <p class="muted" style="font-size:12px;text-align:center;margin:14px 0 4px">Keys: ←/→ pages · Space scrolls then turns · P pause · B break · F flag · G pages · T theme · V layout · +/− zoom</p>`;
  };
  const s = openSheet({ title: 'Settings', body: '' });
  const draw = async () => { const y = s.body.scrollTop; s.body.innerHTML = await render(); s.body.querySelectorAll('input[type=range]').forEach(fillTrack); s.body.scrollTop = y; };
  draw();
  const apply = (k) => {
    savePrefs();
    applyFlowPrefs();
    if (!R.viewer) return;
    if (k === 'view') { reader.classList.toggle('paged-mode', prefs.view !== 'scroll'); R.viewer.setMode(); }
    else if (['fit', 'fontSize', 'flowW', 'notes'].includes(k)) { if (k === 'notes') { R.viewer.lru.clear(); const kk = R.k; R.viewer.k = -1; R.viewer.show(kk, 0); } else R.viewer.relayout(); }
    else if (k === 'hudCompact') renderHud();
  };
  s.body.addEventListener('click', async (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.set) { const v = b.dataset.v; prefs[b.dataset.set] = isNaN(Number(v)) ? v : Number(v); apply(b.dataset.set); draw(); }
    else if (b.dataset.x === 'themes') { s.close(true); openThemes(); }
    else if (b.dataset.x === 'forget') { speed = { wpm: null, slide: null }; await store.del('speed'); draw(); }
    else if (b.dataset.x === 'wipe') {
      if (b.dataset.sure) { for (const d of docs) await store.deleteDoc(d.id); await store.del('session'); await store.del('history'); await store.del('speed'); docs = []; selected.clear(); s.close(); if (!reader.hidden) exitReader(); renderLibrary(); toast('Everything removed'); }
      else { b.dataset.sure = 1; b.textContent = 'Tap again to remove'; }
    }
  });
  s.body.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.k) { prefs[t.dataset.k] = t.checked; apply(t.dataset.k); if (t.dataset.k === 'wake') keepAwake(!!R.pacer); }
    else if (t.dataset.r) { prefs[t.dataset.r] = Number(t.value); fillTrack(t); t.closest('.row').querySelector('output').textContent = t.dataset.r === 'lineH' ? t.value : `${t.value}px`; apply(t.dataset.r); }
  });
}

// ================= actions & keys =================
document.addEventListener('click', (e) => {
  const b = e.target.closest('[data-act]');
  if (!b || b.closest('.sheet')) return;
  const a = b.dataset.act;
  const acts = {
    import: () => filePick.click(), themes: openThemes, settings: openSettings,
    selectAll: () => { const all = docs.every((d) => selected.has(d.id)); docs.forEach((d) => (all ? selected.delete(d.id) : selected.add(d.id))); renderLibrary(); },
    clearSel: () => { selected.clear(); renderLibrary(); },
    readSel: () => openFree(sortedDocs().filter((d) => selected.has(d.id)).map((d) => d.id)),
    paceSel: () => openSetup(sortedDocs().filter((d) => selected.has(d.id)).map((d) => d.id)),
    resume: resumeSession,
    endSaved: async () => { const ses = await store.get('session'); if (ses) { const p = Pacer.restore(ses); p.finish(); await learnSpeed(p); } await store.del('session'); renderHero(); toast('Session ended'); },
    exit: () => exitReader(), prev: prevPage, next: nextPage, pause: togglePause, ring: togglePause,
    flag: toggleFlag, overview: openOverview, stats: () => openStats(), more: () => moreMenu(b), paceThis,
  };
  acts[a]?.();
});
$('#sortSeg').onclick = (e) => { const b = e.target.closest('[data-sort]'); if (!b) return; prefs.sort = b.dataset.sort; savePrefs(); renderLibrary(); };
document.addEventListener('keydown', (e) => {
  if (e.target.closest('input, textarea, select, [contenteditable]')) return;
  if (e.metaKey || e.ctrlKey || e.altKey) {
    if (!reader.hidden && (e.key === '=' || e.key === '+' || e.key === '-' || e.key === '0')) { e.preventDefault(); showZoom(R.viewer.setZoom(e.key === '0' ? 1 : R.viewer.zoom * (e.key === '-' ? 1 / 1.15 : 1.15))); }
    return;
  }
  if (e.key === 'Escape') {
    if (themeBarOpen()) { openThemes(); return; }
    if (menuEl) { closeMenu(); return; }
    if (sheets.length) { sheets[sheets.length - 1].close(); return; }
    if (R.veil) return;
    if (!reader.hidden) { if (reader.classList.contains('immersive')) reader.classList.remove('immersive'); else exitReader(); }
    return;
  }
  if (sheets.length) return;
  const k = e.key.toLowerCase();
  if (k === 't') { openThemes(); return; }
  if (reader.hidden) { if (k === 'i' || k === 'o') filePick.click(); return; }
  if (R.veil) { if (k === 'enter') R.veil.el.querySelector('.btn.primary')?.click(); return; }
  const acts = {
    arrowright: nextPage, arrowleft: prevPage, pagedown: () => R.viewer.scrollBy(1) || nextPage(), pageup: () => R.viewer.scrollBy(-1) || prevPage(),
    ' ': () => (e.shiftKey ? R.viewer.scrollBy(-1) || prevPage() : R.viewer.scrollBy(1) || nextPage()),
    j: nextPage, k: prevPage, p: togglePause, f: toggleFlag, g: openOverview, s: () => openStats(), v: toggleView,
    b: () => { if (R.pacer) { R.pacer.takeBreakNow(); showBreak(); } },
    home: () => go(0), end: () => go(R.units.length - 1),
    '+': () => showZoom(R.viewer.setZoom(R.viewer.zoom * 1.15)), '=': () => showZoom(R.viewer.setZoom(R.viewer.zoom * 1.15)), '-': () => showZoom(R.viewer.setZoom(R.viewer.zoom / 1.15)),
    h: () => reader.classList.toggle('immersive'),
  };
  if (k === 'arrowdown' || k === 'arrowup') { if (R.viewer.mode === 'scroll') return; if (!R.viewer.scrollBy(k === 'arrowdown' ? 1 : -1)) (k === 'arrowdown' ? nextPage : prevPage)(); e.preventDefault(); return; }
  if (acts[k]) { e.preventDefault(); acts[k](); }
});

// pause when you leave, and say so when you come back
document.addEventListener('visibilitychange', () => {
  const p = R.pacer;
  if (document.visibilityState === 'hidden') {
    if (p && !p.s.finished) { p.tick(); if (prefs.autoPause && !p.s.paused && !p.s.brk) { p.pause(); R.awayPaused = Date.now(); } saveSession(); }
  } else {
    if (p) keepAwake(true);
    if (p && R.awayPaused) {
      const away = (Date.now() - R.awayPaused) / 1000;
      R.awayPaused = false;
      if (away > 4) toast(`Paused while you were away (${human(away)})`, { ms: 8000, actions: [{ label: 'Resume', primary: true, run: () => { if (p.s.paused) togglePause(); } }] });
      else p.resume();
      paintHud(true);
    }
  }
});
addEventListener('pagehide', () => { if (R.pacer) { R.pacer.tick(); saveSession(); } });

// ================= start =================
applyFlowPrefs();
(async () => {
  await loadSpeed();
  await loadLibrary();
  // pick up an unfinished session straight away if the app was closed mid-read
  const ses = await store.get('session');
  if (ses && !ses.pacer?.s?.finished && Date.now() - (ses.pacer?.s?.last || 0) < 6 * 3600e3 && new URLSearchParams(location.search).get('resume') !== '0') {
    toast('Pick up where you left off?', { ms: 9000, actions: [{ label: 'Resume', primary: true, run: resumeSession }] });
  }
})();
if ('serviceWorker' in navigator && location.protocol !== 'file:' && !/localhost|127\.0\.0\.1/.test(location.hostname)) navigator.serviceWorker.register('sw.js').catch(() => {});
window.__sp = { R, prefs, store, docs: () => docs };
