// Reading themes: the app's colours, and real paper for the pages.
// A theme is a background, a page colour, an ink colour, an accent and a paper; everything else is derived.

export const THEMES = [
  { id: 'auto', name: 'System', auto: true },
  { id: 'daylight', name: 'Daylight', bg: '#f5f4f1', page: '#ffffff', ink: '#1b1b1d', accent: '#2f54eb', paper: 'smooth' },
  { id: 'paper', name: 'Paper', bg: '#ece8df', page: '#f9f7f1', ink: '#25231f', accent: '#3a5bd9', paper: 'fine' },
  { id: 'cream', name: 'Cream', bg: '#e9dfc6', page: '#f6eed9', ink: '#2f281d', accent: '#a0612a', paper: 'book' },
  { id: 'sepia', name: 'Sepia', bg: '#dccaa6', page: '#efe1c3', ink: '#4a3923', accent: '#9a5216', paper: 'natural' },
  { id: 'sage', name: 'Sage', bg: '#d7e2d0', page: '#e8f0e2', ink: '#22301d', accent: '#3d7a4b', paper: 'rice' },
  { id: 'mist', name: 'Mist', bg: '#dfe4ec', page: '#eff2f7', ink: '#1c2431', accent: '#3a62c3', paper: 'laid' },
  { id: 'blush', name: 'Blush', bg: '#ecdedb', page: '#f8efed', ink: '#33231f', accent: '#b04757', paper: 'handmade' },
  { id: 'dusk', name: 'Dusk', bg: '#303237', page: '#3b3e44', ink: '#e7e5e0', accent: '#8ea2ff', dark: true, paper: 'felt' },
  { id: 'night', name: 'Night', bg: '#141517', page: '#1e1f23', ink: '#ecebe8', accent: '#7d95ff', dark: true, paper: 'card' },
  { id: 'ember', name: 'Night sepia', bg: '#1b1611', page: '#271f18', ink: '#e9d9bd', accent: '#e0a15a', dark: true, paper: 'natural' },
  { id: 'navy', name: 'Midnight', bg: '#0d1421', page: '#162033', ink: '#d5deeb', accent: '#7aa7ff', dark: true, paper: 'linen' },
  { id: 'forest', name: 'Forest', bg: '#111914', page: '#1a251e', ink: '#d5e3d7', accent: '#78c08f', dark: true, paper: 'felt' },
  { id: 'black', name: 'Black', bg: '#000000', page: '#0c0c0c', ink: '#d4d4d4', accent: '#8ea2ff', dark: true, paper: 'smooth' },
  { id: 'custom', name: 'Custom', custom: true },
];
// Real paper: scanned textures (Subtle Patterns, CC BY-SA 3.0 — see vendor/paper/LICENSE.txt), evened out
// around mid-grey so they give any page colour its grain.
export const PAPERS = [
  { id: 'theme', name: 'Automatic' },
  { id: 'smooth', name: 'Smooth' },
  { id: 'fine', name: 'Fine' },
  { id: 'book', name: 'Book' },
  { id: 'laid', name: 'Laid' },
  { id: 'rice', name: 'Rice' },
  { id: 'handmade', name: 'Handmade' },
  { id: 'recycled', name: 'Recycled' },
  { id: 'natural', name: 'Natural' },
  { id: 'linen', name: 'Linen' },
  { id: 'felt', name: 'Felt' },
  { id: 'card', name: 'Cardstock' },
];
export const THEME_DEFAULTS = { rtheme: 'auto', tint: true, tintAmt: 85, pageLift: null, paper: 'theme', grain: 50, accent: null, custom: { h: 38, s: 32, b: 86 } };
// accent colours to pick from (null = the theme's own); lifted a little on dark themes so they stay readable
// Blur is set on a curve: most of the slider is the light end, where small changes show; the last stretch is
// the heavy frosting. 0 = none, BLUR_DEF = the designed look (about 22 px), 100 = 28 px.
export const BLUR_DEF = 89;
export const blurPx = (v) => 28 * Math.pow(Math.max(0, Math.min(100, v)) / 100, 2);
export const ACCENTS = ['#2f54eb', '#5e5ce6', '#9b51e0', '#d6457b', '#e0533d', '#d9820f', '#2f9e44', '#0e9aa7', '#5f6b7a'];
// absolute, because a url() inside a CSS variable resolves against the stylesheet that uses it
const paperUrl = (id) => (id && id !== 'smooth' ? new URL(`vendor/paper/${id}.webp`, document.baseURI).href : '');

// ---- colour helpers ----
const hex = (h) => { const n = parseInt(h.slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
const toHex = (c) => '#' + c.map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('');
const mix = (a, b, t) => { const x = hex(a), y = hex(b); return toHex(x.map((v, i) => v + (y[i] - v) * t)); };
const rgba = (a, al) => { const [r, g, b] = hex(a); return `rgba(${r}, ${g}, ${b}, ${al})`; };
const light = (a) => hex(a).reduce((s, v) => s + v, 0) / 765;
function hsl(h, s, l) {
  s /= 100; l /= 100;
  const k = (n) => (n + h / 30) % 12;
  const f = (n) => l - s * Math.min(l, 1 - l) * Math.max(-1, Math.min(k(n) - 3, 9 - k(n), 1));
  return toHex([f(0) * 255, f(8) * 255, f(4) * 255]);
}

// The custom theme, from three sliders: hue, colour (how tinted) and brightness.
export function customTheme({ h = 38, s = 32, b = 86 } = {}) {
  const L = 4 + b * 0.92;
  const dark = L < 50;
  return {
    id: 'custom', name: 'Custom', dark, paper: dark ? 'felt' : 'fine',
    bg: hsl(h, s * 0.5, L),
    page: hsl(h, s * 0.55, dark ? Math.min(40, L + 5) : Math.min(98.5, L + 5)),
    ink: dark ? hsl(h, s * 0.25, 88) : hsl(h, s * 0.35, 13),
    accent: s < 8 ? (dark ? '#8ea2ff' : '#2f54eb') : hsl(h, Math.min(70, 35 + s), dark ? 68 : 40),
  };
}

const systemDark = () => matchMedia('(prefers-color-scheme: dark)').matches;
export function resolveTheme(p) {
  const t = baseTheme(p);
  if (!p.accent) return t;
  return { ...t, accent: t.dark ? mix(p.accent, '#ffffff', 0.3) : p.accent };
}
function baseTheme(p) {
  const id = p.rtheme || 'auto';
  if (id === 'custom') return customTheme(p.custom || THEME_DEFAULTS.custom);
  if (id === 'auto') return THEMES.find((t) => t.id === (systemDark() ? 'night' : 'daylight'));
  return THEMES.find((t) => t.id === id && !t.auto && !t.custom) || THEMES[1];
}

// How far the page sits from the background, -100 (much darker) … +100 (much lighter). null = as designed.
const LIFT = 0.6;
export function themeLift(t) {
  const lb = light(t.bg), lp = light(t.page);
  const v = lp >= lb ? (lp - lb) / Math.max(0.001, 1 - lb) : -(lb - lp) / Math.max(0.001, lb);
  return Math.round(Math.max(-1, Math.min(1, v / LIFT)) * 100);
}
export function pageColour(t, lift) {
  if (lift == null) return t.page;
  return lift >= 0 ? mix(t.bg, '#ffffff', (lift / 100) * LIFT) : mix(t.bg, '#000000', (-lift / 100) * LIFT);
}

// All the colour tokens for a theme.
export function themeTokens(t, lift = null) {
  const page = pageColour(t, lift);
  const d = t.dark;
  return {
    '--bg': t.bg,
    '--panel': mix(t.bg, t.ink, 0.03),
    '--card': page,
    '--card-2': mix(page, t.bg, 0.45),
    '--ink': t.ink,
    '--ink-2': mix(t.ink, t.bg, 0.2),
    '--muted': mix(t.ink, t.bg, 0.46),
    '--line': mix(t.bg, t.ink, d ? 0.13 : 0.11),
    '--line-2': mix(t.bg, t.ink, d ? 0.2 : 0.17),
    '--accent': t.accent,
    '--accent-ink': d ? mix(t.accent, '#000000', 0.85) : '#ffffff',
    '--accent-soft': rgba(t.accent, d ? 0.16 : 0.11),
    '--seg-on': d ? mix(page, t.ink, 0.12) : page,
    '--doc-page': page,
    '--doc-ink': t.ink,
  };
}

const APPLIED = Object.keys(themeTokens(THEMES[1]));
const CACHE = 'study-pacer:theme';

export function applyReadingTheme(p) {
  const root = document.documentElement;
  const t = resolveTheme(p);
  const dark = !!t.dark;
  root.dataset.theme = dark ? 'dark' : 'light';
  const tokens = themeTokens(t, p.pageLift ?? null);
  for (const k of APPLIED) root.style.removeProperty(k);
  for (const [k, v] of Object.entries(tokens)) root.style.setProperty(k, v);
  // the pages take on the theme: light themes tint the white paper, dark themes turn the page dark
  const tint = p.tint !== false;
  root.classList.toggle('doc-tint', tint);
  root.classList.toggle('doc-dark', tint && dark);
  root.style.setProperty('--tint', String((p.tintAmt ?? 85) / 100));
  // …and get real paper grain
  const paper = (p.paper || 'theme') === 'theme' ? t.paper || 'smooth' : p.paper;
  const url = paperUrl(paper);
  root.classList.toggle('paper', !!url);
  root.style.setProperty('--paper-img', url ? `url("${url}")` : 'none');
  root.style.setProperty('--grain', String(((p.grain ?? 50) / 100) * (dark ? 0.7 : 1)));
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.content = t.bg;
  tabColours(root, { page: tokens['--doc-page'], card: tokens['--card'], dark, tint, amt: (p.tintAmt ?? 85) / 100, url, grain: ((p.grain ?? 50) / 100) * (dark ? 0.7 : 1) });
  const extra = [['--paper-img', url ? `url("${url}")` : 'none'], ['--grain', root.style.getPropertyValue('--grain')], ['--tint', root.style.getPropertyValue('--tint')]];
  const css = [...Object.entries(tokens), ...extra].map(([k, v]) => `${k}:${v}`).join(';');
  try { localStorage.setItem(CACHE, JSON.stringify({ css, dark, tint, paper: !!url, auto: (p.rtheme || 'auto') === 'auto' })); } catch {}
  return dark;
}

// The page tabs between results are drawn with exactly the layers their page has: the page's own
// colour, the theme's tint over a PDF, and the same paper grain. The grain layer on a page is the paper
// image with brightness, contrast and opacity filters and a blend; a background layer can't take filters,
// so the filtered texture is baked once into an image here and the tab blends it the same way.
const baked = new Map();
function bakeGrain(url, g, dark) {
  const key = url + '|' + g + '|' + dark;
  if (baked.has(key)) return baked.get(key);
  const job = new Promise((res) => {
    const img = new Image();
    img.onload = () => {
      try {
        const c = document.createElement('canvas'); c.width = img.naturalWidth; c.height = img.naturalHeight;
        const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(img, 0, 0);
        const im = x.getImageData(0, 0, c.width, c.height), d = im.data;
        for (let i = 0; i < d.length; i += 4) for (let j = 0; j < 3; j++) {
          const v = d[i + j] / 255;
          // light: multiply by brightness(1.8) at opacity g = multiply by 1 - g(1 - p')
          // dark: screen with brightness(0.42) contrast(1.7) at opacity g = screen with g·q
          const out = dark ? g * Math.max(0, Math.min(1, v * 0.714 - 0.35)) : 1 - g * (1 - Math.min(1, v * 1.8));
          d[i + j] = Math.round(out * 255);
        }
        x.putImageData(im, 0, 0);
        c.toBlob((bl) => res(bl ? { url: URL.createObjectURL(bl), w: c.width, h: c.height } : null));
      } catch { res(null); }
    };
    img.onerror = () => res(null);
    img.src = url;
  });
  baked.set(key, job);
  return job;
}
let tabJob = 0;
async function tabColours(root, o) {
  const job = ++tabJob;
  const set = (k, v) => (v == null ? root.style.removeProperty(k) : root.style.setProperty(k, v));
  const tex = o.url && o.grain > 0 ? await bakeGrain(o.url, Math.round(o.grain * 1000) / 1000, o.dark) : null;
  if (job !== tabJob) return;
  set('--tab-grain', tex ? `url("${tex.url}")` : null);
  set('--tab-gsize', tex ? `${tex.w}px ${tex.h}px` : null);
  set('--tab-gblend', tex ? (o.dark ? 'screen' : 'multiply') : null);
}

// ---- the theme bar: a row of cards along the bottom, so you can watch the results change ----
function mini(t, paperId) {
  const url = paperUrl(paperId || t.paper);
  return `<span class="tc-prev" style="--tb:${t.bg};--tp:${t.page};--ti:${t.ink};--ta:${t.accent}${url ? `;--tx:url('${url}')` : ''}"><span class="tc-page${url ? ' grain' : ''}${t.dark ? ' dk' : ''}"><i></i><i></i><i class="hl"></i><i></i></span></span>`;
}
function themeCard(t, on) {
  let prev;
  if (t.auto) prev = '<span class="tc-prev tc-auto"><span class="l"><i></i><i></i><i></i></span><span class="d"><i></i><i></i><i></i></span></span>';
  else if (t.custom) prev = '<span class="tc-prev tc-custom"><span class="tc-page"><i></i><i></i><i class="hl"></i><i></i></span></span>';
  else prev = mini(t);
  return `<button class="tcard${on ? ' on' : ''}" data-theme="${t.id}" aria-pressed="${on}">${prev}<span>${t.name}</span></button>`;
}
function paperCard(pp, on, t) {
  const id = pp.id === 'theme' ? t.paper || 'smooth' : pp.id;
  return `<button class="tcard${on ? ' on' : ''}" data-paper="${pp.id}" aria-pressed="${on}">${mini(t, id)}<span>${pp.name}</span></button>`;
}

// Glass: which surfaces the sliders below act on (any several at once)
const GLASS_PARTS = [['top', 'Top bar'], ['bottom', 'Pacer bar'], ['menus', 'Menus'], ['sheets', 'Sheets']];
export function openThemeBar(prefs, { save, glass, glassOf } = {}) {
  let bar = document.getElementById('themeBar');
  if (bar && !bar.hidden) { closeBar(bar, save); return; }
  if (!bar) {
    bar = document.createElement('section');
    bar.id = 'themeBar';
    bar.className = 'tbar';
    bar.setAttribute('aria-label', 'Reading theme');
    document.body.append(bar);
  }
  if (glass) bar._glass = glass;
  if (glassOf) bar._glassOf = glassOf;
  bar._parts ||= new Set(GLASS_PARTS.map(([k]) => k));
  let tab = bar._tab || 'colours';
  const commit = () => { clearTimeout(bar._t); bar._t = setTimeout(save, 250); };
  const apply = () => { applyReadingTheme(prefs); commit(); };
  const keepScroll = (fn) => { const r = bar.querySelector('.tbar-row'); const x = r?.scrollLeft; fn(); const r2 = bar.querySelector('.tbar-row'); if (r2 && x != null) r2.scrollLeft = x; };
  const render = (centre) => {
    const t = resolveTheme(prefs);
    const lift = prefs.pageLift ?? themeLift(t);
    const c = prefs.custom || THEME_DEFAULTS.custom;
    let body = '';
    const field = (label, value, input, ends, extra = '') => `<div class="tb-field"${extra}><div class="tb-lab"><b>${label}</b><output>${value}</output></div>${input}${ends ? `<div class="tb-ends"><span>${ends[0]}</span><span>${ends[1]}</span></div>` : ''}</div>`;
    const pct = (v) => `${Math.round(v)}%`;
    if (tab === 'colours') {
      body = `<div class="tbar-scroll"><button class="tbar-arrow l" data-scroll="-1" aria-label="Scroll left" hidden>${CHEV}</button><div class="tbar-row">${THEMES.map((x) => themeCard(x, (prefs.rtheme || 'auto') === x.id)).join('')}</div><button class="tbar-arrow r" data-scroll="1" aria-label="Scroll right" hidden>${CHEV}</button></div>
        ${prefs.rtheme === 'custom' ? `<div class="tbar-sliders">
          ${field('Hue', `${c.h}°`, `<input type="range" min="0" max="360" value="${c.h}" data-k="h" class="hue" style="--sat:${Math.max(20, c.s * 1.4)}%">`)}
          ${field('Colour', pct((c.s / 60) * 100), `<input type="range" min="0" max="60" value="${c.s}" data-k="s">`, ['Grey', 'Rich'])}
          ${field('Brightness', pct(c.b), `<input type="range" min="0" max="100" value="${c.b}" data-k="b" class="bright">`, ['Dark', 'Light'])}</div>`
        : `<p class="tbar-note">${(prefs.rtheme || 'auto') === 'auto' ? 'Light by day, dark by night: follows your device’s appearance setting.' : 'Pick <b>Custom</b> at the end of the row to mix your own.'}</p>`}`;
    } else if (tab === 'paper') {
      const eff = (prefs.paper || 'theme') === 'theme' ? t.paper || 'smooth' : prefs.paper;
      const effName = PAPERS.find((x) => x.id === eff)?.name || 'Smooth';
      const g = prefs.grain ?? 50;
      body = `<div class="tbar-scroll"><button class="tbar-arrow l" data-scroll="-1" aria-label="Scroll left" hidden>${CHEV}</button><div class="tbar-row">${PAPERS.map((pp) => paperCard(pp, (prefs.paper || 'theme') === pp.id, t)).join('')}</div><button class="tbar-arrow r" data-scroll="1" aria-label="Scroll right" hidden>${CHEV}</button></div>
        <div class="tbar-sliders">
          ${eff === 'smooth'
            ? `<p class="tbar-note">${(prefs.paper || 'theme') === 'theme' ? `Automatic picks <b>Smooth</b> for this theme: plain pages, no texture.` : 'Smooth pages have no texture.'} Pick another paper to add grain.</p>`
            : `${(prefs.paper || 'theme') === 'theme' ? `<p class="tbar-note">Automatic picks <b>${effName}</b> for this theme.</p>` : ''}${field('Grain', g ? pct(g) : 'Off', `<input type="range" min="0" max="100" value="${g}" data-grain>`, ['Faint', 'Strong'])}`}
        </div>
        <p class="tbar-credit">Paper textures from Subtle Patterns (CC BY-SA 3.0)</p>`;
    } else if (tab === 'glass') {
      // see-through and blur, for whichever surfaces are picked, live over the results behind
      const parts = [...bar._parts];
      const vals = parts.map((g) => bar._glassOf?.(g) || { t: prefs.glassT ?? 50, b: prefs.blurL ?? BLUR_DEF });
      const gt = vals[0]?.t ?? 50, gb = vals[0]?.b ?? BLUR_DEF;
      const mixed = (k) => vals.some((v) => v[k] !== vals[0][k]);
      const all = parts.length === GLASS_PARTS.length;
      const custom = Object.keys(prefs.glassG || {}).length > 0 || (prefs.glassT ?? 50) !== 50 || (prefs.blurL ?? BLUR_DEF) !== BLUR_DEF;
      body = `<div class="tbar-sliders">
        <div class="tb-parts" role="group" aria-label="Which surfaces"><button data-part="all" aria-pressed="${all}">All</button>${GLASS_PARTS.map(([k, l]) => `<button data-part="${k}" aria-pressed="${!all && bar._parts.has(k)}">${l}</button>`).join('')}</div>
        ${parts.length ? `${field('Transparency', mixed('t') ? 'Mixed' : gt === 50 ? 'As designed' : pct(gt), `<input type="range" min="0" max="100" value="${gt}" data-gt>`, ['Solid', 'Clear'])}
        ${field('Blur', mixed('b') ? 'Mixed' : gb === BLUR_DEF ? 'As designed' : `${blurPx(gb).toFixed(blurPx(gb) < 10 ? 1 : 0)} px`, `<input type="range" min="0" max="100" value="${gb}" data-gb>`, ['Sharp', 'Frosted'])}` : '<p class="tbar-note">Pick what to change.</p>'}
        <button class="tb-reset" data-greset ${custom ? '' : 'disabled'}>Reset all</button>
      </div>`;
    } else {
      const tint = prefs.tint !== false;
      const liftWord = (v) => (v === 0 ? 'Same' : v > 0 ? `Lighter ${v}` : `Darker ${-v}`);
      const changed = prefs.pageLift != null || prefs.accent || !tint || (prefs.tintAmt ?? 85) !== 85;
      body = `<div class="tbar-sliders">
        ${field('Page brightness', `${liftWord(lift)}${prefs.pageLift == null ? ' · theme’s' : ''}`, `<input type="range" min="-100" max="100" value="${lift}" data-lift>`, ['Darker than background', 'Lighter'])}
        <div class="tb-field"><div class="tb-lab"><b>Accent colour</b><output>${prefs.accent ? 'Custom' : 'Theme’s'}</output></div>
          <div class="tb-dots" role="group" aria-label="Accent colour"><button class="tb-dot auto${prefs.accent ? '' : ' on'}" data-accent="" style="--c:${baseTheme(prefs).accent}" aria-label="Theme’s accent"></button>${ACCENTS.map((a) => `<button class="tb-dot${prefs.accent === a ? ' on' : ''}" data-accent="${a}" style="--c:${t.dark ? mix(a, '#ffffff', 0.3) : a}" aria-label="Accent ${a}"></button>`).join('')}</div></div>
        <label class="th-switch"><span><b>Colour the pages</b><small>PDF pages, slides and images take on the theme’s colour and paper${t.dark ? '. Dark themes turn them dark.' : ''}</small></span><input type="checkbox" data-tint ${tint ? 'checked' : ''}><i></i></label>
        ${field('How much', pct(prefs.tintAmt ?? 85), `<input type="range" min="10" max="100" value="${prefs.tintAmt ?? 85}" data-amt>`, ['A hint', 'Fully'], tint ? '' : ' hidden')}
        <button class="tb-reset" data-reset ${changed ? '' : 'disabled'}>Reset these</button>
      </div>`;
    }
    // the panel fits what's in it, with the tabs fixed along the bottom; a change of tab glides to the new height
    const h0 = bar.offsetHeight;
    bar.innerHTML = `<div class="tbar-body">${body}</div><div class="tbar-foot">
        <div class="seg small tbar-tabs">${[['colours', 'Colours'], ['paper', 'Paper'], ['adjust', 'Adjust'], ['glass', 'Glass']].map(([k, l]) => `<button data-tab="${k}" aria-pressed="${tab === k}">${l}</button>`).join('')}</div>
        <button class="tbar-done" data-close>Done</button>
      </div>`;
    const h1 = bar.offsetHeight;
    if (h0 && Math.abs(h1 - h0) > 1 && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
      bar.animate([{ height: `${h0}px` }, { height: `${h1}px` }], { duration: 260, easing: 'cubic-bezier(0.32, 0.72, 0, 1)' });
    }
    bar.querySelectorAll('.tb-field input[type="range"]').forEach(fillTrack);
    glassDemo(bar, tab === 'glass' ? bar._parts : null);
    if (centre) bar.querySelector('.tcard.on')?.scrollIntoView({ block: 'nearest', inline: 'center' });
    scrollable(bar.querySelector('.tbar-row'));
  };
  bar.onclick = (e) => {
    const sc = e.target.closest('[data-scroll]');
    if (sc) { const r = bar.querySelector('.tbar-row'); r.scrollBy({ left: Number(sc.dataset.scroll) * r.clientWidth * 0.8, behavior: 'smooth' }); return; }
    const tb = e.target.closest('[data-tab]');
    if (tb) { tab = bar._tab = tb.dataset.tab; render(true); return; }
    if (e.target.closest('[data-close]')) { closeBar(bar, save); return; }
    const th = e.target.closest('.tcard[data-theme]');
    if (th) { prefs.rtheme = th.dataset.theme; apply(); keepScroll(() => render()); return; }
    const pp = e.target.closest('.tcard[data-paper]');
    if (pp) { prefs.paper = pp.dataset.paper; apply(); keepScroll(() => render()); return; }
    const ac = e.target.closest('[data-accent]');
    if (ac) { prefs.accent = ac.dataset.accent || null; apply(); render(); return; }
    const part = e.target.closest('[data-part]');
    if (part) {
      const k = part.dataset.part, every = GLASS_PARTS.map(([g]) => g);
      if (k === 'all') bar._parts = new Set(every);
      else if (bar._parts.size === every.length) bar._parts = new Set([k]); // from All, a tap picks just that one
      else { bar._parts.has(k) ? bar._parts.delete(k) : bar._parts.add(k); if (!bar._parts.size) bar._parts = new Set(every); }
      render();
      return;
    }
    if (e.target.closest('[data-greset]')) { prefs.glassT = 50; prefs.blurL = BLUR_DEF; prefs.glassG = {}; bar._glass?.(); commit(); render(); return; }
    if (e.target.closest('[data-reset]')) { prefs.pageLift = null; prefs.accent = null; prefs.tint = true; prefs.tintAmt = 85; apply(); render(); }
  };
  bar.oninput = (e) => {
    const t = e.target;
    if (t.dataset.k) { prefs.custom = { ...(prefs.custom || THEME_DEFAULTS.custom), [t.dataset.k]: Number(t.value) }; prefs.rtheme = 'custom'; }
    else if (t.hasAttribute('data-grain')) prefs.grain = Number(t.value);
    else if (t.hasAttribute('data-lift')) prefs.pageLift = Number(t.value);
    else if (t.hasAttribute('data-amt')) prefs.tintAmt = Number(t.value);
    else if (t.hasAttribute('data-gt') || t.hasAttribute('data-gb')) {
      const key = t.hasAttribute('data-gt') ? 't' : 'b';
      const v = Number(t.value);
      prefs.glassG ||= {};
      for (const g of bar._parts) prefs.glassG[g] = { ...(bar._glassOf?.(g) || {}), [key]: v };
      if (bar._parts.size === GLASS_PARTS.length) prefs[key === 't' ? 'glassT' : 'blurL'] = v; // All: the overall setting too
      bar._glass?.();
      commit();
      fillTrack(t);
      const out = t.closest('.tb-field')?.querySelector('output');
      if (out) out.textContent = key === 't' ? (v === 50 ? 'As designed' : `${v}%`) : v === BLUR_DEF ? 'As designed' : `${blurPx(v).toFixed(blurPx(v) < 10 ? 1 : 0)} px`;
      const rs = bar.querySelector('[data-greset]');
      if (rs) rs.disabled = false;
      return;
    }
    else if (t.hasAttribute('data-tint')) { prefs.tint = t.checked; bar.querySelector('[data-amt]').closest('.tb-field').hidden = !t.checked; }
    else return;
    if (t.type === 'range') fillTrack(t);
    // the value beside the label follows the slider
    const out = t.closest('.tb-field')?.querySelector('output');
    if (out) {
      const v = Number(t.value);
      out.textContent = t.dataset.k === 'h' ? `${v}°` : t.dataset.k === 's' ? `${Math.round((v / 60) * 100)}%`
        : t.hasAttribute('data-lift') ? (v === 0 ? 'Same' : v > 0 ? `Lighter ${v}` : `Darker ${-v}`)
        : t.hasAttribute('data-grain') ? (v ? `${v}%` : 'Off') : `${v}%`;
    }
    if (t.dataset.k === 'h' || t.dataset.k === 's') { const hue = bar.querySelector('input.hue'); if (hue) hue.style.setProperty('--sat', `${Math.max(20, (prefs.custom.s || 0) * 1.4)}%`); }
    const rs = bar.querySelector('[data-reset]');
    if (rs) rs.disabled = false;
    applyReadingTheme(prefs);
    commit();
  };
  render(true);
  bar.hidden = false;
  document.body.classList.add('tbar-open');
  bar.animate([{ transform: 'translateY(105%)' }, { transform: 'none' }], { duration: 300, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
}
const fillTrack = (r) => r.style.setProperty('--fillpct', `${((r.value - r.min) / (r.max - r.min)) * 100}%`);
const CHEV = '<svg viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>';
// A row of cards you can move with the mouse wheel, by dragging, or with the arrows at its ends.
function scrollable(row) {
  if (!row) return;
  const wrap = row.parentElement;
  const sync = () => {
    const max = row.scrollWidth - row.clientWidth;
    wrap.querySelector('.l').hidden = row.scrollLeft < 4;
    wrap.querySelector('.r').hidden = row.scrollLeft > max - 4;
  };
  row.addEventListener('scroll', sync, { passive: true });
  requestAnimationFrame(sync);
  row.addEventListener('wheel', (e) => {
    if (Math.abs(e.deltaX) > Math.abs(e.deltaY)) return; // a trackpad already scrolls sideways
    const max = row.scrollWidth - row.clientWidth;
    if (max <= 0) return;
    e.preventDefault();
    row.scrollLeft += e.deltaMode === 1 ? e.deltaY * 32 : e.deltaY;
  }, { passive: false });
  let drag = null;
  row.addEventListener('pointerdown', (e) => {
    if (e.pointerType !== 'mouse' || e.button !== 0) return;
    drag = { x: e.clientX, left: row.scrollLeft, moved: false };
  });
  row.addEventListener('pointermove', (e) => {
    if (!drag) return;
    const dx = e.clientX - drag.x;
    if (!drag.moved && Math.abs(dx) < 5) return;
    if (!drag.moved) { drag.moved = true; row.setPointerCapture(e.pointerId); row.classList.add('dragging'); }
    row.scrollLeft = drag.left - dx;
  });
  const end = () => {
    if (drag?.moved) {
      row.classList.remove('dragging');
      // the drag isn't a tap on the card under the mouse
      const eat = (ev) => { ev.stopPropagation(); ev.preventDefault(); };
      row.addEventListener('click', eat, { capture: true, once: true });
      setTimeout(() => row.removeEventListener('click', eat, { capture: true }), 60);
    }
    drag = null;
  };
  row.addEventListener('pointerup', end);
  row.addEventListener('pointercancel', end);
}
// A live preview of the surfaces picked in Glass that aren't already on screen (the bottom bar is under this panel,
// and menus and sheets only show when opened): stand-ins with the same glass, just above the panel, over whatever
// is behind. The top bar is the real one.
function glassDemo(bar, parts) {
  let demo = document.getElementById('glassDemo');
  const want = parts ? ['heads', 'bottom', 'menus'].filter((g) => parts.has(g)) : [];
  if (!want.length) {
    if (demo && !demo.classList.contains('leaving')) {
      demo.classList.add('leaving');
      demo.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(10px)' }], { duration: 180, easing: 'ease-in' }).onfinish = () => demo.remove();
    }
    return;
  }
  const fresh = !demo || demo.classList.contains('leaving');
  if (fresh) { demo?.remove(); demo = document.createElement('div'); demo.id = 'glassDemo'; demo.setAttribute('aria-hidden', 'true'); document.body.append(demo); }
  const have = new Set([...demo.children].map((c) => c.dataset.g));
  const make = {
    menus: '<div class="gd gd-menu" data-g="menus"><div>Jump to page…</div><div>Flag this page</div><div>Take a break now</div></div>',
    heads: '<div class="gd gd-head" data-g="heads"><b>Chapter 3 · Aerodynamics</b></div>',
    bottom: '<div class="gd gd-bar" data-g="bottom"><span class="gd-ring"></span><b>24:10 left</b><span>18 pages to go</span></div>',
  };
  for (const c of [...demo.children]) if (!want.includes(c.dataset.g)) {
    c.dataset.g = '';
    c.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'scale(0.96)' }], { duration: 160, easing: 'ease-in' }).onfinish = () => c.remove();
  }
  for (const g of want) {
    if (have.has(g)) continue;
    const t = document.createElement('template');
    t.innerHTML = make[g];
    const n = t.content.firstElementChild;
    // in order, top to bottom: a menu, a header, the bottom bar
    const order = ['menus', 'heads', 'bottom'];
    const after = [...demo.children].find((c) => order.indexOf(c.dataset.g) > order.indexOf(g));
    demo.insertBefore(n, after || null);
    n.animate([{ opacity: 0, transform: 'translateY(12px) scale(0.97)' }, { opacity: 1, transform: 'none' }], { duration: 300, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1)' });
  }
  // keep it just above the panel
  demo.style.bottom = `${bar.offsetHeight + 12}px`;
}
function closeBar(bar, save) {
  glassDemo(bar, null);
  save();
  bar.animate([{ transform: 'none' }, { transform: 'translateY(105%)' }], { duration: 220, easing: 'ease-in' }).onfinish = () => {
    bar.hidden = true;
    document.body.classList.remove('tbar-open');
  };
}
export const themeBarOpen = () => { const b = document.getElementById('themeBar'); return !!b && !b.hidden; };
