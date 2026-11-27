// The reader's page stage. Draws any unit (a PDF page, a slide, an image, a reflowed page) at the right size,
// one at a time with animated page turns ("paged") or as one continuous column ("scroll").
import { store } from './store.js';
import { pdfjs } from './importers.js';
import { renderSlide } from './pptx.js';

const DPR = () => Math.min(2.5, window.devicePixelRatio || 1);
const reduced = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

// ---- documents, loaded once and kept while the reader is open ----
const cache = new Map(); // docId -> Promise<{ kind, pdf?, deck?, urls?, chunks?, images? }>
export function loadDoc(doc) {
  if (cache.has(doc.id)) return cache.get(doc.id);
  const job = (async () => {
    if (doc.kind === 'pdf') {
      const lib = await pdfjs();
      const blob = await store.getFile(doc.id);
      const data = new Uint8Array(await blob.arrayBuffer());
      const pdf = await lib.getDocument({ data, isEvalSupported: false }).promise;
      return { kind: 'pdf', pdf, sizes: new Map() };
    }
    const c = await store.getContent(doc.id);
    if (doc.kind === 'slides') {
      const urls = {};
      for (const [k, b] of Object.entries(c.media || {})) if (b) urls[k] = URL.createObjectURL(b);
      return { kind: 'slides', deck: c, urls };
    }
    if (doc.kind === 'images') return { kind: 'images', urls: c.images.map((b) => URL.createObjectURL(b)) };
    return { kind: 'flow', chunks: c.chunks };
  })();
  cache.set(doc.id, job);
  job.catch(() => cache.delete(doc.id));
  return job;
}
export function releaseDocs() {
  for (const job of cache.values()) job.then((d) => {
    d.pdf?.destroy();
    for (const u of Object.values(d.urls || {})) URL.revokeObjectURL(u);
  }).catch(() => {});
  cache.clear();
}

async function pdfSize(d, n) {
  if (d.sizes.has(n)) return d.sizes.get(n);
  const page = await d.pdf.getPage(n);
  const vp = page.getViewport({ scale: 1 });
  const s = { w: vp.width, h: vp.height, page };
  d.sizes.set(n, s);
  return s;
}

export class Viewer {
  constructor(stage, { docs, units, prefs, onPage, onTap, onSwipe }) {
    this.stage = stage;
    this.docs = docs; // Map id -> doc
    this.units = units; // [{ doc, i }]
    this.prefs = prefs;
    this.onPage = onPage; this.onTap = onTap; this.onSwipe = onSwipe;
    this.k = -1;
    this.zoom = 1;
    this.flags = new Set();
    this.lru = new Map(); // k -> unit element (paged mode)
    this.ro = new ResizeObserver(() => { clearTimeout(this._rt); this._rt = setTimeout(() => this.relayout(), 120); });
    this.ro.observe(stage);
    this.bindGestures();
  }
  get mode() { return this.prefs.view === 'scroll' ? 'scroll' : 'paged'; }
  destroy() { this.ro.disconnect(); this.io?.disconnect(); this.stage.innerHTML = ''; this.lru.clear(); this.dead = true; }

  // ---- sizes ----
  avail() {
    const cs = getComputedStyle(this.stage);
    const pt = parseFloat(cs.getPropertyValue('--pad-top')) || 70, pb = parseFloat(cs.getPropertyValue('--pad-bot')) || 110;
    const W = this.stage.clientWidth - 28, H = this.stage.clientHeight - pt - pb;
    return { W: Math.max(200, W), H: Math.max(200, H) };
  }
  aspectOf(k) {
    const u = this.units[k], d = this.docs.get(u.doc);
    return u.ar || d.aspect || 0.77;
  }
  widthFor(k) {
    const { W, H } = this.avail();
    const ar = this.aspectOf(k);
    const fit = this.prefs.fit || 'page';
    const z = this.zoom;
    if (this.mode === 'scroll') return Math.min(W, (ar > 1 ? 1100 : 900)) * z;
    if (fit === 'width') return Math.min(W, ar > 1 ? 1400 : 1000) * z;
    // leave room for speaker notes under a slide
    const notes = this.notesFor(k) ? Math.min(160, 40 + this.notesFor(k).length * 0.25) : 0;
    return Math.min(W, (H - notes) * ar) * z;
  }

  // ---- building a unit ----
  build(k) {
    const u = this.units[k], d = this.docs.get(u.doc);
    const el = document.createElement('div');
    el.className = 'unit loading';
    el.dataset.k = k;
    if (d.kind === 'flow') {
      el.classList.add('flow');
      el.classList.remove('loading');
      el.innerHTML = '<div class="fc"></div>';
    } else {
      el.classList.add('tintable');
      if (d.kind === 'slides') el.classList.add('is-slide');
    }
    if (this.flags.has(k)) el.classList.add('flagged');
    this.size(el, k);
    return el;
  }
  size(el, k) {
    const d = this.docs.get(this.units[k].doc);
    if (d.kind === 'flow') {
      el.style.width = '';
      el.style.setProperty('--flow-size', `${(this.prefs.fontSize || 18) * this.zoom}px`);
      el.style.setProperty('--flow-w', `${(this.prefs.flowW || 720) * Math.max(1, this.zoom * 0.9)}px`);
      return;
    }
    const w = Math.round(this.widthFor(k));
    const h = Math.round(w / this.aspectOf(k));
    el.style.width = `${w}px`;
    el.style.height = `${h}px`;
    el._w = w;
  }
  async draw(el, k) {
    if (el._drawnW && el._drawnW === el._w && el._drawn) return;
    const u = this.units[k], meta = this.docs.get(u.doc);
    const d = await loadDoc(meta);
    if (this.dead) return;
    try {
      if (d.kind === 'pdf') await this.drawPdf(el, d, u.i + 1, k);
      else if (d.kind === 'slides') this.drawSlide(el, d, u.i);
      else if (d.kind === 'images') this.drawImage(el, d, u.i, k);
      else this.drawFlow(el, d, u.i, meta);
      el._drawn = true; el._drawnW = el._w;
    } catch (e) {
      if (e?.name !== 'RenderingCancelledException') console.warn(e);
    }
    el.classList.remove('loading');
  }
  async drawPdf(el, d, n, k) {
    const s = await pdfSize(d, n);
    const ar = s.w / s.h;
    if (Math.abs(ar - this.aspectOf(k)) > 0.01) { this.units[k].ar = ar; this.size(el, k); }
    const w = el._w;
    el._task?.cancel();
    const scale = (w * DPR()) / s.w;
    const vp = s.page.getViewport({ scale });
    const c = document.createElement('canvas');
    c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
    const ctx = c.getContext('2d', { alpha: false });
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
    const task = s.page.render({ canvasContext: ctx, viewport: vp });
    el._task = task;
    await task.promise;
    el._task = null;
    const old = el.querySelector('canvas');
    old ? old.replaceWith(c) : el.prepend(c);
    c.animate?.([{ opacity: old ? 0.6 : 0 }, { opacity: 1 }], { duration: old ? 120 : 220, easing: 'ease-out' });
  }
  drawSlide(el, d, i) {
    let host = el.querySelector('.slide-host');
    if (!host) {
      host = document.createElement('div'); host.className = 'slide-host';
      host.append(renderSlide(d.deck, i, d.urls));
      el.prepend(host);
      host.animate?.([{ opacity: 0 }, { opacity: 1 }], { duration: 200 });
    }
    const slide = host.firstElementChild;
    const scale = el._w / (d.deck.w * (96 / 72));
    slide.style.transform = `scale(${scale})`;
    slide.style.transformOrigin = '0 0';
  }
  drawImage(el, d, i, k) {
    if (el.querySelector('img')) return;
    const img = new Image();
    img.decoding = 'async'; img.draggable = false; img.alt = '';
    img.onload = () => {
      const ar = img.naturalWidth / img.naturalHeight;
      if (Math.abs(ar - this.aspectOf(k)) > 0.01) { this.units[k].ar = ar; this.size(el, k); }
    };
    img.src = d.urls[i];
    el.prepend(img);
  }
  drawFlow(el, d, i, meta) {
    const fc = el.querySelector('.fc');
    if (fc.childElementCount) return;
    fc.innerHTML = d.chunks[i] || '';
    if (!el.querySelector('.pg-foot')) {
      const f = document.createElement('div');
      f.className = 'pg-foot';
      f.textContent = `${i + 1} / ${meta.units}`;
      el.append(f);
    }
  }

  // ---- paged ----
  async show(k, dir = 0) {
    if (k < 0 || k >= this.units.length) return;
    const first = this.k === -1;
    this.k = k;
    if (this.mode === 'scroll') { this.scrollTo(k, !first); return; }
    this.stage.className = 'stage paged';
    const old = this.stage.querySelector('.pane:not(.leaving)');
    const pane = document.createElement('div');
    pane.className = 'pane';
    const col = document.createElement('div');
    col.className = 'pane-col';
    let el = this.lru.get(k);
    if (!el) el = this.build(k); else this.size(el, k);
    this.lru.delete(k); this.lru.set(k, el);
    while (this.lru.size > 6) { const [kk, e] = this.lru.entries().next().value; e._task?.cancel(); this.lru.delete(kk); }
    el.classList.toggle('flagged', this.flags.has(k));
    col.append(el);
    const notes = this.notesFor(k);
    if (notes) { const n = document.createElement('div'); n.className = 'notes'; n.style.setProperty('--nw', `${el._w || 720}px`); n.textContent = notes; col.append(n); }
    pane.append(col);
    this.stage.append(pane);
    pane.scrollTop = 0;
    this.fitClass(pane, el, k);
    this.draw(el, k).then(() => this.fitClass(pane, el, k));
    // the page turn
    const style = reduced() || first ? 'none' : this.prefs.turn || 'slide';
    if (old) {
      old.classList.add('leaving');
      const from = old._dx || 0;
      const D = 0.36 * window.innerWidth;
      let outF, inF, dur = 420, ease = 'cubic-bezier(0.32, 0.72, 0, 1)';
      if (style === 'slide') {
        outF = [{ transform: `translateX(${from}px)`, opacity: 1 }, { transform: `translateX(${-dir * D}px) scale(0.96)`, opacity: 0 }];
        inF = [{ transform: `translateX(${dir * D + from * 0.2}px) scale(0.96)`, opacity: 0 }, { transform: 'none', opacity: 1 }];
      } else if (style === 'fade') {
        outF = [{ opacity: 1 }, { opacity: 0 }]; inF = [{ opacity: 0 }, { opacity: 1 }]; dur = 260; ease = 'ease';
      } else if (style === 'stack') {
        outF = [{ transform: `translateX(${from}px)`, opacity: 1 }, { transform: dir > 0 ? 'scale(0.9)' : `translateX(${window.innerWidth}px) rotate(4deg)`, opacity: dir > 0 ? 0 : 1 }];
        inF = [{ transform: dir > 0 ? `translateX(${window.innerWidth * 0.9}px) rotate(4deg)` : 'scale(0.9)', opacity: dir > 0 ? 1 : 0 }, { transform: 'none', opacity: 1 }];
        if (dir > 0) pane.style.zIndex = 2; else old.style.zIndex = 2;
        dur = 520; ease = 'cubic-bezier(0.2, 0.9, 0.25, 1)';
      } else { outF = [{ opacity: 1 }, { opacity: 0 }]; inF = [{ opacity: 1 }, { opacity: 1 }]; dur = 1; }
      if (!dir) { outF = [{ opacity: 1 }, { opacity: 0 }]; inF = [{ opacity: 0, transform: 'scale(0.98)' }, { opacity: 1, transform: 'none' }]; dur = 260; }
      old.animate(outF, { duration: dur, easing: ease, fill: 'forwards' }).onfinish = () => old.remove();
      pane.animate(inF, { duration: dur, easing: ease });
      for (const p of this.stage.querySelectorAll('.pane.leaving')) if (p !== old) p.remove();
    } else if (!reduced()) {
      el.animate([{ opacity: 0, transform: 'translateY(16px) scale(0.98)' }, { opacity: 1, transform: 'none' }], { duration: 500, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.1)' });
    }
    // get the next ones ready
    for (const j of [k + 1, k - 1, k + 2]) {
      if (j < 0 || j >= this.units.length || this.lru.has(j)) continue;
      const pre = this.build(j);
      this.lru.set(j, pre);
      setTimeout(() => this.draw(pre, j), 140);
    }
  }
  fitClass(pane, el, k) { pane.classList.toggle('fits', this.docs.get(this.units[k].doc).kind !== 'flow' && (el.offsetHeight <= pane.clientHeight - 20 || !el.offsetHeight)); }
  notesFor(k) {
    if (!this.prefs.notes) return '';
    const u = this.units[k], d = this.docs.get(u.doc);
    if (d.kind !== 'slides') return '';
    return this._notes?.[u.doc]?.[u.i] || '';
  }
  async primeNotes() {
    this._notes = {};
    for (const d of this.docs.values()) if (d.kind === 'slides') { const x = await loadDoc(d); this._notes[d.id] = x.deck.slides.map((s) => s.notes); }
  }
  // Space / Page Down on a long page scrolls it first
  scrollBy(dir) {
    const pane = this.mode === 'scroll' ? this.stage : this.stage.querySelector('.pane:not(.leaving)');
    if (!pane) return false;
    if (this.mode === 'scroll') { pane.scrollBy({ top: dir * pane.clientHeight * 0.85, behavior: 'smooth' }); return true; }
    const max = pane.scrollHeight - pane.clientHeight;
    if (dir > 0 && pane.scrollTop < max - 4) { pane.scrollBy({ top: pane.clientHeight * 0.85, behavior: 'smooth' }); return true; }
    if (dir < 0 && pane.scrollTop > 4) { pane.scrollBy({ top: -pane.clientHeight * 0.85, behavior: 'smooth' }); return true; }
    return false;
  }

  // ---- scroll ----
  buildScroll() {
    this.stage.className = 'stage scroll';
    this.stage.innerHTML = '';
    this.io?.disconnect();
    const sc = document.createElement('div');
    sc.className = 'scroller';
    this.els = this.units.map((_, k) => this.build(k));
    sc.append(...this.els);
    this.stage.append(sc);
    this.io = new IntersectionObserver((ents) => {
      for (const e of ents) {
        const k = Number(e.target.dataset.k);
        if (e.isIntersecting) this.draw(e.target, k);
      }
    }, { root: this.stage, rootMargin: '120% 0px' });
    this.els.forEach((e) => this.io.observe(e));
    // flow pages have their words right away so the column doesn't jump about
    this.units.forEach((u, k) => { if (this.docs.get(u.doc).kind === 'flow') this.draw(this.els[k], k); });
    this.stage.onscroll = () => {
      if (this._lock) return;
      cancelAnimationFrame(this._sf);
      this._sf = requestAnimationFrame(() => {
        const r = this.stage.getBoundingClientRect();
        const y = r.top + r.height * 0.38;
        let best = this.k;
        for (let k = 0; k < this.els.length; k++) {
          const b = this.els[k].getBoundingClientRect();
          if (b.top <= y && b.bottom >= y) { best = k; break; }
          if (b.top > y) { best = Math.max(0, k - 1); break; }
        }
        if (best !== this.k) { const from = this.k; this.k = best; this.onPage?.(best, best > from ? 1 : -1, 'scroll'); }
      });
    };
  }
  scrollTo(k, smooth) {
    if (!this.els || this.stage.classList.contains('paged')) this.buildScroll();
    const el = this.els[k];
    if (!el) return;
    this._lock = true;
    const top = el.offsetTop - (parseFloat(getComputedStyle(this.stage).getPropertyValue('--pad-top')) || 70) + 2;
    this.stage.scrollTo({ top, behavior: smooth && !reduced() ? 'smooth' : 'auto' });
    clearTimeout(this._lt);
    this._lt = setTimeout(() => { this._lock = false; }, smooth ? 700 : 60);
  }

  // ---- changes ----
  setMode() {
    const k = Math.max(0, this.k);
    this.stage.innerHTML = ''; this.els = null; this.lru.clear(); this.io?.disconnect(); this.stage.onscroll = null;
    this.k = -1;
    this.show(k, 0);
  }
  relayout() {
    if (this.dead || this.k < 0) return;
    if (this.mode === 'scroll' && this.els) {
      const k = this.k;
      this.els.forEach((e, j) => { this.size(e, j); e._drawn = e._drawn && e._drawnW === e._w; });
      for (const e of this.els) { const b = e.getBoundingClientRect(); if (b.bottom > -innerHeight && b.top < innerHeight * 2) this.draw(e, Number(e.dataset.k)); }
      this.scrollTo(k, false);
      return;
    }
    for (const [j, e] of this.lru) { this.size(e, j); }
    const pane = this.stage.querySelector('.pane:not(.leaving)');
    const el = pane?.querySelector('.unit');
    if (el) { this.draw(el, this.k).then(() => this.fitClass(pane, el, this.k)); const n = pane.querySelector('.notes'); if (n) n.style.setProperty('--nw', `${el._w || 720}px`); }
  }
  setZoom(z) { this.zoom = Math.max(0.5, Math.min(3, z)); this.relayout(); return this.zoom; }
  setFlag(k, on) {
    on ? this.flags.add(k) : this.flags.delete(k);
    const el = (this.els && this.els[k]) || this.lru.get(k);
    el?.classList.toggle('flagged', on);
    if (on && el && !reduced()) el.animate([{ transform: 'scale(1)' }, { transform: 'scale(1.015)' }, { transform: 'scale(1)' }], { duration: 380, easing: 'ease-out' });
  }

  // ---- touch: swipe pages (the page follows your finger), tap to show/hide the bars ----
  bindGestures() {
    let g = null;
    const st = this.stage;
    st.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      g = { x: e.clientX, y: e.clientY, t: performance.now(), id: e.pointerId, type: e.pointerType, moved: false, swipe: false };
    });
    st.addEventListener('pointermove', (e) => {
      if (!g || e.pointerId !== g.id) return;
      const dx = e.clientX - g.x, dy = e.clientY - g.y;
      if (!g.moved && Math.hypot(dx, dy) > 8) {
        g.moved = true;
        const pane = st.querySelector('.pane:not(.leaving)');
        const canPan = pane && pane.scrollWidth > pane.clientWidth + 2;
        g.swipe = this.mode === 'paged' && g.type !== 'mouse' && Math.abs(dx) > Math.abs(dy) * 1.3 && !canPan;
        g.pane = pane;
      }
      if (g.swipe && g.pane) {
        const edge = (dx > 0 && this.k === 0) || (dx < 0 && this.k === this.units.length - 1);
        const v = edge ? dx * 0.25 : dx;
        g.pane.style.transform = `translateX(${v}px) rotate(${v / 120}deg)`;
        g.pane._dx = v;
      }
    }, { passive: true });
    const end = (e) => {
      if (!g || e.pointerId !== g.id) return;
      const dx = e.clientX - g.x, dt = performance.now() - g.t;
      if (g.swipe && g.pane) {
        const fast = Math.abs(dx) / dt > 0.45;
        const dir = dx < 0 ? 1 : -1;
        const ok = (Math.abs(dx) > 70 || (fast && Math.abs(dx) > 24)) && this.onSwipe?.(dir);
        if (!ok) {
          const p = g.pane, from = p._dx || 0;
          p.style.transform = ''; p._dx = 0;
          p.animate([{ transform: `translateX(${from}px) rotate(${from / 120}deg)` }, { transform: 'none' }], { duration: 420, easing: 'cubic-bezier(0.2, 0.9, 0.3, 1.2)' });
        } else { g.pane.style.transform = ''; }
      } else if (!g.moved && e.type === 'pointerup' && !e.target.closest('a, button, input, textarea')) {
        if (!window.getSelection()?.toString()) this.onTap?.(e);
      }
      g = null;
    };
    st.addEventListener('pointerup', end);
    st.addEventListener('pointercancel', end);
    // ctrl/⌘ + wheel or a trackpad pinch zooms
    st.addEventListener('wheel', (e) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      this._zw = (this._zw || this.zoom) * Math.exp(-e.deltaY * 0.01);
      clearTimeout(this._zt);
      this._zt = setTimeout(() => { this.onZoom?.(this.setZoom(this._zw)); this._zw = 0; }, 90);
    }, { passive: false });
  }

  // ---- small pictures of pages for the overview ----
  async thumb(k, w = 120) {
    const u = this.units[k], meta = this.docs.get(u.doc);
    const d = await loadDoc(meta);
    if (d.kind === 'pdf') {
      const s = await pdfSize(d, u.i + 1);
      const vp = s.page.getViewport({ scale: (w * DPR()) / s.w });
      const c = document.createElement('canvas');
      c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
      await s.page.render({ canvasContext: c.getContext('2d', { alpha: false }), viewport: vp }).promise;
      return { el: c, ar: s.w / s.h };
    }
    if (d.kind === 'slides') {
      const box = document.createElement('div');
      box.style.cssText = 'position:absolute;inset:0;overflow:hidden';
      const sl = renderSlide(d.deck, u.i, d.urls);
      sl.style.transformOrigin = '0 0';
      sl.style.transform = `scale(${w / (d.deck.w * (96 / 72))})`;
      box.append(sl);
      return { el: box, ar: d.deck.w / d.deck.h };
    }
    if (d.kind === 'images') { const im = new Image(); im.src = d.urls[u.i]; return { el: im, ar: meta.aspect }; }
    const t = document.createElement('div');
    t.className = 'ov-text';
    const tmp = document.createElement('div'); tmp.innerHTML = d.chunks[u.i] || '';
    const h = tmp.querySelector('h1,h2,h3')?.textContent || '';
    t.innerHTML = `${h ? `<b></b>` : ''}<span></span>`;
    if (h) t.querySelector('b').textContent = h;
    t.querySelector('span').textContent = tmp.textContent.replace(h, '').trim().slice(0, 420);
    return { el: t, ar: 0.77 };
  }
}
