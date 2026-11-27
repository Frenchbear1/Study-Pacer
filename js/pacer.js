// The pacer: how long each page gets, and how you're doing against it.
//
// A session is a run of units (pages and slides) across one or more documents. You give it time three ways:
// a total ("90 minutes"), a pace ("45 s a page") or a deadline ("done by 4:30"). Each unit's share of that time
// comes from its weight:
//   smart  – by how much is on it: a dense page of text gets more than a title slide (strength adjustable)
//   even   – every unit the same
//   front / back – more time early on (you're fresh) or towards the end (the hard part)
// When you finish a page early or late the difference is handled one of three ways:
//   spread – the time left is shared again over every page still to go (the default; nothing piles up)
//   carry  – the difference rolls onto the next page only (the original Study Pacer's way)
//   fixed  – every page keeps its planned time; you just see whether you're ahead or behind
// Breaks can come every N minutes of reading, every N pages, or between documents, and always land on a
// page turn so they never cut a page in half.
// Pure logic: no DOM. The app calls tick() and the movement functions and renders snapshot().

export const DEFAULTS = {
  mode: 'total', total: 60 * 60, per: 60, by: null,
  profile: 'smart', strength: 70, rebalance: 'spread',
  breakEvery: 0, breakUnit: 'min', breakLen: 5 * 60, breakDocs: false,
  method: 'none', longLen: 20 * 60, longEvery: 0, interrupt: true, flowtime: false, recall: true,
};

// Study methods, as starting points you can change. Times in minutes.
export const METHODS = [
  { id: 'none', name: 'No breaks', blurb: 'Just the pacer. Take a break whenever you like.', focus: 0 },
  { id: 'pomodoro', name: 'Pomodoro', blurb: '25 minutes on, 5 off. Every fourth break is a long one. Francesco Cirillo’s classic.', focus: 25, short: 5, long: 20, every: 4 },
  { id: 'pomo50', name: 'Long Pomodoro', blurb: '50 on, 10 off: room to get into harder material, still with regular rests.', focus: 50, short: 10, long: 25, every: 3 },
  { id: '5217', name: '52 / 17', blurb: 'From DeskTime’s study of its most productive users: 52 minutes of focus, 17 of real rest.', focus: 52, short: 17, long: 0, every: 0 },
  { id: 'ultradian', name: 'Ultradian 90', blurb: 'One deep 90-minute block, then 20 minutes off, following the body’s ~90-minute alertness cycle.', focus: 90, short: 20, long: 0, every: 0 },
  { id: 'flowtime', name: 'Flowtime', blurb: 'No timer interrupts you. Break when your focus fades; you earn a fifth of the time you studied.', focus: 0, flowtime: true },
  { id: 'sprints', name: 'Page sprints', blurb: 'A short break every 10 pages, wherever the clock is. The original Study Pacer’s way.', pages: 10, short: 5 },
];
export function methodCfg(m, over = {}) {
  const x = { ...m, ...over };
  if (x.flowtime) return { method: m.id, breakEvery: 0, breakUnit: 'min', flowtime: true, longEvery: 0, interrupt: false };
  if (x.pages) return { method: m.id, breakEvery: x.pages, breakUnit: 'pages', breakLen: (x.short || 5) * 60, longEvery: 0, flowtime: false };
  return { method: m.id, breakEvery: x.focus || 0, breakUnit: 'min', breakLen: (x.short || 5) * 60, longLen: (x.long || 0) * 60, longEvery: x.long ? x.every || 0 : 0, flowtime: false };
}

const sum = (a) => a.reduce((s, v) => s + v, 0);

export function weightsFor(units, { profile = 'smart', strength = 70 } = {}) {
  const n = units.length;
  if (!n) return [];
  let w;
  if (profile === 'smart') {
    const k = Math.max(0, Math.min(100, strength)) / 100;
    // the median rather than the mean, so one wall of text doesn't shrink everything else
    const sorted = units.map((u) => u.words || 0).sort((a, b) => a - b);
    const med = sorted[Math.floor(n / 2)] || 0;
    const ref = Math.max(30, med || sum(sorted) / n || 0);
    w = units.map((u) => {
      const r = Math.min(4, (u.words || 0) / ref);
      // a page with nothing to read (a picture, a divider) still gets a moment
      return Math.max(0.25, 1 - k + k * r);
    });
  } else w = units.map(() => 1);
  if (profile === 'front' || profile === 'back') {
    const a = 1.4;
    w = w.map((x, i) => { const t = n > 1 ? i / (n - 1) : 0; return x * (profile === 'front' ? a - (a - 1) * t : 1 + (a - 1) * t); });
  }
  const mean = sum(w) / n;
  return w.map((x) => x / mean); // mean 1, so "per page" times read naturally
}

export class Pacer {
  constructor(units, cfg, state) {
    this.units = units; // [{ doc, i, words }]
    this.cfg = { ...DEFAULTS, ...cfg };
    this.w = weightsFor(units, this.cfg);
    const n = units.length;
    this.s = state || {
      cur: 0, done: new Array(n).fill(false), spent: new Array(n).fill(0),
      alloc: new Array(n).fill(null), arriveLeft: null, visit: 0,
      work: 0, breakTotal: 0, added: 0, started: null, last: null, paused: false,
      brk: null, workSinceBreak: 0, pagesSinceBreak: 0, breaks: 0, log: [], finished: null, overNotified: false,
    };
    if (!state) {
      this.s.budget = this.initialBudget();
      this.arrive(0);
    }
  }

  // ---- time ----
  initialBudget(now = Date.now()) {
    const c = this.cfg, n = this.units.length;
    if (c.mode === 'per') return c.per * n;
    if (c.mode === 'by' && c.by) return Math.max(60, (c.by - now) / 1000 - this.plannedBreaks(n, (c.by - now) / 1000, this.docsLeftFrom(0)));
    return c.total;
  }
  // seconds of reading left in the plan
  remaining(now = Date.now()) {
    const c = this.cfg, s = this.s;
    if (c.mode === 'by' && c.by) {
      const raw = (c.by - now) / 1000 + s.added;
      return raw - this.plannedBreaks(this.left(), Math.max(0, raw), this.docsLeft()) - (s.brk ? Math.max(0, (s.brk.until - now) / 1000) : 0);
    }
    return s.budget + s.added - s.work;
  }
  start(now = Date.now()) { this.s.started ??= now; this.s.last = now; this.s.paused = false; this.log('start'); }
  tick(now = Date.now()) {
    const s = this.s;
    if (s.finished || s.started == null) return;
    const dt = Math.min(5, Math.max(0, (now - (s.last ?? now)) / 1000)); // a sleeping tab doesn't count
    s.last = now;
    if (s.paused || s.due) return;
    if (s.brk) {
      s.breakTotal += dt;
      // the break is up: wait for you to say you're ready, so the clock never runs while you're away
      if (now >= s.brk.until) s.brk.ready = true;
      return;
    }
    s.spent[s.cur] += dt;
    s.work += dt;
    s.workSinceBreak += dt;
    // a focus block ran out: interrupt (the app shows it; nothing counts until you choose)
    const c = this.cfg;
    if (c.interrupt && c.breakUnit === 'min' && c.breakEvery > 0 && !s.defer && this.left() > 0 && s.workSinceBreak >= c.breakEvery * 60 + (s.snooze || 0)) {
      s.due = { at: now, len: this.nextBreakLen(), long: this.isLong() };
      this.log('due');
    }
  }
  focusLeft() { const c = this.cfg; return c.breakUnit === 'min' && c.breakEvery > 0 ? c.breakEvery * 60 + (this.s.snooze || 0) - this.s.workSinceBreak : null; }
  isLong() { const c = this.cfg; return !!(c.longEvery && c.longLen && ((this.s.cycles || 0) + 1) % c.longEvery === 0); }
  nextBreakLen() {
    const c = this.cfg;
    if (c.flowtime) return Math.round(Math.max(120, Math.min(30 * 60, this.s.workSinceBreak / 5)) / 30) * 30;
    return this.isLong() ? c.longLen : c.breakLen;
  }
  // answers to an interruption
  acceptDue(now = Date.now()) { const d = this.s.due; if (!d) return; this.s.due = null; this.s.last = now; this.startBreak({ len: d.len, why: 'focus', long: d.long }, now); }
  snoozeDue(sec, now = Date.now()) { this.s.due = null; this.s.last = now; this.s.snooze = (this.s.snooze || 0) + sec + Math.max(0, this.s.workSinceBreak - (this.cfg.breakEvery * 60 + (this.s.snooze || 0))); this.log('snooze', { sec }); }
  deferDue(now = Date.now()) { this.s.due = null; this.s.last = now; this.s.defer = true; this.log('defer'); }
  skipDue(now = Date.now()) { this.s.due = null; this.s.last = now; this.s.workSinceBreak = 0; this.s.snooze = 0; this.s.cycles = (this.s.cycles || 0) + 1; this.log('skip'); }
  pause(now = Date.now()) { this.tick(now); this.s.paused = true; this.log('pause'); }
  resume(now = Date.now()) { this.s.last = now; this.s.paused = false; this.log('resume'); }
  addTime(sec) { this.s.added += sec; this.s.overNotified = false; this.reallocCurrent(); this.log('add', { sec }); }

  // ---- movement ----
  left() { return this.s.done.reduce((a, d) => a + (d ? 0 : 1), 0); }
  docsLeftFrom(i) { const set = new Set(); for (let j = i; j < this.units.length; j++) set.add(this.units[j].doc); return set.size; }
  docsLeft() { const set = new Set(); this.units.forEach((u, j) => { if (!this.s.done[j]) set.add(u.doc); }); return set.size; }
  arrive(i, now = Date.now()) {
    const s = this.s, c = this.cfg;
    s.cur = i;
    s.visit = s.spent[i];
    if (s.done[i]) return; // reviewing a page you've read: the clock just counts
    if (c.rebalance === 'spread') {
      // what's left, shared by weight over every page still to read; time already spent here stays this page's
      const wl = sum(this.w.filter((_, j) => !s.done[j])) || 1;
      s.alloc[i] = s.spent[i] + Math.max(5, (Math.max(0, this.remaining(now)) * this.w[i]) / wl);
    } else if (s.alloc[i] == null) {
      const base = (this.baseBudget() * this.w[i]) / this.units.length;
      s.alloc[i] = c.rebalance === 'carry' ? Math.max(5, base + (s.carry || 0)) : base;
      s.carry = 0;
    }
  }
  baseBudget() { return this.cfg.mode === 'by' ? this.s.budget : this.s.budget + this.s.added; }
  reallocCurrent() { const i = this.s.cur; if (!this.s.done[i] && this.cfg.rebalance === 'spread') this.arrive(i); }
  // move to unit j. Leaving forward marks the page read.
  go(j, now = Date.now()) {
    const s = this.s;
    if (j === s.cur || j < 0 || j >= this.units.length || s.finished) return { moved: false };
    this.tick(now);
    const from = s.cur;
    const forward = j > from;
    let breakNow = null;
    if (forward && !s.done[from]) {
      if (this.cfg.rebalance === 'carry') s.carry = (s.alloc[from] ?? 0) - s.spent[from];
      s.done[from] = true;
      s.pagesSinceBreak++;
      this.log('page', { i: from, sec: s.spent[from], planned: s.alloc[from] });
      breakNow = this.breakDue(from, j);
    }
    this.arrive(j, now);
    if (breakNow && s.started != null && this.left() > 0) this.startBreak(breakNow, now);
    return { moved: true, from, to: j, break: breakNow };
  }
  next(now) {
    const s = this.s;
    if (s.cur < this.units.length - 1) return this.go(s.cur + 1, now);
    // the last page: finishing it ends the session if everything's read
    this.tick(now);
    if (!s.done[s.cur]) { s.done[s.cur] = true; this.log('page', { i: s.cur, sec: s.spent[s.cur], planned: s.alloc[s.cur] }); }
    const unread = s.done.indexOf(false);
    if (unread === -1) { this.finish(now); return { moved: false, finished: true }; }
    return { moved: false, unread };
  }
  prev(now) { return this.go(this.s.cur - 1, now); }
  finish(now = Date.now()) { this.tick(now); this.s.finished = now; this.s.brk = null; this.log('finish'); }

  // ---- breaks ----
  breakDue(from, to) {
    const c = this.cfg, s = this.s;
    if (s.defer) return { why: 'focus', len: this.nextBreakLen(), long: this.isLong() };
    if (c.breakDocs && this.units[from].doc !== this.units[to].doc) return { why: 'doc', len: this.nextBreakLen() || c.breakLen || 300, long: this.isLong() };
    if (!c.breakEvery) return null;
    if (c.breakUnit === 'min' && !c.interrupt && s.workSinceBreak >= c.breakEvery * 60) return { why: 'focus', len: this.nextBreakLen(), long: this.isLong() };
    if (c.breakUnit === 'pages' && s.pagesSinceBreak >= c.breakEvery) return { why: 'pages', len: this.nextBreakLen(), long: this.isLong() };
    return null;
  }
  startBreak({ len, why, long = false }, now = Date.now()) {
    const s = this.s;
    // what you read since the last break, for the recall check
    const since = s.log.filter((e) => e.type === 'page' && e.t > (s.lastBreakAt || 0)).map((e) => e.i);
    s.brk = { start: now, until: now + len * 1000, len, why, long, since };
    s.breaks++;
    s.defer = false; s.snooze = 0;
    if (why !== 'manual' && why !== 'doc') s.cycles = (s.cycles || 0) + 1;
    this.log('break', { len, why });
  }
  extendBreak(sec) { if (this.s.brk) { this.s.brk.until += sec * 1000; this.s.brk.len += sec; } }
  endBreak(now = Date.now()) {
    const s = this.s;
    if (!s.brk) return;
    this.log('breakEnd', { sec: Math.round((now - s.brk.start) / 1000) });
    s.brk = null; s.workSinceBreak = 0; s.pagesSinceBreak = 0; s.last = now; s.lastBreakAt = now;
    this.reallocCurrent();
  }
  takeBreakNow(now = Date.now()) { this.tick(now); this.s.due = null; this.startBreak({ len: this.cfg.flowtime ? this.nextBreakLen() : this.cfg.breakLen || 300, why: this.cfg.flowtime ? 'flow' : 'manual' }, now); }
  // how much break time is still to come if things go to plan
  plannedBreaks(pagesLeft = this.left(), workLeft = Math.max(0, this.remaining()), docsLeft = this.docsLeft()) {
    const c = this.cfg, s = this.s;
    let n = 0;
    if (c.breakEvery > 0) {
      if (c.breakUnit === 'min') n = Math.floor((workLeft + (s?.workSinceBreak || 0)) / (c.breakEvery * 60));
      else n = Math.floor((pagesLeft - 1 + (s?.pagesSinceBreak || 0)) / c.breakEvery);
    }
    if (c.breakDocs) n = Math.max(n, docsLeft - 1);
    n = Math.max(0, n);
    const longs = c.longEvery && c.longLen ? Math.floor(((s?.cycles || 0) % c.longEvery + n) / c.longEvery) : 0;
    return n * (c.breakLen || 0) + longs * Math.max(0, c.longLen - c.breakLen);
  }

  // ---- what to show ----
  snapshot(now = Date.now()) {
    const s = this.s, n = this.units.length, i = s.cur;
    const remaining = this.remaining(now);
    const review = s.done[i];
    const pageSpent = s.spent[i] - s.visit;
    const alloc = s.alloc[i] ?? 0;
    const pageLeft = review ? null : alloc - s.spent[i];
    const doneN = n - this.left();
    // your own pace: seconds per unit of weight, from what you've finished (blended in as it becomes reliable)
    let doneW = 0, doneT = 0;
    s.done.forEach((d, j) => { if (d) { doneW += this.w[j]; doneT += s.spent[j]; } });
    const planRate = Math.max(1, (this.cfg.mode === 'by' ? s.budget : s.budget + s.added)) / n;
    const myRate = doneW > 0 ? doneT / doneW : planRate;
    const trust = Math.min(1, doneN / 5, doneT / 120);
    const rate = planRate + (myRate - planRate) * trust;
    let leftW = 0;
    s.done.forEach((d, j) => { if (!d) leftW += this.w[j]; });
    const curExtra = review ? 0 : Math.min(this.w[i] * rate, s.spent[i]);
    const predWork = Math.max(0, leftW * rate - curExtra);
    const brkLeft = s.brk ? Math.max(0, (s.brk.until - now) / 1000) : 0;
    const plannedBrk = this.plannedBreaks(this.left(), Math.max(0, predWork), this.docsLeft());
    const predicted = s.finished || now + (predWork + plannedBrk + brkLeft) * 1000;
    const planned = this.cfg.mode === 'by' && this.cfg.by ? this.cfg.by + s.added * 1000 : now + (Math.max(0, remaining) + this.plannedBreaks() + brkLeft) * 1000;
    // per document
    const docs = [];
    const byDoc = new Map();
    this.units.forEach((u, j) => {
      let d = byDoc.get(u.doc);
      if (!d) { d = { doc: u.doc, total: 0, done: 0, first: j, last: j }; byDoc.set(u.doc, d); docs.push(d); }
      d.total++; d.last = j; if (s.done[j]) d.done++;
    });
    const unitPlanned = (j) => s.alloc[j] ?? (Math.max(0, remaining) * this.w[j]) / (leftW || 1);
    return {
      i, n, unit: this.units[i], review, finished: !!s.finished, paused: s.paused, brk: s.brk ? { ...s.brk, left: brkLeft } : null,
      pageAlloc: alloc, pageLeft, pageSpent, pageTotalSpent: s.spent[i],
      remaining, budget: s.budget + s.added, work: s.work, breakTotal: s.breakTotal,
      left: this.left(), doneN, docs, docsLeft: this.docsLeft(),
      planned, predicted, delta: (planned - predicted) / 1000, trust,
      over: this.cfg.mode !== 'open' && remaining <= 0 && this.left() > 0,
      open: this.cfg.mode === 'open', due: s.due, focusLeft: this.focusLeft(), cycles: s.cycles || 0, defer: !!s.defer,
      unitPlanned,
    };
  }
  stats() {
    const s = this.s;
    const pages = this.units.map((u, j) => ({ j, doc: u.doc, i: u.i, words: u.words, sec: s.spent[j], planned: s.alloc[j], done: s.done[j] })).filter((p) => p.sec > 0.5);
    const timed = pages.filter((p) => p.done);
    const secs = timed.map((p) => p.sec);
    const words = timed.reduce((a, p) => a + (p.words || 0), 0);
    const wsec = timed.filter((p) => p.words > 40).reduce((a, p) => a + p.sec, 0);
    const wwords = timed.filter((p) => p.words > 40).reduce((a, p) => a + p.words, 0);
    return {
      pages, work: s.work, breaks: s.breaks, breakTotal: s.breakTotal,
      avg: secs.length ? sum(secs) / secs.length : 0,
      fastest: timed.length ? timed.reduce((a, b) => (b.sec < a.sec ? b : a)) : null,
      slowest: timed.length ? timed.reduce((a, b) => (b.sec > a.sec ? b : a)) : null,
      words, wpm: wsec > 30 ? Math.round(wwords / (wsec / 60)) : null,
    };
  }
  log(type, extra = {}) { this.s.log.push({ t: Date.now(), type, ...extra }); if (this.s.log.length > 4000) this.s.log.splice(0, 1000); }
  toJSON() { return { cfg: this.cfg, s: this.s }; }
}

// A sensible amount of time for a set of units, from how fast you've read before (or a careful study pace).
export function suggestTime(units, history) {
  const wpm = history?.wpm || 160;
  const slide = history?.slide || 40;
  let t = 0;
  for (const u of units) t += Math.max(u.kind === 'slides' ? slide * 0.5 : 20, ((u.words || 0) / wpm) * 60 + (u.kind === 'slides' ? slide * 0.4 : 8));
  return Math.round(t / 60) * 60 || 60;
}

// ---- formatting ----
export function clock(sec, { sign = false, hours = true } = {}) {
  const neg = sec < 0;
  let v = Math.round(Math.abs(sec));
  const h = Math.floor(v / 3600); v -= h * 3600;
  const m = Math.floor(v / 60); const s = v - m * 60;
  const body = h && hours ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${h * 60 + m}:${String(s).padStart(2, '0')}`;
  return (neg ? '−' : sign ? '+' : '') + body;
}
export function human(sec) {
  sec = Math.round(Math.abs(sec));
  if (sec < 60) return `${sec}s`;
  const h = Math.floor(sec / 3600), m = Math.round((sec % 3600) / 60);
  if (!h) return `${m} min`;
  return m ? `${h} h ${m} min` : `${h} h`;
}
export const timeOfDay = (ms) => new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
