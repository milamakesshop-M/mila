/*
 * Mila for Excel: the bridge.
 * Mila Pro was written for Google Sheets, where a script reads and writes cells one by one and the sheet recalculates
 * in between. Excel add-ins talk to the workbook in batches instead. So for every message:
 *   1. load a snapshot of the cells Mila reads (the pages, and the parts of the Engine she uses);
 *   2. run Mila on the snapshot ("first pass"), writing nothing, only noting each change;
 *   3. send all the changes to Excel in one go, and let Excel recalculate;
 *   4. load the Engine again and run Mila a second time on the same message ("second pass"): her decisions are the
 *      same (she sees the same pages), but her answer now uses the recalculated numbers. Nothing is written twice.
 * Mila's own code (Brain, Mind, Persona, Code) is shared with Mila Pro, unchanged.
 */
var MilaBridge = (function () {
  'use strict';
  function colNum(c) { var n = 0; for (var i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64); return n; }
  function colLet(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseA1(a1) { var m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1); return { c: colNum(m[1]), r: +m[2] }; }
  function rangeRC(a1) {
    var p = a1.replace(/\$/g, '').split(':'), A = parseA1(p[0]), B = parseA1(p[1] || p[0]);
    return { r0: Math.min(A.r, B.r), c0: Math.min(A.c, B.c), r1: Math.max(A.r, B.r), c1: Math.max(A.c, B.c) };
  }
  // Excel stores dates as numbers; a cell formatted as a date is a date for Mila (as in Google Sheets)
  var EPOCH = Date.UTC(1899, 11, 30);
  function serialToDate(n) { var ms = Math.round(n * 86400000) + EPOCH, d = new Date(ms); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()); }
  function dateToSerial(d) { return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EPOCH) / 86400000; }
  function isDateFormat(f) { f = String(f || '').replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''); return /[dy]/i.test(f) || /m{3,}/i.test(f); }
  function ymd(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }

  // ---------------------------------------------------------------- a snapshot of the workbook, as Mila reads it
  // book = { sheets: { name: { cells: { A1: [value, formula] } } }, order: [names], active: name, theme: 'light'|'dark' }
  function makeWorld(book, opts) {
    opts = opts || {};
    var writes = [], firstWriteDone = false, active = book.active, overlay = {}, readAfter = { engine: false };
    var props = opts.props || { user: {}, doc: {} }, propWrites = [];
    function cellOf(sheet, a) {
      var o = overlay[sheet] && overlay[sheet][a];
      if (o) return o;
      if (sheet === 'Engine' && firstWriteDone) readAfter.engine = true;
      var src = (opts.engineAfter && sheet === 'Engine' && firstWriteDone) ? opts.engineAfter : book;
      var sh = src.sheets[sheet] || book.sheets[sheet];
      return (sh && sh.cells[a]) || null;
    }
    function record(w) {
      firstWriteDone = true;
      writes.push(w);
      if (w.a) {
        var rc = rangeRC(w.a);
        overlay[w.s] = overlay[w.s] || {};
        for (var r = rc.r0; r <= rc.r1; r++) for (var c = rc.c0; c <= rc.c1; c++) {
          var key = colLet(c) + r, v;
          if (w.op === 'clear') v = ['', null];
          else if (w.op === 'formula') v = [cellOf(w.s, key) ? cellOf(w.s, key)[0] : '', w.f];
          else if (w.op === 'values') v = [w.v[r - rc.r0][c - rc.c0], null];
          else v = [w.v, null];
          if (w.op === 'values' && typeof v[0] === 'string' && v[0].charAt(0) === '=') v = ['', v[0]];
          overlay[w.s][key] = v;
        }
      }
    }
    function Range(sheet, a1) {
      this.sheet = sheet; this.a1 = a1.replace(/\$/g, '');
      var rc = rangeRC(this.a1); this.r0 = rc.r0; this.c0 = rc.c0; this.r1 = rc.r1; this.c1 = rc.c1;
    }
    Range.prototype = {
      _cells: function () { var out = []; for (var r = this.r0; r <= this.r1; r++) { var row = []; for (var c = this.c0; c <= this.c1; c++) row.push(colLet(c) + r); out.push(row); } return out; },
      getValues: function () { var s = this.sheet.name; return this._cells().map(function (row) { return row.map(function (a) { var x = cellOf(s, a); return x && x[0] != null ? x[0] : ''; }); }); },
      getValue: function () { return this.getValues()[0][0]; },
      getFormulas: function () { var s = this.sheet.name; return this._cells().map(function (row) { return row.map(function (a) { var x = cellOf(s, a); return x && x[1] ? x[1] : ''; }); }); },
      getFormula: function () { return this.getFormulas()[0][0]; },
      getDisplayValue: function () { var v = this.getValue(); return v === true ? 'TRUE' : v === false ? 'FALSE' : v instanceof Date ? ymd(v) : String(v); },
      getDisplayValues: function () { return this.getValues().map(function (r) { return r.map(function (v) { return v instanceof Date ? ymd(v) : String(v); }); }); },
      setValue: function (v) {
        if (this.r0 === this.r1 && this.c0 === this.c1) {
          if (typeof v === 'string' && v.charAt(0) === '=') record({ op: 'formula', s: this.sheet.name, a: this.a1, f: v });
          else record({ op: 'value', s: this.sheet.name, a: this.a1, v: v });
        } else { var rows = this._cells().map(function (row) { return row.map(function () { return v; }); }); this.setValues(rows); }
        return this;
      },
      setValues: function (vals) { record({ op: 'values', s: this.sheet.name, a: this.a1, v: vals }); return this; },
      setFormula: function (f) { record({ op: 'formula', s: this.sheet.name, a: this.a1, f: f }); return this; },
      clearContent: function () { record({ op: 'clear', s: this.sheet.name, a: this.a1 }); return this; },
      clear: function () { return this.clearContent(); },
      getRow: function () { return this.r0; }, getColumn: function () { return this.c0; }, getSheet: function () { return this.sheet; },
      getNumRows: function () { return this.r1 - this.r0 + 1; }, getNumColumns: function () { return this.c1 - this.c0 + 1; },
      getA1Notation: function () { return this.a1; },
      getDataValidation: function () { return null; }, clearDataValidations: function () { return this; }, setDataValidation: function () { return this; },
      insertCheckboxes: function () { return this; }, setBackground: function () { return this; }, setFontColor: function () { return this; },
      setFontWeight: function () { return this; }, setHorizontalAlignment: function () { return this; }, merge: function () { return this; },
      setRichTextValue: function (rt) { record({ op: 'value', s: this.sheet.name, a: this.a1, v: rt.text }); return this; },
      activate: function () { active = this.sheet.name; record({ op: 'activate', s: this.sheet.name, a: this.a1 }); return this; }
    };
    var sheets = {};
    (book.order || Object.keys(book.sheets)).forEach(function (name, i) {
      sheets[name] = {
        name: name, id: 100 + i,
        getName: function () { return name; }, getSheetId: function () { return 100 + i; }, getImages: function () { return []; },
        getRange: function (a, c, nr, nc) {
          if (typeof a === 'number') a = colLet(c) + a + ':' + colLet(c + (nc || 1) - 1) + (a + (nr || 1) - 1);
          return new Range(sheets[name], a);
        },
        getRangeList: function (list) { var sh = this; return { check: function () { list.forEach(function (x) { sh.getRange(x).setValue(true); }); }, uncheck: function () { list.forEach(function (x) { sh.getRange(x).setValue(false); }); } }; },
        createTextFinder: function () { return { matchFormulaText: function () { return this; }, replaceAllWith: function () { return 0; } }; },
        activate: function () { active = name; record({ op: 'activate', s: name }); return this; }
      };
    });
    var toasts = [];
    var ss = {
      getSheetByName: function (n) { return sheets[n] || null; }, getSheets: function () { return Object.keys(sheets).map(function (k) { return sheets[k]; }); },
      getSpreadsheetTimeZone: function () { try { return Intl.DateTimeFormat().resolvedOptions().timeZone; } catch (e) { return 'UTC'; } },
      getUrl: function () { return ''; }, getId: function () { return 'excel'; }, getName: function () { return 'Budget Planner with Mila'; },
      setActiveSheet: function (s) { active = s.name; record({ op: 'activate', s: s.name }); return s; },
      getActiveSheet: function () { return sheets[active]; }, toast: function (m) { toasts.push(m); }
    };
    function store(kind) {
      return {
        getProperty: function (k) { return props[kind][k] != null ? props[kind][k] : null; },
        setProperty: function (k, v) { props[kind][k] = String(v); propWrites.push({ kind: kind, k: k, v: String(v) }); return this; },
        deleteProperty: function (k) { delete props[kind][k]; propWrites.push({ kind: kind, k: k, v: null }); return this; }
      };
    }
    var env = {
      SpreadsheetApp: {
        getActive: function () { return ss; }, getActiveSpreadsheet: function () { return ss; }, getActiveSheet: function () { return sheets[active]; },
        openById: function () { return ss; }, flush: function () { },
        newRichTextValue: function () { var b = { text: '', url: '' }; return { setText: function (t) { b.text = t; return this; }, setLinkUrl: function (u) { b.url = u; return this; }, build: function () { return b; } }; },
        DataValidationCriteria: { CHECKBOX: 'CHECKBOX' },
        newDataValidation: function () { return { requireValueInList: function () { return this; }, setAllowInvalid: function () { return this; }, setHelpText: function () { return this; }, build: function () { return {}; } }; },
        getUi: function () { return { createMenu: function () { return { addItem: function () { return this; }, addSeparator: function () { return this; }, addToUi: function () { } }; }, showModalDialog: function () { }, showSidebar: function () { } }; }
      },
      Utilities: {
        formatDate: function (d, tz, f) {
          if (f === 'H') return String(d.getHours());
          if (f === 'yyyy-MM-dd') return ymd(d);
          return d.toISOString();
        },
        parseDate: function (s) { var p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
      },
      PropertiesService: { getUserProperties: function () { return store('user'); }, getDocumentProperties: function () { return store('doc'); }, getScriptProperties: function () { return store('doc'); } },
      LockService: { getDocumentLock: function () { return { waitLock: function () { }, releaseLock: function () { } }; } },
      ScriptApp: {   // Excel has no timers for add-ins: daily auto-fill runs when Mila's panel opens instead
        getProjectTriggers: function () { return []; }, deleteTrigger: function () { },
        newTrigger: function () { return { timeBased: function () { return this; }, everyDays: function () { return this; }, atHour: function () { return this; }, create: function () { throw new Error('no timers in Excel'); } }; },
        getService: function () { return { getUrl: function () { return null; } }; }
      },
      HtmlService: { createTemplateFromFile: function () { return { evaluate: function () { return { setTitle: function () { return this; }, addMetaTag: function () { return this; }, setWidth: function () { return this; } }; } }; },
        createHtmlOutputFromFile: function () { return { getContent: function () { return ''; } }; }, createHtmlOutput: function (h) { return { html: h, setWidth: function () { return this; }, setHeight: function () { return this; } }; } },
      Session: { getScriptTimeZone: function () { return ss.getSpreadsheetTimeZone(); } },
      console: (typeof console !== 'undefined') ? console : { log: function () { } }
    };
    return { env: env, writes: writes, propWrites: propWrites, props: props, toasts: toasts, active: function () { return active; }, readAfter: readAfter };
  }

  // ---------------------------------------------------------------- what to load (everything Mila reads)
  function engineRanges(M, withImp) {
    var out = ['A1:B' + (M.excel.scalarRows)], E = M.eng;
    var ansCols = Object.keys(E.ansCol).map(function (k) { return colNum(E.ansCol[k]); });
    out.push(colLet(Math.min.apply(null, ansCols)) + E.ans0 + ':' + colLet(Math.max.apply(null, ansCols)) + (E.ans0 + E.intents.length));
    var bc = Object.keys(E.bilCol).map(function (k) { return colNum(E.bilCol[k]); });
    out.push(colLet(Math.min.apply(null, bc)) + E.bil0 + ':' + colLet(Math.max.apply(null, bc)) + (E.bil0 + M.setup.nbill - 1));
    out.push(E.moptCol + E.mopt0 + ':' + E.moptCol + (E.mopt0 + E.nmopt - 1));
    if (M.imp && withImp !== false) {         // the Bank import part of the Engine: only when something was pasted there
      var ic = Object.keys(M.imp.eng).map(function (k) { return colNum(M.imp.eng[k]); });
      out.push('A' + M.imp.eng0 + ':' + colLet(Math.max.apply(null, ic)) + (M.imp.eng0 + M.imp.n - 1));
    }
    (M.ex && M.ex.check || []).forEach(function (c) { out.push(c[0]); });
    return out;
  }
  var SKIP_SHEETS = { 'Engine': 1 };

  // ---------------------------------------------------------------- Excel side (Office.js)
  function toBook(sheetData) {      // [{name, address, values, formulas, numberFormat}] -> book
    var book = { sheets: {}, order: [] };
    sheetData.forEach(function (sd) {
      var sh = book.sheets[sd.name] || (book.sheets[sd.name] = { cells: {} });
      if (book.order.indexOf(sd.name) < 0) book.order.push(sd.name);
      if (!sd.address) return;
      var addr = sd.address.split('!').pop(), rc = rangeRC(addr);
      for (var i = 0; i < sd.values.length; i++) for (var j = 0; j < sd.values[i].length; j++) {
        var v = sd.values[i][j], f = sd.formulas ? sd.formulas[i][j] : null, nf = sd.numberFormat ? sd.numberFormat[i][j] : null;
        if (typeof v === 'number' && isDateFormat(nf)) v = serialToDate(v);
        if (v === '' && (f == null || f === '')) continue;
        sh.cells[colLet(rc.c0 + j) + (rc.r0 + i)] = [v, typeof f === 'string' && f.charAt(0) === '=' ? f : null];
      }
    });
    return book;
  }
  // Excel on the web answers at most about 5 MB per request, so the planner is read in pieces (one page, or part of a page, at a time)
  function pieces(addr, maxCells) {
    var rc = rangeRC(addr), cols = rc.c1 - rc.c0 + 1, step = Math.max(1, Math.floor(maxCells / cols)), out = [];
    for (var r = rc.r0; r <= rc.r1; r += step) out.push(colLet(rc.c0) + r + ':' + colLet(rc.c1) + Math.min(rc.r1, r + step - 1));
    return out;
  }
  async function loadBook(Excel, M, engineOnly, onPart) {
    return Excel.run(async function (ctx) {
      var wsc = ctx.workbook.worksheets; wsc.load('items/name');
      var act = ctx.workbook.worksheets.getActiveWorksheet(); act.load('name');
      var impRow = M.eng.rows.impLines, impCell = null;
      if (impRow) { impCell = ctx.workbook.worksheets.getItem('Engine').getRange('B' + impRow); impCell.load('values'); }
      await ctx.sync();
      var impN = impCell ? (Number(impCell.values[0][0]) || 0) : 1;
      var data = wsc.items.map(function (ws) { return { name: ws.name }; });       // every page, in order (some stay empty)
      var targets = [], used = [];
      wsc.items.forEach(function (ws) {
        if (ws.name === 'Engine') {
          // Mila reads the Engine's numbers; only its first block (her settings and answers) needs the formulas too
          engineRanges(M, impN > 0).forEach(function (a, i) { targets.push({ name: ws.name, ws: ws, a: a, f: i === 0 }); });
        } else if (!engineOnly && !(ws.name === 'Bank import' && !impN)) {
          var u = ws.getUsedRangeOrNullObject(true); u.load('address'); used.push({ name: ws.name, ws: ws, u: u });
        }
      });
      if (used.length) await ctx.sync();
      used.forEach(function (x) { if (!x.u.isNullObject) targets.push({ name: x.name, ws: x.ws, a: x.u.address.split('!').pop(), f: true }); });
      // cut big pages into pieces, then send several pieces per request (Excel on the web answers about 5 MB at most)
      var parts = [];
      targets.forEach(function (t) { pieces(t.a, 12000).forEach(function (p) { var rc = rangeRC(p); parts.push({ t: t, a: p, n: (rc.r1 - rc.r0 + 1) * (rc.c1 - rc.c0 + 1) }); }); });
      var batches = [], cur = [], size = 0;
      parts.forEach(function (p) { if (cur.length && size + p.n > 30000) { batches.push(cur); cur = []; size = 0; } cur.push(p); size += p.n; });
      if (cur.length) batches.push(cur);
      for (var b = 0; b < batches.length; b++) {
        if (onPart) onPart(b + 1, batches.length);
        batches[b].forEach(function (p) { p.r = p.t.ws.getRange(p.a); p.r.load(p.t.f ? ['values', 'formulas', 'numberFormat'] : ['values', 'numberFormat']); });
        await ctx.sync();
        batches[b].forEach(function (p) { data.push({ name: p.t.name, address: p.a, values: p.r.values, formulas: p.t.f ? p.r.formulas : null, numberFormat: p.r.numberFormat }); });
      }
      var book = toBook(data);
      book.active = act.name;
      return book;
    });
  }
  function toCell(v) { return v instanceof Date ? dateToSerial(v) : (v == null ? '' : v); }
  async function applyWrites(Excel, writes) {
    if (!writes.length) return;
    return Excel.run(async function (ctx) {
      var book = ctx.workbook, act = null, touched = {};
      writes.forEach(function (w) { if (w.op !== 'activate') touched[w.s] = 1; });
      var prot = Object.keys(touched).map(function (n) { var p = book.worksheets.getItem(n).protection; p.load('protected'); return { n: n, p: p }; });
      await ctx.sync();
      prot.forEach(function (x) { x.was = !!x.p.protected; if (x.was) x.p.unprotect(); });
      for (var k = 0; k < writes.length; k++) {
        var w = writes[k];
        if (k && k % 250 === 0) await ctx.sync();
        if (w.op === 'activate') { act = w.s; continue; }
        var r = book.worksheets.getItem(w.s).getRange(w.a);
        if (w.op === 'clear') r.clear('Contents');
        else if (w.op === 'formula') r.formulas = [[w.f]];
        else if (w.op === 'value') {
          if (w.v instanceof Date) { r.values = [[dateToSerial(w.v)]]; r.numberFormat = [['mmm d, yyyy']]; }
          else r.values = [[w.v == null ? '' : w.v]];
        } else if (w.op === 'values') {
          var hasF = w.v.some(function (row) { return row.some(function (x) { return typeof x === 'string' && x.charAt(0) === '='; }); });
          var data = w.v.map(function (row) { return row.map(toCell); });
          if (hasF) r.formulas = data; else r.values = data;
        }
      }
      prot.forEach(function (x) { if (x.was) x.p.protect({ allowFormatColumns: true, allowFormatRows: true, allowAutoFilter: x.n === 'Transactions', allowSort: x.n === 'Transactions' }); });
      if (act) book.worksheets.getItem(act).activate();
      await ctx.sync();
    });
  }

  // ---------------------------------------------------------------- running Mila's code on a world
  var compiled = { code: null, names: null, fn: null };
  function loadMila(env, code) {
    var names = Object.keys(env);
    if (compiled.code === code && compiled.names === names.join(',')) return compiled.fn.apply(null, names.map(function (n) { return env[n]; }));
    var fn = new Function(names.join(','), code + '\n;return { milaChat: milaChat, milaHello: milaHello, milaPage: milaPage, milaWhere: milaWhere, milaExcelEdit: (typeof milaExcelEdit === "function" ? milaExcelEdit : null), milaExcelDaily: (typeof milaExcelDaily === "function" ? milaExcelDaily : null), MILA: MILA };');
    compiled = { code: code, names: names.join(','), fn: fn };
    return fn.apply(null, names.map(function (n) { return env[n]; }));
  }
  function seeded(seed) { var a = seed >>> 0; return function () { a = (a + 0x6D2B79F5) >>> 0; var t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
  function withRandom(seed, f) { var keep = Math.random; Math.random = seeded(seed); try { return f(); } finally { Math.random = keep; } }

  function messageCells(writes) {      // the last word Mila wrote in each message cell (Ask Mila Q2, Q3, Q4 and Q30 onwards)
    var out = {};
    writes.forEach(function (w) {
      if (w.s !== 'Ask Mila' || !w.a) return;
      var rc = rangeRC(w.a);
      for (var r = rc.r0; r <= rc.r1; r++) for (var c = rc.c0; c <= rc.c1; c++) {
        if (colLet(c) !== 'Q' || r === 1 || r === 5 || (r > 5 && r < 30)) continue;
        var v = w.op === 'clear' ? '' : w.op === 'values' ? w.v[r - rc.r0][c - rc.c0] : w.op === 'value' ? w.v : undefined;
        if (v !== undefined) out['Q' + r] = v == null ? '' : v;
      }
    });
    return out;
  }
  function messageFixes(w1, w2) {
    var a = messageCells(w1), b = messageCells(w2), fix = [];
    Object.keys(b).forEach(function (k) { if (!(k in a) || String(a[k]) !== String(b[k])) fix.push({ op: 'value', s: 'Ask Mila', a: k, v: b[k] }); });
    return fix;
  }

  /* run one call (milaChat, milaHello…) the Excel way: first pass -> write -> recalc -> second pass for the answer
     io = { load(engineOnly) -> book, apply(writes), props: {user, doc}, saveProps(propWrites) } */
  async function call(io, code, fnName, args) {
    var seed = Math.floor(Math.random() * 2147483647), props0 = JSON.parse(JSON.stringify(io.props));
    var S0 = await io.load(false);
    var w1 = makeWorld(S0, { props: JSON.parse(JSON.stringify(props0)) });
    var r1 = withRandom(seed, function () { var m = loadMila(w1.env, code); return m[fnName].apply(null, args); });
    var real = w1.writes.filter(function (w) { return w.op !== 'activate'; });
    if (!real.length) { if (w1.writes.length) await io.apply(w1.writes); await io.saveProps(w1.propWrites, w1.props); return { result: r1, writes: w1.writes, toasts: w1.toasts }; }
    await io.apply(w1.writes);
    await io.saveProps(w1.propWrites, w1.props);
    if (!w1.readAfter.engine) return { result: r1, writes: w1.writes, toasts: w1.toasts, passes: 1 };
    var S1 = await io.load(true);
    var w2 = makeWorld(S0, { props: JSON.parse(JSON.stringify(props0)), engineAfter: S1 });
    var r2;
    try { r2 = withRandom(seed, function () { var m = loadMila(w2.env, code); return m[fnName].apply(null, args); }); } catch (e) { r2 = r1; w2 = null; }
    // Mila's messages in the planner (Ask Mila, column Q) are written in the first pass, before Excel recalculated.
    // When the second pass says something different (the fresh numbers), put the second pass's words there instead.
    var fix = w2 ? messageFixes(w1.writes, w2.writes) : [];
    if (fix.length) await io.apply(fix);
    return { result: r2, writes: w1.writes.concat(fix), toasts: w1.toasts.concat(w2 ? w2.toasts : []), passes: 2, fixed: fix.length };
  }

  return { call: call, makeWorld: makeWorld, loadBook: loadBook, applyWrites: applyWrites, toBook: toBook, engineRanges: engineRanges,
    messageFixes: messageFixes, serialToDate: serialToDate, dateToSerial: dateToSerial, isDateFormat: isDateFormat, colLet: colLet, colNum: colNum, loadMila: loadMila };
})();
if (typeof module !== 'undefined') module.exports = MilaBridge;
