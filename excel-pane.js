/* Mila for Excel: the panel. Connects the chat (the same one as Mila Pro) to Excel through MilaBridge. */
(function () {
  'use strict';
  var X = window.MILA_EXCEL;                 // config, palettes, avatars (from mila-code.js)
  var CODE = window.MILA_CODE;               // Mila's code: the same as Mila Pro for Google Sheets
  var M = X.cfg;
  var props = { user: {}, doc: {} };
  var queue = Promise.resolve(), busy = 0, quietUntil = 0;
  // ---------------------------------------------------------------- what Mila is doing right now (shown in the panel)
  var diag = { step: 'starting', since: Date.now() };
  function chatEl() { return document.getElementById('chat'); }
  function statusLine(txt) {
    var chat = chatEl(); if (!chat) return;
    var el = document.getElementById('mila-status');
    if (!txt) { if (el && el.parentNode) el.parentNode.removeChild(el); return; }
    if (!el) { el = document.createElement('div'); el.id = 'mila-status'; el.className = 'msg mila typing'; el.appendChild(document.createElement('div')); el.children[0].className = 'bubble'; chat.appendChild(el); }
    el.children[0].textContent = txt; chat.scrollTop = chat.scrollHeight;
  }
  function step(name, txt) { diag.step = name; diag.since = Date.now(); if (txt) statusLine(txt); }
  function showErr(where, e) {
    var msg = (e && (e.message || e.description)) || String(e);
    if (e && e.debugInfo && e.debugInfo.errorLocation) msg += ' [' + e.debugInfo.errorLocation + ']';
    console.error('Mila', where, e);
    var chat = chatEl(); if (!chat) return;
    var m = document.createElement('div'); m.className = 'msg mila';
    var b = document.createElement('div'); b.className = 'bubble';
    b.style.borderColor = '#E9A0A0';
    b.textContent = '⚠️ Something went wrong while ' + where + ': ' + msg + ' (Please send a screenshot of this to the shop.)';
    m.appendChild(b); chat.appendChild(m); chat.scrollTop = chat.scrollHeight;
  }
  window.addEventListener('error', function (ev) { showErr(diag.step, (ev.message || 'error') + ' @ ' + String(ev.filename || '').split('/').pop() + ':' + ev.lineno); });
  window.addEventListener('unhandledrejection', function (ev) { showErr(diag.step, ev.reason); });
  setInterval(function () {             // a step that takes long says so, instead of looking frozen
    var s = Math.round((Date.now() - diag.since) / 1000);
    if (diag.step !== 'ready' && s >= 15 && document.getElementById('mila-status')) statusLine('Still working on: ' + diag.step + ' (' + s + ' s)…');
  }, 5000);
  function withTimeout(p, ms, what) {
    return new Promise(function (ok, fail) {
      var t = setTimeout(function () { fail(new Error('no answer from Excel after ' + Math.round(ms / 1000) + ' s (' + what + ')')); }, ms);
      p.then(function (v) { clearTimeout(t); ok(v); }, function (e) { clearTimeout(t); fail(e); });
    });
  }
  var NEED = ['Ask Mila', 'Today', 'Setup', 'Engine'];          // a workbook without these pages is not the planner
  var SETUP_KEY = 'mila_excel_setup', SETUP_VER = '2';

  function readUserProps() { try { return JSON.parse(localStorage.getItem('mila_user_props') || '{}'); } catch (e) { return {}; } }
  function readDocProps() { try { return JSON.parse(Office.context.document.settings.get('mila_props') || '{}'); } catch (e) { return {}; } }
  function saveSettings() { return new Promise(function (ok) { var t = setTimeout(ok, 4000); try { Office.context.document.settings.saveAsync(function () { clearTimeout(t); ok(); }); } catch (e) { clearTimeout(t); ok(); } }); }
  function saveProps(pw, all) {
    props = { user: all.user, doc: all.doc };
    try { localStorage.setItem('mila_user_props', JSON.stringify(props.user)); } catch (e) { }
    try { Office.context.document.settings.set('mila_props', JSON.stringify(props.doc)); } catch (e) { }
    return saveSettings();
  }
  var io = {
    get props() { return props; },
    load: function (engineOnly) { diag.step = engineOnly ? 'reading the new numbers' : 'reading your planner'; diag.since = Date.now(); return withTimeout(MilaBridge.loadBook(Excel, M, engineOnly), 60000, diag.step); },
    apply: function (writes) { quietUntil = Date.now() + 2500; diag.step = 'writing in your planner'; diag.since = Date.now(); return withTimeout(MilaBridge.applyWrites(Excel, writes), 60000, diag.step); },
    saveProps: saveProps
  };
  // one message at a time, in order
  function run(fn, args) {
    var job = queue.then(function () { busy++; return MilaBridge.call(io, CODE, fn, args); })
      .then(function (r) { busy--; quietUntil = Date.now() + 1500; diag.step = 'ready'; statusLine(null); return r.result; },
            function (e) {
              busy--; diag.step = 'ready'; statusLine(null);
              if (e && e.debugInfo && e.debugInfo.errorLocation && e.message) e.message += ' [' + e.debugInfo.errorLocation + ']';
              if (fn !== 'milaChat' && fn !== 'milaHello') showErr('answering', e);     // the chat shows its own message for these
              throw e;
            });
    queue = job.catch(function () { });
    return job;
  }
  // the chat panel calls google.script.run.withSuccessHandler(..).milaChat(text, page): same calls, answered here
  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          var p = name === 'milaWhere' ? where() : run(name, args);
          p.then(function (r) { if (ok) ok(r); }, function (e) { console.error(e); if (fail) fail(e); });
        };
      }
    });
  }
  function where() {     // which page is open (fast: no snapshot)
    return Excel.run(function (ctx) { var s = ctx.workbook.worksheets.getActiveWorksheet(); s.load('name'); return ctx.sync().then(function () { return { sheet: s.name }; }); });
  }
  function firstCell(a) { return String(a || '').split('!').pop().split(':')[0].replace(/\$/g, ''); }
  function rowOf(a) { return +(/(\d+)$/.exec(firstCell(a)) || [0, 0])[1]; }
  function colOf(a) { return (/^([A-Z]+)/.exec(firstCell(a)) || ['', ''])[1]; }

  // ---------------------------------------------------------------- is this the planner?
  function pageNames() {
    return Excel.run(function (ctx) { var w = ctx.workbook.worksheets; w.load('items/name'); return ctx.sync().then(function () { return w.items.map(function (s) { return s.name; }); }); });
  }

  // ---------------------------------------------------------------- once per planner: the same boxes as Mila Pro
  // Every page's "Ask me" button becomes a blue box you type in, the Today box talks to Mila too,
  // and the examples on Ask Mila become tap boxes (pick ✓ next to one, Mila answers or does it).
  function setupOnce(force) {
    return Excel.run(async function (ctx) {
      var book = ctx.workbook, ask = book.worksheets.getItem('Ask Mila');
      var q1 = ask.getRange('Q1'); q1.load('values');
      await ctx.sync();
      var done = false;
      try { done = Office.context.document.settings.get(SETUP_KEY) === SETUP_VER; } catch (e) { }
      if (!force && done && String(q1.values[0][0]) === 'pro') return false;
      var pal = M.palette, bar = M.bar || {}, names = Object.keys(bar).filter(function (n) { return bar[n].box; });
      var touched = names.concat(['Ask Mila']).filter(function (n, i, a) { return a.indexOf(n) === i; });
      var prot = touched.map(function (n) { var ws = book.worksheets.getItemOrNullObject(n); ws.load('name'); return { n: n, ws: ws }; });   // isNullObject is known after the sync
      await ctx.sync();
      prot = prot.filter(function (x) { return !x.ws.isNullObject; });
      prot.forEach(function (x) { x.p = x.ws.protection; x.p.load('protected'); });
      await ctx.sync();
      prot.forEach(function (x) { x.was = !!x.p.protected; if (x.was) x.p.unprotect(); });
      var dv = Office.context.requirements.isSetSupported('ExcelApi', '1.8');
      prot.forEach(function (x) {
        var n = x.n, b = bar[n];
        if (!b || !b.box) return;
        var rg = x.ws.getRange(b.rng);
        try { rg.clear('Hyperlinks'); } catch (e) { }
        x.ws.getRange(b.a).values = [['']];
        rg.format.fill.color = '#' + pal.input;
        rg.format.font.color = '#' + pal.input_ink;
        rg.format.font.bold = true;
        rg.format.font.underline = 'None';
        rg.format.horizontalAlignment = 'Left';
        rg.format.protection.locked = false;
        if (b.label) x.ws.getRange(b.label).values = [['💬 Talk to Mila · box →']];
        ask.getRange(b.reply).values = [[n === 'Today'
          ? '✏️ Type in the blue box above and press Enter: “spent 12 on lunch”, “paid rent”, or a question. (Mila for Excel, on a computer.)'
          : '✏️ Type in the blue box next to me and press Enter: “spent 12 on lunch”, “add bill Water 40 on the 5th”, or any question.']];
      });
      // tap boxes: pick ✓ next to a line
      var t0 = M.tap0, nt = (M.taps || []).length;
      if (nt) {
        ask.getRange('C' + t0 + ':C' + (t0 + nt - 1)).values = M.taps.map(function (x) { return [x]; });
        var boxes = ask.getRange('B' + t0 + ':B' + (t0 + nt - 1));
        boxes.values = M.taps.map(function () { return ['']; });
        boxes.format.protection.locked = false;
        boxes.format.horizontalAlignment = 'Center';
        boxes.format.font.color = '#' + pal.rose_deep;
        boxes.format.font.bold = true;
        boxes.format.fill.color = '#' + pal.input;
      }
      q1.values = [['pro']];
      await ctx.sync();
      // nice extras (newer Excel): a ✓ list in the tap boxes, and a hint when you click a blue box. If they fail, typing still works.
      var picks = false;
      if (dv) {
        try {
          if (nt) {
            var bx = ask.getRange('B' + t0 + ':B' + (t0 + nt - 1));
            bx.dataValidation.clear();
            bx.dataValidation.rule = { list: { inCellDropDown: true, source: '✓' } };
            bx.dataValidation.prompt = { showPrompt: true, title: 'Mila', message: 'Pick ✓ and I answer, or do it.' };
          }
          await ctx.sync(); picks = true;
        } catch (e) { console.log('Mila: no ✓ lists', e); }
        try {
          prot.forEach(function (x) {
            var b = bar[x.n]; if (!b || !b.box || x.n === 'Today') return;
            var rg = x.ws.getRange(b.rng);
            rg.dataValidation.clear();
            rg.dataValidation.prompt = { showPrompt: true, title: 'Talk to Mila', message: 'Type here, then press Enter: "spent 12 on lunch", "paid rent", or any question.' };
          });
          await ctx.sync();
        } catch (e) { console.log('Mila: no box hints', e); }
      }
      if (nt) ask.getRange('C' + (t0 - 1)).values = [[picks ? '👈 P I C K   ✓   I N   A   B O X :   I   A N S W E R ,   O R   D O   I T' : '👈 T Y P E   x   I N   A   B O X :   I   A N S W E R ,   O R   D O   I T']];
      prot.forEach(function (x) { if (x.was) x.p.protect({ allowFormatColumns: true, allowFormatRows: true, allowAutoFilter: x.n === 'Transactions', allowSort: x.n === 'Transactions' }); });
      await ctx.sync();
      return true;
    }).then(function (changed) {
      try { Office.context.document.settings.set(SETUP_KEY, SETUP_VER); } catch (e) { }
      return saveSettings().then(function () { return changed; });
    });
  }

  // ---------------------------------------------------------------- Mila in the background
  // After the first time, Mila starts with this planner (hidden): the boxes work and daily auto-fill runs, with or without her panel.
  var background = false;
  function startWithPlanner() {
    try {
      if (!Office.context.requirements.isSetSupported('SharedRuntime', '1.1') || !Office.addin) return Promise.resolve(false);
      background = true;
      return Promise.resolve(Office.addin.setStartupBehavior(Office.StartupBehavior.load)).then(function () { return true; }, function () { return false; });
    } catch (e) { return Promise.resolve(false); }
  }
  // Excel left open overnight: the next morning's bills and paychecks are logged by themselves (if daily auto-fill is on)
  var lastDay = new Date().toDateString();
  function dailyCheck() {
    var d = new Date().toDateString();
    if (d === lastDay) return;
    lastDay = d;
    run('milaExcelDaily', []).then(function (r) { if (r && r.said) pane.say('✨ ' + r.said); }, function () { });
  }
  // opening the panel hours later: a fresh hello instead of this morning's
  function onShown() {
    try {
      if (!Office.addin || !Office.addin.onVisibilityModeChanged) return;
      Office.addin.onVisibilityModeChanged(function (a) {
        if (a && a.visibilityMode === 'Taskpane' && window.milaSayHello && Date.now() - (window.milaHelloAt || 0) > 20 * 60000) window.milaSayHello();
      });
    } catch (e) { }
  }

  // ---------------------------------------------------------------- typing in the planner while the panel is open
  // the blue box on every page, the big box on Ask Mila, the Today box, and the tap boxes
  function watched() {
    var map = {}, bar = M.bar || {};
    Object.keys(bar).forEach(function (n) { if (bar[n].box) (map[n] = map[n] || {})[firstCell(bar[n].a)] = 'box'; });
    (map['Ask Mila'] = map['Ask Mila'] || {})[firstCell(M.ask)] = 'ask';
    for (var i = 0; i < (M.taps || []).length; i++) map['Ask Mila']['B' + (M.tap0 + i)] = 'tap';
    return map;
  }
  function mine(ev) {          // a change Mila made herself (not you typing)
    if (ev.triggerSource) return ev.triggerSource === 'ThisLocalAddin';
    return Date.now() < quietUntil;
  }
  function watchBoxes() {
    var map = watched();
    return Excel.run(async function (ctx) {
      var all = ctx.workbook.worksheets; all.load('items/name');
      await ctx.sync();
      var have = all.items.map(function (x) { return x.name; });
      Object.keys(map).filter(function (n) { return have.indexOf(n) >= 0; }).forEach(function (n) {
        var ws = ctx.workbook.worksheets.getItem(n);
        ws.onChanged.add(function (ev) {
          if (mine(ev)) return Promise.resolve();
          var a = firstCell(ev.address), kind = map[n][a];
          if (!kind && n === 'Ask Mila' && colOf(ev.address) === 'B') {          // pasting over several tap boxes: the first one counts
            var r0 = rowOf(ev.address); if (map[n]['B' + r0] === 'tap') { a = 'B' + r0; kind = 'tap'; }
          }
          if (!kind) return Promise.resolve();
          return Excel.run(function (c2) {
            var r = c2.workbook.worksheets.getItem(n).getRange(a); r.load('values');
            return c2.sync().then(function () {
              var v = String(r.values[0][0] == null ? '' : r.values[0][0]).trim();
              if (!v) return;
              var tapText = kind === 'tap' ? null : v;
              return run('milaExcelEdit', [n, a, kind === 'tap' ? 'TRUE' : v]).then(function (r) { pane.note(kind, n, tapText, r && r.said); });
            });
          });
        });
      });
      return ctx.sync();
    }).catch(function (e) { console.log('Mila: boxes not watched', e); });
  }
  // the chat shows what happened in the planner, so the panel and the boxes tell one story
  var pane = { say: function (txt) {
    var chat = document.getElementById('chat'); if (!chat) return;
    var m = document.createElement('div'); m.className = 'msg mila';
    var b = document.createElement('div'); b.className = 'bubble'; b.textContent = txt;
    m.appendChild(b); chat.appendChild(m); chat.scrollTop = chat.scrollHeight;
  }, note: function (kind, sheet, text, said) {
    var chat = document.getElementById('chat'); if (!chat) return;
    function bubble(who, txt) {
      var m = document.createElement('div'); m.className = 'msg ' + who;
      var b = document.createElement('div'); b.className = 'bubble'; b.textContent = txt;
      m.appendChild(b); chat.appendChild(m);
    }
    if (text) bubble('me', '📝 ' + sheet + ': ' + text);
    if (said || !(kind === 'tap' || kind === 'ask')) {
      bubble('mila', said || 'My answer is at the top of ' + sheet + '. 🦋'); chat.scrollTop = chat.scrollHeight; return;
    }
    // a question: the planner's own formulas answer it on Ask Mila (C14), so read it once Excel has recalculated
    Excel.run(function (ctx) { var r = ctx.workbook.worksheets.getItem('Ask Mila').getRange('C14'); r.load('values'); return ctx.sync().then(function () { return r.values[0][0]; }); })
      .then(function (v) { bubble('mila', String(v || '') || 'My answer is on Ask Mila. 🦋'); chat.scrollTop = chat.scrollHeight; },
            function () { bubble('mila', 'My answer is on Ask Mila. 🦋'); });
  } };

  function theme() {
    return Excel.run(function (ctx) {
      var r = ctx.workbook.worksheets.getItem('Welcome').getRange('B40:B60'); r.load('values');
      return ctx.sync().then(function () { return r.values.some(function (x) { return /dark/i.test(String(x[0])); }) ? 'dark' : 'light'; });
    }).catch(function () { return 'light'; });
  }
  function notPlanner() {
    var p = X.palettes.light;
    document.body.innerHTML = '<div style="font:14px/1.55 Segoe UI,Arial,sans-serif;color:#' + p.ink + ';background:#' + p.bg + ';min-height:100vh;padding:22px;box-sizing:border-box">' +
      '<img alt="" src="data:image/png;base64,' + X.avatars.light + '" style="width:64px;height:64px;border-radius:50%;display:block;margin-bottom:12px">' +
      '<p style="font-size:16px;margin:0 0 10px"><b>Hi, I’m Mila 🦋</b></p>' +
      '<p>I live inside the <b>Budget Planner with Mila</b> workbook, and this workbook isn’t it.</p>' +
      '<p>Open your planner (the .xlsx file from your download), then click <b>Talk to Mila</b> on the Home tab.</p>' +
      (X.shop ? '<p>Don’t have the planner yet? <a target="_blank" rel="noopener" style="color:#' + p.rose_deep + '" href="' + X.shop + '">Get the Budget Planner with Mila</a>.</p>' : '') +
      '<p style="color:#' + p.muted + ';font-size:12px">Nothing in this workbook was read or changed.</p></div>';
  }

  step('connecting to Excel', 'Connecting to Excel…');
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { document.body.innerHTML = '<p style="padding:16px">Mila works in Excel.</p>'; return; }
    step('checking your workbook', 'Checking your workbook…');
    withTimeout(pageNames(), 30000, 'listing the pages').then(function (names) {
      if (NEED.some(function (n) { return names.indexOf(n) < 0; })) { notPlanner(); return; }
      props = { user: readUserProps(), doc: readDocProps() };
      window.google = { script: {} };
      Object.defineProperty(window.google.script, 'run', { get: function () { return runner(null, null); } });
      return withTimeout(theme(), 15000, 'reading the theme').catch(function () { return 'light'; }).then(function (t) {
        window.MILA_PANEL_DATA = { palette: X.palettes[t], avatar: X.avatars[t], theme: t, version: M.version, app: false };
        // the planner's boxes are made first (once), then Mila says hello: both wait in the same line, in order
        queue = queue.then(function () {
          step('getting your planner ready', 'Getting your planner ready (only the first time)…');
          return withTimeout(setupOnce(false), 45000, 'getting your planner ready').catch(function (e) { showErr('getting your planner ready', e); });
        }).then(function () { step('reading your planner', 'Reading your planner…'); });
        window.milaStartPanel();
        watchBoxes();
        startWithPlanner();
        onShown();
        setInterval(dailyCheck, 10 * 60000);
      });
    }).catch(function (e) { showErr(diag.step, e); });
  });
  window.MilaExcelPane = { setupOnce: setupOnce, watched: watched, dailyCheck: dailyCheck, setDay: function (d) { lastDay = d; }, isBackground: function () { return background; } };
})();
