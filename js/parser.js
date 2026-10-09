/* Excel -> dashboard.json converter (runs in the browser, needs SheetJS) */
(function (root) {
  function DashParser(X) {
    var num = function (v) { if (typeof v === 'string' && v.trim() !== '' && isFinite(+v.replace(/,/g, ''))) v = +v.replace(/,/g, ''); return typeof v === 'number' && isFinite(v) ? v : 0; };
    var rnd = function (x) { return Math.round(x * 100) / 100; };
    var cl = function (v) { return v == null ? '' : String(v).replace(/\s+/g, ' ').trim(); };
    var sum = function (a, f) { return a.reduce(function (t, x) { return t + f(x); }, 0); };
    function dt(v) {
      if (v instanceof Date) { var p = function (n) { return (n < 10 ? '0' : '') + n; }; return v.getFullYear() + '-' + p(v.getMonth() + 1) + '-' + p(v.getDate()); }
      return v == null ? '' : String(v).slice(0, 10);
    }
    function grid(ws) {
      var r = X.utils.decode_range(ws['!ref'] || 'A1'); r.s = { r: 0, c: 0 };
      return X.utils.sheet_to_json(ws, { header: 1, defval: null, range: r, raw: true });
    }
    function find(wb, re) { var n = wb.SheetNames.filter(function (s) { return re.test(s.trim()); })[0]; return n ? grid(wb.Sheets[n]) : null; }

    var NAMES = { 11: 'New Lines', 14: 'Gauge Conversion', 15: 'Doubling', 16: 'Traffic Facilities', 30: 'ROB/RUB', 32: 'Bridge Works', 33: 'S&T (Con)', 35: 'Electrification', 42: 'Workshops', 51: 'Staff Amenities', 53: 'Passenger Amenities', 64: 'Other Specified Works', 65: 'HRD / Training', 73: 'Capital Suspense', 29: 'Road Safety / LX' };
    var SECN = { A: 'CAO/Con', B: 'CSTE/Con', C: 'CSTE/Project', D: 'CEE/C/TRD & CAO/RSP (Gati Shakti)' };
    var K = ['org', 'voa', 're', 'fme', 'aug', 'sep', 'cum'];

    function parseLabels(wb) {
      var g = find(wb, /^PH-Wise/i); if (!g) return null;
      var m = null, u = null, re = /[A-Za-z]+'\d\d/;
      for (var i = 3; i < 10; i++) { var a = String((g[i] || [])[14] == null ? '' : g[i][14]), b = String((g[i] || [])[13] == null ? '' : g[i][13]); if (!m && re.test(a)) m = a.match(re)[0]; if (!u && re.test(b)) u = b.match(re)[0]; }
      return { month: m || 'Current month', upto: 'Upto ' + (u || ''), ly: 'Last year (same period)', lyFull: '2025-26' };
    }
    function parseExp(wb) {
      var g = find(wb, /^PH-Wise/i); if (!g) return null;
      var st = [], seen = {};
      g.forEach(function (r, i) { var m = String(r[7] == null ? '' : r[7]).trim().match(/^([A-D])(?:\s*[-\u2013]|\s+)/); if (m && !seen[m[1]]) { seen[m[1]] = 1; st.push({ i: i, c: m[1] }); } });
      return st.map(function (s, k) {
        var end = k + 1 < st.length ? st[k + 1].i : g.length, heads = [], cur = null;
        for (var i = s.i + 1; i < end; i++) {
          var r = g[i] || [], f = r[5] != null ? cl(r[5]) : cl(r[6]), v4 = r[4], ph = null;
          if (typeof v4 === 'number' && Number.isInteger(v4)) ph = v4; else if (typeof v4 === 'string' && /^\d+(\.0)?$/.test(v4.trim())) ph = parseInt(v4, 10);
          if (ph != null) { cur = heads.filter(function (h) { return h.ph === ph; })[0]; if (!cur) { cur = { ph: ph, name: NAMES[ph] || 'PH-' + ph, funds: [] }; heads.push(cur); } }
          if (f === '' && cur && ph != null && r[7] != null && [1, 3, 9, 10, 11, 12, 13, 14, 15].some(function (j) { return num(r[j]) !== 0; })) f = cl(r[7]);   /* head row without a fund label, e.g. CAP SUSP. */
          if (!cur || f === '' || /^total/i.test(f) || /^total/i.test(cl(r[4]))) continue;
          var fd = { fund: f }; K.forEach(function (k2, j) { fd[k2] = rnd(num(r[9 + j])); }); fd.ly = rnd(num(r[3])); fd.lyfy = rnd(num(r[1])); cur.funds.push(fd);   /* ly = same period last year, lyfy = full year 2025-26 */
        }
        heads = heads.filter(function (h) { return h.funds.length; });
        var LK = K.concat(['ly', 'lyfy']);
        heads.forEach(function (h) { LK.forEach(function (k2) { h[k2] = rnd(sum(h.funds, function (f2) { return f2[k2]; })); }); });
        var o = { code: s.c, name: SECN[s.c], heads: heads }; LK.forEach(function (k2) { o[k2] = rnd(sum(heads, function (h) { return h[k2]; })); });
        return o;
      });
    }

    /* summary lines that leak into the work list are dropped; a real work with a stray label in front keeps its name */
    var LABEL = /(grand\s*total|plan\s*head\s*\d+\s*totals?|sub\s*total|\btotals?\b|re-?appropriation|name\s+of\s+work)/ig;
    function junk(n) { return !n.replace(LABEL, ' ').replace(/[\s\W_]+/g, '') || /plan\s*head\s*\d+\s*totals?|grand\s*total/i.test(n); }
    function tidy(n) { return n.replace(/^((re-?appropriation|name\s+of\s+work)\s+)+/i, ''); }
    function works(g) {
      var cnt = {}, dc = 0, best = -1;
      g.forEach(function (r) { r.forEach(function (v, c) { if (typeof v === 'string' && /^\s*Total/.test(v)) { cnt[c] = (cnt[c] || 0) + 1; if (cnt[c] > best) { best = cnt[c]; dc = c; } } }); });
      var out = [], blk = [];
      g.forEach(function (r, i) {
        if (typeof r[dc] === 'string' && /^\s*Total/.test(r[dc])) {
          var rows = blk.map(function (j) { return g[j]; }).concat([r]); blk = [];
          var nm = cl(rows.map(function (x) { return x[4] == null ? '' : String(x[4]).trim(); }).filter(Boolean).join(' ')); if (!nm || junk(nm)) return; nm = tidy(nm);
          var gt = function (o) { return num(r[dc + o]); };
          var phs = []; rows.forEach(function (x) { if (typeof x[2] === 'number' && x[2] > 0 && x[2] < 100) phs.push(Math.trunc(x[2])); });
          var pbr = rows.filter(function (x) { return x[1] != null; })[0], fund = []; rows.forEach(function (x) { if (x[7] != null && fund.indexOf(String(x[7]).trim()) < 0) fund.push(String(x[7]).trim()); });
          var w = { name: nm.slice(0, 110), pb: pbr ? String(pbr[1]) : '', fund: fund.join(', '), sanc: rnd(sum(rows, function (x) { return num(x[8]); }) / 1e4), exp: rnd(sum(rows, function (x) { return num(x[9]); }) / 1e4), outlay: rnd(gt(1) / 1e4), m: [0, 1, 2, 3, 4, 5].map(function (k) { return rnd(gt(4 + k) / 1e4); }), yr: rnd(gt(16) / 1e4), ph: phs.length ? phs[0] : null };
          if (w.sanc || w.exp || w.yr || w.m.some(function (x) { return x; })) out.push(w);
        } else blk.push(i);
      });
      return out;
    }
    function attachWorks(wb, exp) {
      var wk = {}, cs = [];
      wb.SheetNames.forEach(function (n) {
        var s = n.trim();
        if (/^PH-/.test(s)) wk['A|' + s.slice(3).split(/\s+/)[0]] = works(grid(wb.Sheets[n]));
        else if (/^GSU/i.test(s)) wk['D|' + s.split(/\s+/)[1]] = works(grid(wb.Sheets[n]));
        else if (/^CSTE/i.test(s)) cs = works(grid(wb.Sheets[n]));
      });
      exp.forEach(function (s) { s.heads.forEach(function (h) {
        h.works = 'AD'.indexOf(s.code) >= 0 ? (wk[s.code + '|' + h.ph] || []) : s.code === 'C' ? cs.filter(function (w) { return w.ph === h.ph; }) : [];
        h.works.sort(function (a, b) { return sum(b.m, function (x) { return x; }) - sum(a.m, function (x) { return x; }); });
      }); });
    }

    function parseFiles(wb) {
      var w = find(wb, /WEEKLY/i), e = find(wb, /^Executive/i), res = {};
      if (w) {
        var secs = [], sec = null, last = '';
        w.forEach(function (r) {
          var a = cl(r[0]), nm = cl(r[1]), hasNum = [2, 3, 4, 5].some(function (j) { return r[j] != null; });
          if (/Section$/.test(a)) { sec = { name: a.replace(' Section', ''), rows: [] }; secs.push(sec); last = ''; }
          else if (sec && /^total$/i.test(nm)) sec.tot = { ob: num(r[2]), acc: num(r[3]), clr: num(r[4]), cb: num(r[5]) };
          else if (sec && hasNum && (/^\d+(\.0)?$/.test(a) || nm) && !/^by name$/i.test(nm)) {
            if (nm) last = nm; else nm = last;
            var ex = sec.rows.filter(function (x) { return x.by === nm; })[0], ty = cl(r[6]);
            if (ex) { ex.ob += num(r[2]); ex.acc += num(r[3]); ex.clr += num(r[4]); ex.cb += num(r[5]); if (ty && ex.type.indexOf(ty) < 0) ex.type = (ex.type + '; ' + ty).slice(0, 60); }
            else sec.rows.push({ by: nm, ob: num(r[2]), acc: num(r[3]), clr: num(r[4]), cb: num(r[5]), type: ty.slice(0, 60) });
          }
        });
        res.files = secs;
      }
      if (e) {
        var grp = '', p = [];
        e.forEach(function (r) {
          var a = cl(r[0]);
          if (r[2] == null && r[1] == null && a && a !== 'SL No.' && !/^\d+(\.0)?$/.test(a)) { grp = a; return; }
          if (/^\d+(\.0)?$/.test(a) && r[2] != null)
            p.push({ grp: grp, desc: cl(r[2]).slice(0, 120), file: cl(r[1]).slice(0, 40), dept: cl(r[3]), rcvd: dt(r[4]), name: cl(r[5]).toLowerCase(), status: cl(r[6]) });
        });
        res.pending = p;
      }
      return res;
    }

    function parseAudit(wb) {
      var g = grid(wb.Sheets[wb.SheetNames[0]]), out = [];
      for (var i = 2; i < g.length; i++) {
        var r = g[i], o = r[0] != null ? 0 : 1, n = r[o];
        if (typeof n !== 'string' || !n.trim()) continue;
        out.push({ n: n.trim(), ob: num(r[o + 1]), acc: num(r[o + 2]), tot: num(r[o + 3]), clo: num(r[o + 4]) + num(r[o + 5]), cb: num(r[o + 7]) });
      }
      return out;
    }
    function parseRb(wb) {
      var g = grid(wb.Sheets[wb.SheetNames[0]]), out = [];
      for (var i = 2; i < g.length; i++) if (g[i][1] != null) out.push({ para: String(g[i][1]).replace(/\.0$/, ''), sec: cl(g[i][2]), since: dt(g[i][3]), rem: g[i][4] == null ? '' : String(g[i][4]), sect: g[i][5] == null ? '' : String(g[i][5]) });
      return out;
    }


    /* ---------- simplified input template (one workbook, flat sheets) ---------- */
    var TPL = ['Expenditure', 'Works', 'FileSummary', 'PendingFiles', 'Audit', 'RB_Inspection', 'Info'];
    var nk = function (h) { return String(h == null ? '' : h).replace(/\(.*?\)/g, '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
    function tbl(wb, name) {
      var n = wb.SheetNames.filter(function (x) { return x.trim().toLowerCase() === name.toLowerCase(); })[0]; if (!n) return [];
      return X.utils.sheet_to_json(wb.Sheets[n], { defval: null, raw: true }).map(function (r) {
        var o = {}; Object.keys(r).forEach(function (k) { o[nk(k)] = r[k]; });
        o.v = function () { for (var i = 0; i < arguments.length; i++) if (o[arguments[i]] != null) return o[arguments[i]]; return null; };
        return o;
      }).filter(function (o) { return Object.keys(o).some(function (k) { return k !== 'v' && o[k] != null && String(o[k]).trim() !== ''; }); });
    }
    function parseTemplate(wb, D) {
      var notes = [], ex = tbl(wb, 'Expenditure'), wk = tbl(wb, 'Works'), fl = tbl(wb, 'FileSummary'), pf = tbl(wb, 'PendingFiles'), au = tbl(wb, 'Audit'), rb = tbl(wb, 'RB_Inspection'), info = tbl(wb, 'Info');
      var ia = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames.filter(function (x) { return x.trim().toLowerCase() === 'info'; })[0]] || {}, { header: 1, defval: null, raw: true });
      if (ia[0] && ia[0][1]) { D.asOn = dt(ia[0][1]).split('-').reverse().join('.'); notes.push('Position date'); }
      if (ex.length) {
        var by = {}, order = [];
        ex.forEach(function (r) {
          var c = String(r.v('section') || '').trim().toUpperCase().charAt(0), ph = parseInt(r.v('planheadno', 'planhead'), 10); if (!SECN[c] || isNaN(ph)) return;
          var s = by[c] || (by[c] = { code: c, name: SECN[c], heads: [] }); if (order.indexOf(c) < 0) order.push(c);
          var h = s.heads.filter(function (x) { return x.ph === ph; })[0] || (s.heads.push({ ph: ph, name: cl(r.v('planheadname')) || NAMES[ph] || 'PH-' + ph, funds: [], works: [] }), s.heads[s.heads.length - 1]);
          var fd = { fund: cl(r.v('fund')) || '-', org: rnd(num(r.v('org202627', 'org'))), voa: rnd(num(r.v('voa'))), re: rnd(num(r.v('re'))), fme: rnd(num(r.v('fme'))), aug: rnd(num(r.v('uptoaug'))), sep: rnd(num(r.v('sep'))), cum: rnd(num(r.v('cumulative', 'cum'))) };
          h.funds.push(fd);
        });
        var old = {}; (D.exp || []).forEach(function (s) { s.heads.forEach(function (h) { old[s.code + '|' + h.ph] = h.works || []; }); });
        D.exp = order.sort().map(function (c) {
          var s = by[c]; s.heads.forEach(function (h) { K.forEach(function (k) { h[k] = rnd(sum(h.funds, function (f) { return f[k]; })); }); h.works = old[c + '|' + h.ph] || []; });
          K.forEach(function (k) { s[k] = rnd(sum(s.heads, function (h) { return h[k]; })); }); return s;
        });
        notes.push('Expenditure');
      }
      if (wk.length && D.exp) {
        D.exp.forEach(function (s) { s.heads.forEach(function (h) { h.works = []; }); });
        wk.forEach(function (r) {
          var c = String(r.v('section') || '').trim().toUpperCase().charAt(0), ph = parseInt(r.v('planheadno', 'planhead'), 10);
          var h = (D.exp.filter(function (s) { return s.code === c; })[0] || { heads: [] }).heads.filter(function (x) { return x.ph === ph; })[0]; if (!h) return;
          var mo = ['apr', 'may', 'jun', 'jul', 'aug', 'sep'].map(function (k) { return rnd(num(r.v(k))); });
          h.works.push({ name: cl(r.v('workname', 'work')), pb: r.v('pbno') == null ? '' : String(r.v('pbno')), fund: cl(r.v('fund')), sanc: rnd(num(r.v('sanctionedcost'))), exp: rnd(num(r.v('expuptomar26'))), outlay: rnd(num(r.v('outlay202627', 'outlay'))), m: mo, yr: rnd(num(r.v('202627approx'))), ph: ph });
        });
        D.exp.forEach(function (s) { s.heads.forEach(function (h) { h.works.sort(function (a, b) { return sum(b.m, function (x) { return x; }) - sum(a.m, function (x) { return x; }); }); }); });
        notes.push('Works');
      }
      if (fl.length) {
        var secs = [];
        fl.forEach(function (r) {
          var n = cl(r.v('section')), s = secs.filter(function (x) { return x.name === n; })[0] || (secs.push({ name: n, rows: [], tot: { ob: 0, acc: 0, clr: 0, cb: 0 } }), secs[secs.length - 1]);
          var o = { by: cl(r.v('officer', 'byname')), ob: num(r.v('opening')), acc: num(r.v('accretion')), clr: num(r.v('clearance')), cb: num(r.v('closing')), type: cl(r.v('typeoffiles')).slice(0, 60) };
          s.rows.push(o); ['ob', 'acc', 'clr', 'cb'].forEach(function (k) { s.tot[k] += o[k]; });
        });
        D.files = secs; notes.push('File summary');
      }
      if (pf.length) {
        D.pending = pf.map(function (r) { return { grp: cl(r.v('group')), desc: cl(r.v('description')).slice(0, 120), file: cl(r.v('fileno')).slice(0, 40), dept: cl(r.v('department')), rcvd: dt(r.v('receivedon')), name: cl(r.v('officer')).toLowerCase(), status: cl(r.v('status')) }; });
        notes.push('Pending files');
      }
      if (au.length) {
        var run = { ob: 0, acc: 0, clo: 0 };
        D.audit = au.map(function (r) {
          var n = cl(r.v('nature')), o;
          if (/total/i.test(n)) o = { n: n, ob: run.ob, acc: run.acc, tot: run.ob + run.acc, clo: run.clo, cb: run.ob + run.acc - run.clo };
          else {
            var ob = num(r.v('opening')), ac = num(r.v('accrued')), cl2 = num(r.v('closed')), cb = r.v('closing') == null ? ob + ac - cl2 : num(r.v('closing'));
            o = { n: n, ob: ob, acc: ac, tot: ob + ac, clo: cl2, cb: cb }; run.ob += ob; run.acc += ac; run.clo += cl2;
          }
          return o;
        });
        notes.push('Audit');
      }
      if (rb.length) {
        D.rb = rb.map(function (r) { return { para: String(r.v('para')).replace(/\.0$/, ''), sec: cl(r.v('section')), since: dt(r.v('pendingsince')), rem: r.v('auditremarks') == null ? '' : String(r.v('auditremarks')), sect: r.v('sectionalremarks') == null ? '' : String(r.v('sectionalremarks')) }; });
        notes.push('RB inspection');
      }
      return notes;
    }

    /* decide what a workbook contains */
    function detect(wb, fname) {
      var n = wb.SheetNames.map(function (s) { return s.trim(); });
      if (n.filter(function (s) { return TPL.indexOf(s) >= 0; }).length >= 2) return 'template';
      if (n.some(function (s) { return /^PH-Wise/i.test(s); })) return 'exp';
      if (n.some(function (s) { return /WEEKLY/i.test(s); })) return 'files';
      if (n.some(function (s) { return /^(GSU|CSTE)/i.test(s); }) || n.filter(function (s) { return /^PH-/.test(s); }).length > 2) return 'works';
      if (/inspect|board|rb/i.test(fname)) return 'rb';
      if (/audit/i.test(fname)) return 'audit';
      return null;
    }
    return { detect: detect, parseExp: parseExp, parseLabels: parseLabels, attachWorks: attachWorks, parseFiles: parseFiles, parseAudit: parseAudit, parseRb: parseRb, parseTemplate: parseTemplate };
  }
  if (typeof module !== 'undefined') module.exports = DashParser; else root.DashParser = DashParser;
})(this);
