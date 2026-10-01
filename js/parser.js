/* Excel -> dashboard.json converter (runs in the browser, needs SheetJS) */
(function (root) {
  function DashParser(X) {
    var num = function (v) { return typeof v === 'number' && isFinite(v) ? v : 0; };
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
          if (!cur || f === '' || /^total/i.test(f) || /^total/i.test(cl(r[4]))) continue;
          var fd = { fund: f }; K.forEach(function (k2, j) { fd[k2] = rnd(num(r[9 + j])); }); cur.funds.push(fd);
        }
        heads = heads.filter(function (h) { return h.funds.length; });
        heads.forEach(function (h) { K.forEach(function (k2) { h[k2] = rnd(sum(h.funds, function (f2) { return f2[k2]; })); }); });
        var o = { code: s.c, name: SECN[s.c], heads: heads }; K.forEach(function (k2) { o[k2] = rnd(sum(heads, function (h) { return h[k2]; })); });
        return o;
      });
    }

    function works(g) {
      var cnt = {}, dc = 0, best = -1;
      g.forEach(function (r) { r.forEach(function (v, c) { if (typeof v === 'string' && /^\s*Total/.test(v)) { cnt[c] = (cnt[c] || 0) + 1; if (cnt[c] > best) { best = cnt[c]; dc = c; } } }); });
      var out = [], blk = [];
      g.forEach(function (r, i) {
        if (typeof r[dc] === 'string' && /^\s*Total/.test(r[dc])) {
          var rows = blk.map(function (j) { return g[j]; }).concat([r]); blk = [];
          var nm = cl(rows.map(function (x) { return x[4] == null ? '' : String(x[4]).trim(); }).filter(Boolean).join(' ')); if (!nm) return;
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
        var secs = [], sec = null;
        w.forEach(function (r) {
          var a = cl(r[0]);
          if (/Section$/.test(a)) { sec = { name: a.replace(' Section', ''), rows: [] }; secs.push(sec); }
          else if (sec && /^\d+(\.0)?$/.test(a) && r[1] != null && [2, 3, 4, 5].some(function (j) { return r[j] != null; }))
            sec.rows.push({ by: cl(r[1]), ob: num(r[2]), acc: num(r[3]), clr: num(r[4]), cb: num(r[5]), type: cl(r[6]).slice(0, 60) });
          else if (sec && cl(r[1]) === 'Total') sec.tot = { ob: num(r[2]), acc: num(r[3]), clr: num(r[4]), cb: num(r[5]) };
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
      for (var i = 2; i < g.length; i++) if (g[i][0] != null) out.push({ n: String(g[i][0]).trim(), ob: num(g[i][1]), acc: num(g[i][2]), tot: num(g[i][3]), clo: num(g[i][4]) + num(g[i][5]), cb: num(g[i][7]) });
      return out;
    }
    function parseRb(wb) {
      var g = grid(wb.Sheets[wb.SheetNames[0]]), out = [];
      for (var i = 2; i < g.length; i++) if (g[i][1] != null) out.push({ para: String(g[i][1]).replace(/\.0$/, ''), sec: cl(g[i][2]), since: dt(g[i][3]), rem: g[i][4] == null ? '' : String(g[i][4]), sect: g[i][5] == null ? '' : String(g[i][5]) });
      return out;
    }

    /* decide what a workbook contains */
    function detect(wb, fname) {
      var n = wb.SheetNames.map(function (s) { return s.trim(); });
      if (n.some(function (s) { return /^PH-Wise/i.test(s); })) return 'exp';
      if (n.some(function (s) { return /WEEKLY/i.test(s); })) return 'files';
      if (n.some(function (s) { return /^(GSU|CSTE)/i.test(s); }) || n.filter(function (s) { return /^PH-/.test(s); }).length > 2) return 'works';
      if (/inspect|board|rb/i.test(fname)) return 'rb';
      if (/audit/i.test(fname)) return 'audit';
      return null;
    }
    return { detect: detect, parseExp: parseExp, attachWorks: attachWorks, parseFiles: parseFiles, parseAudit: parseAudit, parseRb: parseRb };
  }
  if (typeof module !== 'undefined') module.exports = DashParser; else root.DashParser = DashParser;
})(this);
