$(function () {
  var D, charts = {}, PAL = ['#1f6feb', '#178a63', '#d98a0b', '#c9423f', '#7b57c9', '#12a4c4', '#e0702a', '#6b7a90'];
  var PIE = ['#1f6feb', '#178a63', '#d98a0b', '#c9423f', '#7b57c9', '#12a4c4', '#e0702a', '#6b7a90', '#d6336c', '#8a9a1b', '#0b7285', '#5f3dc4', '#a0522d', '#2f9e44'];
  var MON = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'];
  var f2 = function (v) { return (v || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  var f0 = function (v) { return (v || 0).toLocaleString('en-IN'); };
  var pct = function (a, b) { return b ? Math.round(a / b * 1000) / 10 : 0; };
  var sum = function (a, k) { return a.reduce(function (t, x) { return t + (k ? x[k] : x); }, 0); };
  var esc = function (s) { return $('<i>').text(s == null ? '' : s).html(); };
  var clip = function (s, n) { return s.length > n ? s.slice(0, n - 2) + '..' : s; };
  Chart.defaults.font.family = '"Segoe UI",system-ui,Arial,sans-serif';
  Chart.defaults.plugins.legend.position = 'bottom';
  Chart.defaults.plugins.legend.labels.usePointStyle = true;

  $.getJSON('data/dashboard.json').done(function (j) { D = j; init(); }).fail(function () {
    $('#loadErr').removeClass('d-none').html('<strong>Could not load data/dashboard.json.</strong> Browsers block JSON loading from a file:// address. Start the bundled <code>run.bat</code> (Windows) or <code>run.sh</code> and open <code>http://localhost:8080</code>, or host the folder on any web server (IIS, Apache, nginx).');
  });

  function draw(id, cfg, click) {
    if (charts[id]) charts[id].destroy();
    cfg.options = $.extend(true, { responsive: true, maintainAspectRatio: false }, cfg.options || {});
    if (click) {
      cfg.options.onClick = function (e, els) { if (els.length) click(els[0].index, els[0].datasetIndex); };
      cfg.options.onHover = function (e, els) { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; };
    }
    charts[id] = new Chart(document.getElementById(id), cfg);
  }
  function bar(p) { return '<div class="progress"><div class="progress-bar" data-w="' + Math.min(p, 100) + '"></div></div>'; }
  function tbl(sel, head, rows, click) {
    $(sel).html('<thead><tr>' + head + '</tr></thead><tbody>' + rows.map(function (r, i) { return '<tr data-i="' + i + '">' + r + '</tr>'; }).join('') + '</tbody>');
    $(sel).find('.progress-bar').each(function () { this.style.width = $(this).data('w') + '%'; });
    $(sel).find('tbody tr').on('click', function () { click($(this).data('i')); });
  }
  function kpiCard(k) { return '<div class="col-6 col-lg-4 col-xl-2"><div class="card kpi ' + (k.c || '') + (k.go ? '' : ' static') + '" data-k="' + k.id + '"><div class="lbl">' + k.l + '</div><div class="val">' + k.v + '</div><div class="note">' + (k.n || '&nbsp;') + '</div></div></div>'; }

  /* ---------- drill engine ---------- */
  var stack = [], modal;
  function open(trail) { stack = trail; render(); modal = bootstrap.Modal.getOrCreateInstance('#drill'); modal.show(); }
  function push(label, fn) { stack.push({ l: label, f: fn }); render(); }
  function render() {
    var top = stack[stack.length - 1], v = top.f();
    $('#dTitle').text(v.title || top.l);
    $('#dCrumb').html(stack.map(function (s, i) { return i === stack.length - 1 ? '<li class="breadcrumb-item active">' + esc(s.l) + '</li>' : '<li class="breadcrumb-item"><a data-i="' + i + '">' + esc(s.l) + '</a></li>'; }).join(''))
      .find('a').on('click', function () { stack = stack.slice(0, $(this).data('i') + 1); render(); });
    $('#dKpi').html((v.kpis || []).map(function (k) { return '<div class="col-6 col-md-3"><div class="card kpi static p-2 ' + (k[2] || '') + '"><div class="lbl">' + k[0] + '</div><div class="fw-bold fs-6">' + k[1] + '</div></div></div>'; }).join(''));
    $('#dChartCol').toggle(!!v.chart); $('#dTblCol').toggle(!!v.rows);
    $('#dText').html(v.text || '');
    if (charts.dChart) { charts.dChart.destroy(); delete charts.dChart; }
    if (v.chart) setTimeout(function () { draw('dChart', v.chart, v.onChart); }, 60);
    if (v.rows) tbl('#dTable', v.head, v.rows, v.onRow || function () { });
    $('#dTable').toggleClass('nohover', !v.onRow);
  }

  /* ---------- expenditure views ---------- */
  function vSections() {
    var S = D.exp;
    return { title: 'Expenditure - Sections', kpis: [['ORG 2026-27', f2(sum(S, 'org'))], ['Cumulative', f2(sum(S, 'cum')), 'g'], ['% on ORG', pct(sum(S, 'cum'), sum(S, 'org')) + '%', 'a'], ['Sep\'26', f2(sum(S, 'sep'))]],
      chart: barCfg(S.map(function (s) { return s.code + ' ' + s.name; }), [['ORG', S.map(function (s) { return s.org; })], ['Cumulative', S.map(function (s) { return s.cum; })]]),
      onChart: function (i) { push(S[i].name, function () { return vHeads(S[i]); }); },
      head: '<th>Section</th><th class="num">ORG</th><th class="num">Cum.</th><th>% ORG</th>',
      rows: S.map(function (s) { return '<td>' + esc(s.code + ' - ' + s.name) + '</td><td class="num">' + f2(s.org) + '</td><td class="num">' + f2(s.cum) + '</td><td>' + bar(pct(s.cum, s.org)) + '<small>' + pct(s.cum, s.org) + '%</small></td>'; }),
      onRow: function (i) { push(S[i].name, function () { return vHeads(S[i]); }); } };
  }
  function vHeads(s) {
    var H = s.heads;
    return { title: s.name + ' - Plan Heads', kpis: [['ORG', f2(s.org)], ['Cumulative', f2(s.cum), 'g'], ['% on ORG', pct(s.cum, s.org) + '%', 'a'], ['Plan heads', H.length]],
      chart: barCfg(H.map(function (h) { return 'PH-' + h.ph; }), [['ORG', H.map(function (h) { return h.org; })], ['Cumulative', H.map(function (h) { return h.cum; })]]),
      onChart: function (i) { push('PH-' + H[i].ph, function () { return vWorks(H[i]); }); },
      head: '<th>Plan head</th><th class="num">ORG</th><th class="num">RE(P)</th><th class="num">Cum.</th><th>% ORG</th>',
      rows: H.map(function (h) { return '<td>PH-' + h.ph + ' ' + esc(h.name) + '</td><td class="num">' + f2(h.org) + '</td><td class="num">' + f2(h.re) + '</td><td class="num">' + f2(h.cum) + '</td><td>' + bar(pct(h.cum, h.org)) + '<small>' + pct(h.cum, h.org) + '%</small></td>'; }),
      onRow: function (i) { push('PH-' + H[i].ph, function () { return vWorks(H[i]); }); } };
  }
  function vWorks(h) {
    var W = h.works, top = W.slice(0, 10);
    return { title: 'PH-' + h.ph + ' ' + h.name + ' - Works', kpis: [['ORG', f2(h.org)], ['Cumulative', f2(h.cum), 'g'], ['Fund heads', h.funds.length], ['Works listed', W.length]],
      chart: { type: 'bar', data: { labels: top.map(function (w) { return clip(w.name, 28); }), datasets: [{ label: 'Apr-Sep\'26 (Rs. Cr)', data: top.map(function (w) { return +sum(w.m).toFixed(2); }), backgroundColor: PIE[1] }] }, options: { indexAxis: 'y', plugins: { legend: { display: false } } } },
      onChart: function (i) { push(clip(top[i].name, 24), function () { return vWork(top[i]); }); },
      head: '<th>Work</th><th class="num">Sanct.</th><th class="num">Outlay</th><th class="num">Apr-Sep</th>',
      rows: W.length ? W.map(function (w) { return '<td>' + esc(clip((w.pb ? w.pb + ' | ' : '') + w.name, 70)) + '</td><td class="num">' + f2(w.sanc) + '</td><td class="num">' + f2(w.outlay) + '</td><td class="num">' + f2(sum(w.m)) + '</td>'; }) : ['<td class="text-muted">No work-wise data for this head</td>'],
      onRow: function (i) { if (W[i]) push(clip(W[i].name, 24), function () { return vWork(W[i]); }); } };
  }
  function vWork(w) {
    return { title: (w.pb ? w.pb + ' - ' : '') + w.name, kpis: [['Fund', esc(w.fund || '-')], ['Sanctioned cost', f2(w.sanc)], ['Exp. to Mar\'26', f2(w.exp)], ['Outlay 2026-27', f2(w.outlay), 'a'], ['Apr-Sep\'26', f2(sum(w.m)), 'g'], ['2026-27 approx.', f2(w.yr)]],
      chart: { type: 'bar', data: { labels: MON, datasets: [{ label: 'Monthly expenditure (Rs. Cr)', data: w.m, backgroundColor: PIE[0] }] }, options: { plugins: { legend: { display: false } } } } };
  }
  function barCfg(lab, ds) { return { type: 'bar', data: { labels: lab.map(function (l) { return clip(l, 26); }), datasets: ds.map(function (d, i) { return { label: d[0], data: d[1], backgroundColor: PIE[i] }; }) } }; }

  /* ---------- file views ---------- */
  function key(nm) {
    nm = (nm || '').toLowerCase().replace(/\(.*?\)/g, '').trim();
    if (/^t[\. ]/.test(nm) || nm.indexOf('tribhuwan') === 0) return 'tribhuwan';
    return nm.replace('deepankar', 'dipankar').replace('gagan kr', 'gagan').split(/[\s\/]+/)[0];
  }
  var FS = function () { return D.files.filter(function (s) { return s.tot; }); };
  function moveCfg(lab, T) { return barCfg(lab, [['Opening', T.map(function (t) { return t.ob; })], ['Accretion', T.map(function (t) { return t.acc; })], ['Clearance', T.map(function (t) { return t.clr; })], ['Closing', T.map(function (t) { return t.cb; })]]); }
  function vFileSecs() {
    var S = FS(), T = S.map(function (s) { return s.tot; }), go = function (i) { push(S[i].name, function () { return vOfficers(S[i]); }); };
    return { title: 'File pendency - Sections', kpis: [['Opening', f0(sum(T, 'ob'))], ['Accretion', f0(sum(T, 'acc')), 'a'], ['Clearance', f0(sum(T, 'clr')), 'g'], ['Closing', f0(sum(T, 'cb')), 'r']],
      chart: moveCfg(S.map(function (s) { return s.name; }), T), onChart: go,
      head: '<th>Section</th><th class="num">Open</th><th class="num">Accr.</th><th class="num">Clear</th><th class="num">Close</th>',
      rows: S.map(function (s) { return '<td>' + esc(s.name) + '</td><td class="num">' + s.tot.ob + '</td><td class="num">' + s.tot.acc + '</td><td class="num">' + s.tot.clr + '</td><td class="num fw-bold">' + s.tot.cb + '</td>'; }), onRow: go };
  }
  function vOfficers(s) {
    var R = s.rows, go = function (i) { push(R[i].by.split('(')[0].trim(), function () { return vFiles(R[i]); }); };
    return { title: s.name + ' Section - officer-wise', kpis: [['Opening', f0(s.tot.ob)], ['Accretion', f0(s.tot.acc), 'a'], ['Clearance', f0(s.tot.clr), 'g'], ['Closing', f0(s.tot.cb), 'r']],
      chart: moveCfg(R.map(function (r) { return r.by.replace(/\(.*?\)/g, '').trim(); }), R), onChart: go,
      head: '<th>Officer</th><th class="num">Open</th><th class="num">Accr.</th><th class="num">Clear</th><th class="num">Close</th><th>Type of files</th>',
      rows: R.map(function (r) { return '<td>' + esc(r.by) + '</td><td class="num">' + r.ob + '</td><td class="num">' + r.acc + '</td><td class="num">' + r.clr + '</td><td class="num fw-bold">' + r.cb + '</td><td class="small">' + esc(r.type) + '</td>'; }), onRow: go };
  }
  function vFiles(r) {
    var k = key(r.by), L = D.pending.filter(function (p) { return key(p.name) === k; }), g = {};
    L.forEach(function (p) { g[p.grp || 'Other'] = (g[p.grp || 'Other'] || 0) + 1; });
    return { title: 'Pending files - ' + r.by, kpis: [['Opening', r.ob], ['Accretion', r.acc, 'a'], ['Clearance', r.clr, 'g'], ['Closing', r.cb, 'r']],
      chart: L.length ? { type: 'doughnut', data: { labels: Object.keys(g), datasets: [{ data: Object.keys(g).map(function (x) { return g[x]; }), backgroundColor: PIE }] } } : null,
      head: '<th>Description</th><th>Dept.</th><th>Received</th><th>Status</th>',
      rows: L.length ? L.map(function (p) { return '<td>' + esc(p.desc) + '<div class="small text-muted">' + esc(p.file) + '</div></td><td>' + esc(p.dept) + '</td><td>' + esc(p.rcvd) + '</td><td>' + esc(p.status) + '</td>'; }) : ['<td class="text-muted">No file-level list available for this officer.</td>'] };
  }

  /* ---------- audit / RB views ---------- */
  var isTot = function (n) { return /Total/.test(n); };
  function vAudit(a) {
    return { title: a.n, kpis: [['Opening 01.04.26', a.ob], ['Accrued Apr-Sep\'26', a.acc, 'a'], ['Total', a.tot], ['Closed', a.clo, 'g'], ['Closing 30.09.26', a.cb, 'r']],
      chart: { type: 'bar', data: { labels: ['Opening', 'Accrued', 'Closed', 'Closing'], datasets: [{ data: [a.ob, a.acc, a.clo, a.cb], backgroundColor: [PAL[7], PAL[2], PAL[1], PAL[3]] }] }, options: { plugins: { legend: { display: false } } } } };
  }
  function vRb(r) {
    return { title: 'Inspection Report Part-I - Para ' + r.para, kpis: [['Para', r.para], ['Section', esc(r.sec)], ['Pending since', r.since, 'r']],
      text: '<h6>Audit remarks</h6><p>' + esc(r.rem) + '</p>' + (r.sect ? '<h6>Sectional remarks</h6><p>' + esc(r.sect) + '</p>' : '<p class="text-muted">No sectional remarks recorded.</p>') };
  }

  /* ---------- dashboard ---------- */
  function init() {
    $('#asOn').text(D.asOn);
    var S = D.exp, T = FS().map(function (s) { return s.tot; }), A = D.audit, at = A[A.length - 1];
    var all = []; S.forEach(function (s) { s.heads.forEach(function (h) { h.works.forEach(function (w) { all.push({ w: w, h: h, s: s }); }); }); });
    all.sort(function (a, b) { return sum(b.w.m) - sum(a.w.m); });
    var K = [
      { id: 'org', l: 'ORG 2026-27', v: f2(sum(S, 'org')), n: 'Rs. crore', go: vSections },
      { id: 'cum', l: 'Cumulative exp.', v: f2(sum(S, 'cum')), n: pct(sum(S, 'cum'), sum(S, 'org')) + '% of ORG', c: 'g', go: vSections },
      { id: 'sep', l: 'Sep\'26 to date', v: f2(sum(S, 'sep')), n: 'Rs. crore', c: 'a', go: vSections },
      { id: 'fil', l: 'Files pending', v: f0(sum(T, 'cb')), n: 'closing balance', c: 'r', go: vFileSecs },
      { id: 'aud', l: 'Audit outstanding', v: f0(at.cb), n: 'paras & notes', c: 'a', go: function () { return vAudit(at); } },
      { id: 'rb', l: 'RB Inspection paras', v: D.rb.length, n: 'pending', c: 'r', go: function () { return vRbList(); } }];
    $('#kpis').html(K.map(kpiCard).join('')).find('.kpi').not('.static').on('click', function () {
      var k = K.filter(function (x) { return x.id === $(this).data('k'); }.bind(this))[0]; open([{ l: k.l, f: k.go }]);
    });
    function vRbList() { return { title: 'RB Inspection Report - pending paras', kpis: [['Paras', D.rb.length, 'r']], head: '<th>Para</th><th>Section</th><th>Since</th>', rows: D.rb.map(function (r) { return '<td>' + esc(r.para) + '</td><td>' + esc(r.sec) + '</td><td>' + esc(r.since) + '</td>'; }), onRow: function (i) { push('Para ' + D.rb[i].para, function () { return vRb(D.rb[i]); }); } }; }

    draw('cSec', barCfg(S.map(function (s) { return s.code + ' ' + s.name.split(' ')[0]; }), [['ORG', S.map(function (s) { return s.org; })], ['Cumulative', S.map(function (s) { return s.cum; })]]),
      function (i) { open([{ l: 'Sections', f: vSections }, { l: S[i].name, f: function () { return vHeads(S[i]); } }]); });
    var HP = []; S.forEach(function (s) { s.heads.forEach(function (h) { if (h.cum > 0) HP.push({ s: s, h: h }); }); });
    HP.sort(function (a, b) { return b.h.cum - a.h.cum; });
    draw('cPh', { type: 'bar', data: { labels: HP.map(function (x) { return 'PH-' + x.h.ph + ' ' + x.s.code; }), datasets: [{ label: 'Cumulative (Rs. Cr)', data: HP.map(function (x) { return x.h.cum; }), backgroundColor: HP.map(function (x) { return PAL[x.s.code.charCodeAt(0) - 65]; }) }] }, options: { plugins: { legend: { display: false } } } },
      function (i) { var x = HP[i]; open([{ l: 'Sections', f: vSections }, { l: x.s.name, f: function () { return vHeads(x.s); } }, { l: 'PH-' + x.h.ph, f: function () { return vWorks(x.h); } }]); });
    var TW = all.slice(0, 8);
    tbl('#tWorks', '<th>Work</th><th>Section / PH</th><th class="num">Sanct. cost</th><th class="num">Outlay 26-27</th><th class="num">Apr-Sep\'26</th><th>Utilisation</th>', TW.map(function (x) {
      return '<td>' + esc(clip(x.w.name, 75)) + '</td><td><span class="pill">' + esc(x.s.code) + ' / PH-' + x.h.ph + '</span></td><td class="num">' + f2(x.w.sanc) + '</td><td class="num">' + f2(x.w.outlay) + '</td><td class="num">' + f2(sum(x.w.m)) + '</td><td>' + bar(pct(sum(x.w.m), x.w.outlay)) + '</td>'; }),
      function (i) { var x = TW[i]; open([{ l: x.s.name, f: function () { return vHeads(x.s); } }, { l: 'PH-' + x.h.ph, f: function () { return vWorks(x.h); } }, { l: 'Work', f: function () { return vWork(x.w); } }]); });

    var FSL = FS();
    draw('cFile', moveCfg(FSL.map(function (s) { return s.name; }), T), function (i) { open([{ l: 'Sections', f: vFileSecs }, { l: FSL[i].name, f: function () { return vOfficers(FSL[i]); } }]); });
    tbl('#tFile', '<th>Section</th><th class="num">Open</th><th class="num">Accr.</th><th class="num">Clear</th><th class="num">Close</th>', FSL.map(function (s) { return '<td>' + esc(s.name) + '</td><td class="num">' + s.tot.ob + '</td><td class="num">' + s.tot.acc + '</td><td class="num">' + s.tot.clr + '</td><td class="num fw-bold">' + s.tot.cb + '</td>'; }),
      function (i) { open([{ l: 'Sections', f: vFileSecs }, { l: FSL[i].name, f: function () { return vOfficers(FSL[i]); } }]); });

    var AM = A.filter(function (a) { return !isTot(a.n) && a.cb > 0; });
    draw('cAud', { type: 'doughnut', data: { labels: AM.map(function (a) { return a.n; }), datasets: [{ data: AM.map(function (a) { return a.cb; }), backgroundColor: PIE }] }, options: { cutout: '58%' } }, function (i) { open([{ l: AM[i].n, f: function () { return vAudit(AM[i]); } }]); });
    tbl('#tAud', '<th>Nature</th><th class="num">OB</th><th class="num">Accrued</th><th class="num">Closed</th><th class="num">Closing</th>', A.map(function (a) { return '<td class="' + (isTot(a.n) ? 'fw-bold' : '') + '">' + esc(a.n) + '</td><td class="num">' + a.ob + '</td><td class="num">' + a.acc + '</td><td class="num">' + a.clo + '</td><td class="num fw-bold">' + a.cb + '</td>'; }),
      function (i) { open([{ l: A[i].n, f: function () { return vAudit(A[i]); } }]); });

    var RS = {}; D.rb.forEach(function (r) { RS[r.sec] = (RS[r.sec] || 0) + 1; });
    var RK = Object.keys(RS);
    draw('cRb', { type: 'bar', data: { labels: RK, datasets: [{ data: RK.map(function (k) { return RS[k]; }), backgroundColor: PIE[3] }] }, options: { indexAxis: 'y', plugins: { legend: { display: false } } } }, function (i) {
      var L = D.rb.filter(function (r) { return r.sec === RK[i]; });
      open([{ l: 'RB Inspection - ' + RK[i], f: function () { return { title: 'Paras pending with ' + RK[i], kpis: [['Paras', L.length, 'r']], head: '<th>Para</th><th>Since</th><th>Remarks</th>', rows: L.map(function (r) { return '<td>' + esc(r.para) + '</td><td>' + esc(r.since) + '</td><td class="small">' + esc(clip(r.rem, 120)) + '</td>'; }), onRow: function (j) { push('Para ' + L[j].para, function () { return vRb(L[j]); }); } }; } }]);
    });
    tbl('#tRb', '<th>Para</th><th>Section</th><th>Pending since</th><th>Audit remarks</th>', D.rb.map(function (r) { return '<td class="fw-bold">' + esc(r.para) + '</td><td><span class="pill">' + esc(r.sec) + '</span></td><td>' + esc(r.since) + '</td><td class="small">' + esc(clip(r.rem, 110)) + '</td>'; }),
      function (i) { open([{ l: 'Para ' + D.rb[i].para, f: function () { return vRb(D.rb[i]); } }]); });
  }
});
