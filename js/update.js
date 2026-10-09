$(function () {
  var P = DashParser(XLSX), D = { asOn: '', exp: [], files: [], pending: [], audit: [], rb: [] }, wbs = {}, FL = [];
  var LBL = { template: 'Input template (all sheets)', exp: 'Expenditure (PH-wise)', works: 'Work-wise', files: 'File pendency', audit: 'Audit', rb: 'Railway Board Inspection' };
  var esc = function (s) { return $('<i>').text(s == null ? '' : s).html(); };
  var iso = function (dmy) { var m = /^(\d\d)\.(\d\d)\.(\d{4})$/.exec(dmy || ''); return m ? m[3] + '-' + m[2] + '-' + m[1] : ''; };

  $.getJSON('data/dashboard.json').done(function (j) {
    D = j; $('#asOnIn').val(iso(j.asOn));
    $('#baseMsg').text('Current data/dashboard.json loaded (position as on ' + j.asOn + '). Uploaded files will update it.'); manual();
  }).fail(function () {
    $('#baseMsg').removeClass('alert-info').addClass('alert-warning').text('No existing data/dashboard.json could be read (open this page through run.bat / a web server to merge with it). A fresh file will be created from what you provide.');
    $('#asOnIn').val(new Date().toISOString().slice(0, 10)); manual();
  });

  function addFiles(list) {
    $.each(list, function (_, f) {
      var rd = new FileReader();
      rd.onload = function (e) {
        var row = { name: f.name };
        try {
          var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          row.t = P.detect(wb, f.name); row.wb = wb;
          if (row.t) wbs[row.t] = wb;
        } catch (x) { row.err = 'Unreadable file'; }
        FL.push(row); listFiles();
      };
      rd.readAsArrayBuffer(f);
    });
  }
  function listFiles() {
    $('#fList').html(FL.map(function (r, i) {
      var st = r.err ? '<span class="badge text-bg-danger">' + r.err + '</span>' : r.t ? '<span class="badge text-bg-success">' + LBL[r.t] + '</span>' :
        '<select class="form-select form-select-sm" data-i="' + i + '"><option value="">Not recognised - choose...</option><option value="audit">Audit</option><option value="rb">Railway Board Inspection</option></select>';
      return '<tr><td>' + esc(r.name) + '</td><td class="text-end">' + st + '</td></tr>';
    }).join(''));
    $('#fList select').on('change', function () { var r = FL[$(this).data('i')]; r.t = this.value; if (r.t) wbs[r.t] = r.wb; listFiles(); });
  }
  $('#pick').on('change', function () { addFiles(this.files); this.value = ''; });
  $('#drop').on('dragover dragenter', function (e) { e.preventDefault(); $(this).addClass('over'); })
    .on('dragleave drop', function (e) { e.preventDefault(); $(this).removeClass('over'); if (e.type === 'drop') addFiles(e.originalEvent.dataTransfer.files); });

  /* ----- forms ----- */
  function val(box, f) { return $(box).find('[data-f="' + f + '"]'); }
  function manual() {
    var rows = D.audit.map(function (a, i) { return '<tr><td><span class="pill">Audit</span></td><td>' + esc(a.n) + ' &middot; closing ' + a.cb + '</td><td class="text-end"><a href="#" data-k="audit" data-i="' + i + '">remove</a></td></tr>'; })
      .concat(D.rb.map(function (r, i) { return '<tr><td><span class="pill">RB</span></td><td>Para ' + esc(r.para) + ' &middot; ' + esc(r.sec) + '</td><td class="text-end"><a href="#" data-k="rb" data-i="' + i + '">remove</a></td></tr>'; }));
    $('#mList').html(rows.join('')).find('a').on('click', function (e) { e.preventDefault(); D[$(this).data('k')].splice($(this).data('i'), 1); manual(); });
  }
  $('#addAud').on('click', function () {
    var b = '#fAud', n = val(b, 'n').val().trim(); if (!n) return val(b, 'n').focus();
    var g = function (f) { return parseFloat(val(b, f).val()) || 0; }, o = { n: n, ob: g('ob'), acc: g('acc'), tot: g('ob') + g('acc'), clo: g('clo'), cb: g('ob') + g('acc') - g('clo') };
    var at = D.audit.length && /Total/.test(D.audit[D.audit.length - 1].n) ? D.audit.length - 1 : D.audit.length; D.audit.splice(at, 0, o);
    $(b).find('input').val(''); manual();
  });
  $('#addRb').on('click', function () {
    var b = '#fRb', p = val(b, 'para').val().trim(); if (!p) return val(b, 'para').focus();
    D.rb.push({ para: p, sec: val(b, 'sec').val().trim(), since: val(b, 'since').val(), rem: val(b, 'rem').val(), sect: val(b, 'sect').val() });
    $(b).find('input,textarea').val(''); manual();
  });

  /* ----- generate ----- */
  $('#gen').on('click', function () {
    try {
      var tn = [];
      if (wbs.template) tn = P.parseTemplate(wbs.template, D);
      if (wbs.exp) {
        var old = {}; D.exp.forEach(function (s) { s.heads.forEach(function (h) { old[s.code + '|' + h.ph] = h.works; }); });
        D.exp = P.parseExp(wbs.exp);
        D.labels = P.parseLabels(wbs.exp) || D.labels;
        D.exp.forEach(function (s) { s.heads.forEach(function (h) { h.works = old[s.code + '|' + h.ph] || []; }); });
      }
      if (wbs.works) P.attachWorks(wbs.works, D.exp);
      if (wbs.files) { var r = P.parseFiles(wbs.files); if (r.files) D.files = r.files; if (r.pending) D.pending = r.pending; }
      if (wbs.audit) D.audit = P.parseAudit(wbs.audit);
      if (wbs.rb) D.rb = P.parseRb(wbs.rb);
      $('#asOnIn').val(iso(D.asOn)); wbs = {}; FL.forEach(function (r) { r.done = 1; });
      var d = $('#asOnIn').val(); if (d && tn.indexOf('Position date') < 0) D.asOn = d.slice(8) + '.' + d.slice(5, 7) + '.' + d.slice(0, 4);
      var nw = 0; D.exp.forEach(function (s) { s.heads.forEach(function (h) { nw += h.works.length; }); });
      $('#sum').html([['Position as on', D.asOn], ['Expenditure sections', D.exp.length], ['Works', nw], ['File sections', D.files.length], ['Pending files listed', D.pending.length], ['Audit lines', D.audit.length], ['RB inspection paras', D.rb.length]]
        .map(function (x) { return '<tr><td>' + x[0] + '</td><td class="text-end fw-bold">' + esc(x[1]) + '</td></tr>'; }).join(''));
      var blob = new Blob([JSON.stringify(D)], { type: 'application/json' });
      $('#dl').attr('href', URL.createObjectURL(blob)).removeClass('disabled'); manual();
    } catch (x) { alert('Could not process the files: ' + x.message); }
  });
});
