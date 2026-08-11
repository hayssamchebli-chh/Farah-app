/* Price History Pivot — turns a flat sales-lines export into one row per item
   with repeated [R, Date, Qty, Price, Curr.] blocks ordered newest first. */

(function () {
  'use strict';

  var BLOCK = ['R', 'Date', 'Qty', 'Price', 'Curr.'];

  // Header aliases, matched after normalising (lowercase, alphanumerics only).
  var FIELDS = {
    no:    ['no', 'nos', 'itemno', 'item', 'itemnumber', 'partno', 'partnumber'],
    desc:  ['description', 'itemdescription', 'desc'],
    qty:   ['quantity', 'qty'],
    price: ['ocnetprice', 'netpriceoc', 'ocnet', 'priceoc'],
    curr:  ['currency', 'curr', 'currencycode'],
    date:  ['postingdate', 'postingdt', 'date', 'documentdate']
  };

  var workbook = null;
  var pivot = null;

  var el = {
    drop: document.getElementById('drop'),
    file: document.getElementById('file'),
    browse: document.getElementById('browse'),
    options: document.getElementById('options'),
    includeDesc: document.getElementById('includeDesc'),
    sheet: document.getElementById('sheet'),
    status: document.getElementById('status'),
    result: document.getElementById('result'),
    summary: document.getElementById('summary'),
    preview: document.getElementById('preview'),
    previewNote: document.getElementById('previewNote'),
    download: document.getElementById('download'),
    reset: document.getElementById('reset')
  };

  /* ---------- helpers ---------- */

  function norm(s) {
    return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function say(msg, kind) {
    el.status.textContent = msg;
    el.status.className = 'status ' + (kind || '');
  }

  function show(node, on) { node.classList.toggle('hidden', !on); }

  // Map header cells to our field names. Longest alias wins so that
  // "OC Net Price" is not swallowed by a generic "price".
  function mapHeaders(headerRow) {
    var map = {};
    headerRow.forEach(function (cell, i) {
      var n = norm(cell);
      if (!n) return;
      Object.keys(FIELDS).forEach(function (field) {
        if (map[field] !== undefined) return;
        if (FIELDS[field].indexOf(n) !== -1) map[field] = i;
      });
    });
    // Fall back to a plain "Net Price" column if "OC Net Price" is absent.
    if (map.price === undefined) {
      headerRow.forEach(function (cell, i) {
        if (map.price === undefined && norm(cell) === 'netprice') map.price = i;
      });
    }
    return map;
  }

  // Find the header row: the first row (within the first 20) that yields
  // both a No. column and a Posting Date column.
  function findHeader(rows) {
    var limit = Math.min(rows.length, 20);
    for (var r = 0; r < limit; r++) {
      var map = mapHeaders(rows[r] || []);
      if (map.no !== undefined && map.date !== undefined) return { row: r, map: map };
    }
    return null;
  }

  function toDate(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number' && isFinite(v)) {
      // Excel serial date (1900 system, including its phantom 1900-02-29).
      if (v < 1 || v > 2958465) return null;
      var days = Math.floor(v) - (v < 61 ? 0 : 1);
      var ms = Date.UTC(1899, 11, 31) + days * 86400000;
      var u = new Date(ms);
      return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
    }
    var s = String(v == null ? '' : v).trim();
    if (!s) return null;
    // dd/mm/yyyy and dd-mm-yyyy are the common exports; treat them as day-first.
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    if (m) {
      var yy = +m[3];
      if (yy < 100) yy += yy < 70 ? 2000 : 1900;
      return new Date(yy, +m[2] - 1, +m[1]);
    }
    var p = new Date(s);
    return isNaN(p.getTime()) ? null : p;
  }

  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v == null ? '' : v).trim().replace(/[\s,']/g, '');
    if (!s) return null;
    var n = Number(s);
    return isFinite(n) ? n : null;
  }

  function fmtDate(d) {
    if (!(d instanceof Date)) return '';
    var p = function (x) { return (x < 10 ? '0' : '') + x; };
    return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear();
  }

  /* ---------- the pivot ---------- */

  function buildPivot(rows, includeDesc) {
    var found = findHeader(rows);
    if (!found) {
      throw new Error('Could not find a header row containing both "No." and "Posting Date".');
    }
    var map = found.map;
    var missing = ['no', 'qty', 'price', 'curr', 'date'].filter(function (f) {
      return map[f] === undefined;
    });
    if (missing.length) {
      var labels = { no: 'No.', qty: 'Quantity', price: 'OC Net Price', curr: 'Currency', date: 'Posting Date' };
      throw new Error('Missing column(s): ' + missing.map(function (f) { return labels[f]; }).join(', '));
    }

    var groups = {};      // key -> { no, desc, entries[] }
    var order = [];       // first-seen order of keys
    var skipped = 0;

    for (var r = found.row + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var noRaw = row[map.no];
      var no = String(noRaw == null ? '' : noRaw).trim();
      if (!no) { if (row.some(function (c) { return c !== undefined && c !== ''; })) skipped++; continue; }

      var key = no.toUpperCase();
      if (!groups[key]) { groups[key] = { no: no, desc: '', entries: [] }; order.push(key); }
      var g = groups[key];
      if (includeDesc && !g.desc && map.desc !== undefined) {
        g.desc = String(row[map.desc] == null ? '' : row[map.desc]).trim();
      }
      g.entries.push({
        date: toDate(row[map.date]),
        qty: toNum(row[map.qty]),
        price: toNum(row[map.price]),
        curr: String(row[map.curr] == null ? '' : row[map.curr]).trim(),
        seq: r
      });
    }

    if (!order.length) throw new Error('No data rows found under the header.');

    // Newest first; rows without a date sink to the bottom. Same-date rows keep
    // their original file order.
    var maxDup = 0;
    order.forEach(function (key) {
      var e = groups[key].entries;
      e.sort(function (a, b) {
        var ta = a.date ? a.date.getTime() : -Infinity;
        var tb = b.date ? b.date.getTime() : -Infinity;
        if (ta !== tb) return tb - ta;
        return a.seq - b.seq;
      });
      if (e.length > maxDup) maxDup = e.length;
    });

    var header = ['No.'];
    if (includeDesc) header.push('Description');
    for (var i = 1; i <= maxDup; i++) BLOCK.forEach(function (h) { header.push(h); });

    var body = order.map(function (key) {
      var g = groups[key];
      var out = [g.no];
      if (includeDesc) out.push(g.desc);
      for (var i = 0; i < maxDup; i++) {
        var e = g.entries[i];
        if (!e) { out.push('', '', '', '', ''); continue; }
        out.push(i + 1, e.date || '', e.qty == null ? '' : e.qty, e.price == null ? '' : e.price, e.curr);
      }
      return out;
    });

    return {
      header: header,
      body: body,
      maxDup: maxDup,
      items: order.length,
      lines: order.reduce(function (s, k) { return s + groups[k].entries.length; }, 0),
      skipped: skipped,
      includeDesc: includeDesc
    };
  }

  /* ---------- rendering ---------- */

  function render(p) {
    var maxRows = 100;
    var html = '<thead><tr>' + p.header.map(function (h, i) {
      return '<th' + (i === 0 ? ' class="sticky"' : '') + '>' + h + '</th>';
    }).join('') + '</tr></thead><tbody>';

    var firstBlock = p.includeDesc ? 2 : 1;
    var isRCol = function (i) {
      return i >= firstBlock && (i - firstBlock) % 5 === 0;
    };

    p.body.slice(0, maxRows).forEach(function (row) {
      html += '<tr>' + row.map(function (c, i) {
        var v = c instanceof Date ? fmtDate(c) : (c === '' || c == null ? '' : String(c));
        var cls = i === 0 ? ' class="sticky"'
          : (isRCol(i) && v !== '' ? ' class="rcell"'
            : (typeof c === 'number' ? ' class="num"' : ''));
        return '<td' + cls + '>' + v.replace(/[&<>]/g, function (ch) {
          return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[ch];
        }) + '</td>';
      }).join('') + '</tr>';
    });
    el.preview.innerHTML = html + '</tbody>';

    el.summary.textContent = p.items + ' unique item' + (p.items === 1 ? '' : 's') +
      ' · ' + p.lines + ' lines · up to ' + p.maxDup + ' transaction' + (p.maxDup === 1 ? '' : 's') + ' per item';
    el.previewNote.textContent =
      (p.body.length > maxRows ? 'Showing the first ' + maxRows + ' of ' + p.body.length + ' rows. ' : '') +
      (p.skipped ? p.skipped + ' row(s) were skipped for having no No. value.' : '');
    show(el.result, true);
  }

  function exportWorkbook(p) {
    var aoa = [p.header].concat(p.body);
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

    // Date format + column widths.
    var firstBlock = p.includeDesc ? 2 : 1;
    var widths = [{ wch: 18 }];
    if (p.includeDesc) widths.push({ wch: 40 });
    for (var i = 0; i < p.maxDup; i++) {
      widths.push({ wch: 4 }, { wch: 12 }, { wch: 9 }, { wch: 12 }, { wch: 7 });
    }
    ws['!cols'] = widths;
    ws['!freeze'] = { xSplit: firstBlock, ySplit: 1 };

    for (var r = 1; r <= p.body.length; r++) {
      for (var b = 0; b < p.maxDup; b++) {
        var addr = XLSX.utils.encode_cell({ r: r, c: firstBlock + b * 5 + 1 });
        var cell = ws[addr];
        if (cell && cell.t === 'd') cell.z = 'dd/mm/yyyy';
      }
    }

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Price History');
    var stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, 'price-history-pivot-' + stamp + '.xlsx', { cellDates: true });
  }

  /* ---------- wiring ---------- */

  function process() {
    if (!workbook) return;
    try {
      say('Building pivot…');
      var ws = workbook.Sheets[el.sheet.value];
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false, defval: '' });
      pivot = buildPivot(rows, el.includeDesc.checked);
      render(pivot);
      say('Done.', 'ok');
    } catch (err) {
      pivot = null;
      show(el.result, false);
      say(err.message || String(err), 'error');
    }
  }

  function load(file) {
    say('Reading ' + file.name + '…');
    var reader = new FileReader();
    reader.onerror = function () { say('Could not read that file.', 'error'); };
    reader.onload = function (e) {
      try {
        workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      } catch (err) {
        say('That does not look like a readable Excel/CSV file.', 'error');
        return;
      }
      el.sheet.innerHTML = workbook.SheetNames.map(function (n) {
        return '<option>' + n.replace(/[&<>]/g, '') + '</option>';
      }).join('');
      show(el.options, true);
      process();
    };
    reader.readAsArrayBuffer(file);
  }

  el.browse.addEventListener('click', function () { el.file.click(); });
  el.drop.addEventListener('click', function (ev) {
    if (ev.target !== el.browse) el.file.click();
  });
  el.file.addEventListener('change', function () {
    if (el.file.files && el.file.files[0]) load(el.file.files[0]);
  });
  ['dragenter', 'dragover'].forEach(function (t) {
    el.drop.addEventListener(t, function (ev) { ev.preventDefault(); el.drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (t) {
    el.drop.addEventListener(t, function (ev) { ev.preventDefault(); el.drop.classList.remove('over'); });
  });
  el.drop.addEventListener('drop', function (ev) {
    var f = ev.dataTransfer && ev.dataTransfer.files && ev.dataTransfer.files[0];
    if (f) load(f);
  });
  el.sheet.addEventListener('change', process);
  el.includeDesc.addEventListener('change', process);
  el.download.addEventListener('click', function () { if (pivot) exportWorkbook(pivot); });
  el.reset.addEventListener('click', function () {
    workbook = null; pivot = null; el.file.value = '';
    show(el.options, false); show(el.result, false); show(el.status, false);
    el.status.className = 'status hidden';
  });

  // Exposed for the Node-based test harness.
  if (typeof module !== 'undefined' && module.exports) module.exports = { buildPivot: buildPivot };
  window.__pivot = { buildPivot: buildPivot };
})();
