/* ===== [AP v27] top "Download report" follows the Ageing view · 04-Oct-2026 =====
   Paste at the very END of the Javascript field, after the [AP v26] block.
   Nothing above it changes.

   The top "Download report" button (data-a3dl="all") used to run v24's fullBook,
   whose section 1 was always supplier rows in four buckets. It ignored the
   Supplier level / Bill wise pills and the six buckets v25 puts on screen.

   This block catches that one click before v24 does, and builds the workbook again:
     Supplier level -> section 1 = the six-bucket supplier table on screen (v25 rules:
                       debit notes kept out of the buckets, shown on their own line)
     Bill wise      -> section 1 = every open bill, with the same columns and order as
                       v25's "Download bill-wise" CSV, plus a recap by bucket
   Ages come from window.ap24Model(), which v24 recomputes whenever Bill date /
   Posting date / Due date is picked, so the export always uses the selected basis.
   The summary, tables 2 and 3, the Check and the supplier_recon sheet are copied
   unchanged from v24's fullBook.
   Read-only: reads window.ap24Model() and the pills v24 paints. Writes nothing to ERP. */
try { (function () {
  var TAG = '[AP v27]';
  var ACC = 'Accounts Payable - FT';
  var LBL6 = ['Not due', '1-30', '31-60', '61-90', '91-180', 'Over 180'];

  function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
  function inr(n) { var v = Number(n) || 0, p = Math.abs(v).toFixed(2).split('.');
    var i = p[0], l3 = i.slice(-3), r = i.slice(0, -3); if (r) l3 = ',' + l3;
    return (v < 0 ? '-' : '') + r.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + l3 + '.' + p[1]; }
  /* the same bucket rule as v25 */
  function bidx(age) { var g = Number(age);
    return g <= 0 ? 0 : g <= 30 ? 1 : g <= 60 ? 2 : g <= 90 ? 3 : g <= 180 ? 4 : 5; }

  /* the current view, read from the pills v24 paints (.on = selected) */
  function view(rn) {
    var lv = rn && rn.querySelector ? rn.querySelector('.a3pill.on[data-a3lvl]') : null;
    var bs = rn && rn.querySelector ? rn.querySelector('.a3pill.on[data-a3basis]') : null;
    return { bill: !!lv && lv.getAttribute('data-a3lvl') === 'inv',
             basis: bs ? bs.getAttribute('data-a3basis') : 'due' };   /* v24 starts with BASIS = 'due' */
  }

  /* the same model as v25's six(): six buckets, debit notes kept apart */
  function six(m) {
    var tot = [0, 0, 0, 0, 0, 0], cnt = [0, 0, 0, 0, 0, 0], dn = 0, dnN = 0, sup = [], bills = [];
    (m.A || []).forEach(function (v) {
      var b = [0, 0, 0, 0, 0, 0], t = 0;
      (v.rows || []).forEach(function (r) {
        var a = Number(r.amt) || 0;
        bills.push({ sp: v.p, sn: v.n, no: r.no, vno: r.vno, bd: r.bd, due: r.due,
                     age: r.age, amt: a, bk: (a < 0 ? 'Debit note' : LBL6[bidx(r.age)]) });
        if (a < 0) { dn += a; dnN++; return; }
        var i = bidx(r.age); b[i] += a; tot[i] += a; cnt[i]++; t += a;
      });
      if (Math.abs(t) > 0.005 || Math.abs(v.adv || 0) > 0.005)
        sup.push({ p: v.p, n: v.n, b: b, tot: t, adv: v.adv || 0 });
    });
    sup.sort(function (x, y) { return y.tot - x.tot; });
    var sum = tot.reduce(function (a, x) { return a + x; }, 0);
    return { tot: tot, cnt: cnt, sup: sup, dn: dn, dnN: dnN, sum: sum, bills: bills };
  }

  function book(m, vw) {
    var T = function (v) { return [v, 't']; }, B = function (v) { return [v, 'b']; }, N = function (v) { return [v, 'n']; },
        NB = function (v) { return [v, 'N']; }, H = function (v) { return [v, 'h']; }, TT = function (v) { return [v, 'T']; },
        I = function (v) { return [v, 'i']; };
    var basis = vw.basis === 'post' ? 'Posting Date' : vw.basis === 'due' ? 'Due Date' : 'Bill Date';
    var s = six(m), R = [], tie = m.supTie, h1 = 0, e1 = 0, j;
    R.push([TT('FURNISHKA TECH PRIVATE LIMITED')]);
    R.push([T('AP Aging Summary By ' + basis + ' as of ' + m.d), null, null, null, null, T(ACC)]);
    R.push([T('Scope: ' + (m.scopeTxt || 'company-wide'))]);
    R.push([T('View: ' + (vw.bill ? 'Bill wise' : 'Supplier level') + ' · aged on ' + basis.toLowerCase())]);
    R.push([]);
    R.push([B('Bills outstanding, aged (table 1)'), null, null, null, null, NB(m.billTot), T(m.bills + ' bills')]);
    R.push([B('Less: advance paid, bill booked but not allocated (table 2)'), null, null, null, null, NB(-m.advHeld), T(m.advHeldN + ' suppliers')]);
    R.push([B('Less: advance paid, bill not booked (table 3)'), null, null, null, null, NB(-m.advOnly), T(m.B.length + ' suppliers')]);
    R.push([B('Add: payments carrying a payable balance'), null, null, null, null, NB(m.ppTot), T(m.ppN + ' suppliers')]);
    R.push([B('Add: open journal items'), null, null, null, null, NB(m.jv), T(m.jvN + ' journal rows')]);
    R.push([B('Net payable per books'), null, null, null, null, NB(m.net), T('general ledger ' + inr(m.ctl) + ', difference ' + inr(m.delta))]);
    R.push([B('Supplier-wise tie to the ledger'), null, null, null, null, null, T(tie ? tie.ok + ' of ' + tie.n + ' suppliers agree (see supplier_recon)' : 'not computed')]);
    R.push([]);

    if (vw.bill) {
      /* bill wise: v25's "Download bill-wise" columns and order */
      R.push([B('1. Ageing of bills - bill wise')]);
      h1 = R.length + 1;
      R.push([H('vendor_name'), H('supplier_id'), H('bill_no'), H('voucher_no'),
              H('aged_from (' + basis.toLowerCase() + ')'), H('due_date'), H('age_days'), H('bucket'), H('balance'),
              H('comment'), H('status')]);
      var bl = s.bills.slice().sort(function (a, b) {
        return String(a.sn).localeCompare(String(b.sn)) || b.amt - a.amt; });
      bl.forEach(function (r) {
        R.push([T(r.sn), T(r.sp), T(r.no || r.vno), T(r.vno), T(r.bd), T(r.due), I(r.age), T(r.bk), N(r.amt)]);
      });
      e1 = R.length;
      R.push([B('TOTAL (' + bl.length + ' bills, ' + s.sup.length + ' suppliers)'), null, null, null, null, null, null, null, NB(m.billTot)]);
      R.push([]);
      R.push([H('bucket'), H('amount'), H('bills')]);
      for (j = 0; j < 6; j++) R.push([T(LBL6[j]), N(s.tot[j]), I(s.cnt[j])]);
      R.push([T('Debit notes'), N(s.dn), I(s.dnN)]);
      R.push([B('Bills outstanding'), NB(r2(s.sum + s.dn)), I(bl.length)]);
      R.push([T('Difference to table 1 total'), N(r2(s.sum + s.dn - m.billTot))]);
    } else {
      /* supplier level: v25's six-bucket table, same rows, same totals */
      R.push([B('1. Ageing of bills - supplier level')]);
      h1 = R.length + 1;
      R.push([H('vendor_name'), H('not_due'), H('days_1-30'), H('days_31-60'), H('days_61-90'),
              H('days_91-180'), H('days_above-180'), H('total'), H('advance_held'), H('comment'), H('status')]);
      s.sup.forEach(function (v) {
        R.push([T(v.n), N(v.b[0]), N(v.b[1]), N(v.b[2]), N(v.b[3]), N(v.b[4]), N(v.b[5]), N(v.tot),
                v.adv > 0.005 ? N(v.adv) : null]);
      });
      e1 = R.length;
      R.push([B('TOTAL (' + s.sup.length + ' suppliers)'), NB(s.tot[0]), NB(s.tot[1]), NB(s.tot[2]), NB(s.tot[3]),
              NB(s.tot[4]), NB(s.tot[5]), NB(s.sum), NB(m.advHeld)]);
      if (s.dnN) R.push([T('Debit notes (' + s.dnN + ') - their own class, never inside a bucket'),
                         null, null, null, null, null, null, NB(s.dn)]);
      R.push([B('Bills outstanding = buckets + debit notes'), null, null, null, null, null, null, NB(r2(s.sum + s.dn))]);
      R.push([T('Difference to table 1 total'), null, null, null, null, null, null, N(r2(s.sum + s.dn - m.billTot))]);
    }

    /* tables 2 and 3, Check and supplier_recon: unchanged from v24 fullBook */
    R.push([]); R.push([]);
    var L = m.A.filter(function (v) { return v.adv > 0.005; }).sort(function (a, b) { return b.adv - a.adv; });
    var alc = r2(L.reduce(function (a, v) { return a + Math.min(v.adv, Math.max(v.tot, 0)); }, 0));
    R.push([B('2. Advances paid, bill booked but not allocated')]);
    R.push([H('vendor_name'), H('advance_paid'), H('open_bills'), H('allocatable'), null, null, null, H('comment'), H('status')]);
    L.forEach(function (v) { R.push([T(v.n), N(v.adv), N(v.tot), N(Math.min(v.adv, Math.max(v.tot, 0)))]); });
    R.push([B('TOTAL'), NB(m.advHeld), NB(r2(L.reduce(function (a, v) { return a + v.tot; }, 0))), NB(alc)]);
    R.push([]); R.push([]);
    R.push([B('3. Advance payments - bill not booked')]);
    R.push([H('vendor_name'), H('advance_paid'), null, null, null, null, null, H('comment'), H('status')]);
    m.B.forEach(function (v) { R.push([T(v.n), N(v.adv)]); });
    R.push([B('TOTAL'), NB(m.advOnly)]);
    R.push([]); R.push([]);
    R.push([B('Check')]);
    R.push([T('Table 1 total'), null, null, null, null, N(m.billTot)]);
    R.push([T('Less: table 2 total'), null, null, null, null, N(-m.advHeld)]);
    R.push([T('Less: table 3 total'), null, null, null, null, N(-m.advOnly)]);
    R.push([T('Add: payments carrying a payable balance'), null, null, null, null, N(m.ppTot)]);
    R.push([T('Add: open journal items'), null, null, null, null, N(m.jv)]);
    R.push([B('Net payable per books'), null, null, null, null, NB(m.net)]);
    R.push([T('General ledger, ' + ACC), null, null, null, null, N(m.ctl)]);
    R.push([B('Difference'), null, null, null, null, NB(m.delta)]);
    var S2 = [[H('vendor_name'), H('supplier'), H('bills'), H('advance_paid'), H('payments_payable_side'), H('journals'), H('net_per_report'), H('general_ledger'), H('difference')]];
    var tb = 0, ta = 0, tp = 0, tj = 0, tn = 0, tg = 0, td = 0;
    (tie ? tie.L : []).forEach(function (x) {
      S2.push([T(x.n), T(x.p), N(x.bills), N(-x.adv), N(x.pp), N(x.je), N(x.net), N(x.gl), N(x.df)]);
      tb += x.bills; ta += x.adv; tp += x.pp; tj += x.je; tn += x.net; tg += x.gl; td += x.df;
    });
    S2.push([B('TOTAL (' + (tie ? tie.n : 0) + ' suppliers, ' + (tie ? tie.ok : 0) + ' agree)'), null, NB(tb), NB(-ta), NB(tp), NB(tj), NB(tn), NB(tg), NB(r2(td))]);
    return [
      { name: 'ap_aging_summary', rows: R, sel: 1, af: 'A' + h1 + ':K' + e1,
        widths: vw.bill ? [46, 22, 24, 22, 16, 13, 10, 12, 18, 40, 13]
                        : [58, 15, 15, 15, 15, 15, 15, 18, 18, 40, 13] },
      { name: 'supplier_recon', rows: S2, widths: [48, 18, 18, 18, 22, 16, 18, 18, 14], freeze: 1, af: 'A1:I' + (S2.length - 1) }
    ];
  }

  /* workbook writer: v24 xlsxBook, plus the integer style 'i' from v24's single-sheet xlsx() */
  function dl(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
  }
  function xlsxBook(sheets, name) {
    var enc = new TextEncoder();
    var CT = (function () { var a = new Uint32Array(256); for (var i = 0; i < 256; i++) { var c = i; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; a[i] = c >>> 0; } return a; })();
    function crc(u) { var c = 0xFFFFFFFF; for (var i = 0; i < u.length; i++) c = CT[(c ^ u[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
    function xx(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
    function cl(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
    function sheetXml(sh) {
      var sx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
      sx += '<sheetViews><sheetView workbookViewId="0"' + (sh.sel ? ' tabSelected="1"' : '') + '>' + (sh.freeze ? '<pane ySplit="' + sh.freeze + '" topLeftCell="A' + (sh.freeze + 1) + '" activePane="bottomLeft" state="frozen"/>' : '') + '</sheetView></sheetViews>';
      if (sh.widths) sx += '<cols>' + sh.widths.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols>';
      sx += '<sheetData>';
      sh.rows.forEach(function (row, ri) {
        var r = ri + 1, cs = '';
        (row || []).forEach(function (cv, ci) {
          if (!cv) return;
          var ref = cl(ci) + r, v = cv[0], k = cv[1];
          if (v === null || v === undefined || v === '') return;
          if (k === 'n') cs += '<c r="' + ref + '" s="2"><v>' + r2(v) + '</v></c>';
          else if (k === 'N') cs += '<c r="' + ref + '" s="3"><v>' + r2(v) + '</v></c>';
          else if (k === 'i') cs += '<c r="' + ref + '" s="6"><v>' + (Number(v) || 0) + '</v></c>';
          else if (k === 'h') cs += '<c r="' + ref + '" s="4" t="inlineStr"><is><t>' + xx(v) + '</t></is></c>';
          else if (k === 'b') cs += '<c r="' + ref + '" s="1" t="inlineStr"><is><t>' + xx(v) + '</t></is></c>';
          else if (k === 'T') cs += '<c r="' + ref + '" s="5" t="inlineStr"><is><t>' + xx(v) + '</t></is></c>';
          else cs += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xx(v) + '</t></is></c>';
        });
        if (cs) sx += '<row r="' + r + '">' + cs + '</row>';
      });
      sx += '</sheetData>';
      if (sh.af) sx += '<autoFilter ref="' + sh.af + '"/>';
      return sx + '</worksheet>';
    }
    var st = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<numFmts count="1"><numFmt numFmtId="164" formatCode="[&gt;=100000]##\\,##\\,##0.00;[&lt;=-100000]\\-##\\,##\\,##0.00;##,##0.00"/></numFmts>'
      + '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font></fonts>'
      + '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEEF0F3"/><bgColor indexed="64"/></patternFill></fill></fills>'
      + '<borders count="2"><border/><border><bottom style="thin"><color rgb="FF9AA2AE"/></bottom></border></borders>'
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + '<cellXfs count="7">'
      + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
      + '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
      + '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
      + '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
      + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
      + '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
      + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
    var rel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    var files = [];
    sheets.forEach(function (sh, i) {
      ct += '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      wb += '<sheet name="' + xx(sh.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      rel += '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      files.push(['xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(sh)]);
    });
    ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    wb += '</sheets></workbook>';
    rel += '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    files = [['[Content_Types].xml', ct], ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'], ['xl/workbook.xml', wb], ['xl/_rels/workbook.xml.rels', rel], ['xl/styles.xml', st]].concat(files);
    var parts = [], cen = [], off = 0;
    function u16(n) { return [n & 255, (n >> 8) & 255]; }
    function u32(n) { return [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]; }
    var nw = new Date();
    var dt = ((nw.getHours() << 11) | (nw.getMinutes() << 5) | (nw.getSeconds() >> 1)) & 0xFFFF;
    var dd = (((nw.getFullYear() - 1980) << 9) | ((nw.getMonth() + 1) << 5) | nw.getDate()) & 0xFFFF;
    files.forEach(function (f) {
      var nb = enc.encode(f[0]), db = enc.encode(f[1]), c = crc(db);
      var lh = [].concat([80, 75, 3, 4], u16(20), u16(0), u16(0), u16(dt), u16(dd), u32(c), u32(db.length), u32(db.length), u16(nb.length), u16(0));
      parts.push(new Uint8Array(lh), nb, db);
      cen.push({ nb: nb, c: c, s: db.length, o: off });
      off += lh.length + nb.length + db.length;
    });
    var cd = [];
    cen.forEach(function (e) {
      var ch = [].concat([80, 75, 1, 2], u16(20), u16(20), u16(0), u16(0), u16(dt), u16(dd), u32(e.c), u32(e.s), u32(e.s), u16(e.nb.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(e.o));
      cd.push(new Uint8Array(ch), e.nb);
    });
    var cs2 = cd.reduce(function (a, x) { return a + x.length; }, 0);
    var eo = new Uint8Array([].concat([80, 75, 5, 6], u16(0), u16(0), u16(cen.length), u16(cen.length), u32(cs2), u32(off), u16(0)));
    dl(new Blob(parts.concat(cd, [eo]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name + '.xlsx');
  }

  /* window capture runs before v24's capture listener on ROOT, so this click is ours */
  if (!window.__ap27Bound) {
    window.__ap27Bound = 1;
    window.addEventListener('click', function (ev) {
      var path = (ev.composedPath && ev.composedPath()) || [], hit = null;
      for (var i = 0; i < path.length; i++) {
        var n = path[i];
        if (n && n.getAttribute && n.getAttribute('data-a3dl') === 'all') { hit = n; break; }
      }
      if (!hit) return;
      var m = typeof window.ap24Model === 'function' ? window.ap24Model() : null;
      if (!m || !m.A) return;                   /* no model yet: v24 handles it, as before */
      ev.preventDefault(); ev.stopImmediatePropagation();
      try {
        var vw = view(hit.getRootNode ? hit.getRootNode() : document);
        var name = 'AP_Aging_Summary_' + (vw.bill ? 'Bill_Wise' : 'Supplier_Level') + '_By_' +
          (vw.basis === 'post' ? 'Posting' : vw.basis === 'due' ? 'Due' : 'Bill') + '_Date_as_of_' + m.d;
        xlsxBook(book(m, vw), name);
        console.log(TAG, name);
      } catch (e) {
        console.error(TAG, 'download failed', e);
        try { frappe.show_alert({ message: 'Download failed - see the browser console.', indicator: 'red' }); } catch (e2) {}
      }
    }, true);
  }
  console.log(TAG, '04-Oct-2026 ready');
})(); } catch (e) { console.error('[AP v27] patch failed to initialise', e); }
