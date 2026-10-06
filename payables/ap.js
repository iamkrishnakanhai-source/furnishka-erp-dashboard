/* ── Furnishka shared libraries · load first · one synchronous XHR ───────────
   Finance-Lib-FiscalCalendar (FurnishkaFY) · Finance-Lib-NumberFormat (FurnishkaFmt)
   · Finance-Lib-SchemaGuard (FurnishkaGuard) · Finance-Lib-Typography (FurnishkaType)
   Reference: claude/06_SHARED_LIBRARIES.md. Synchronous by design — everything
   below depends on these being present. Each library is idempotent, so a double
   load is safe. Dependency order matters: Typography formats through NumberFormat. */
if (!window.FurnishkaFY) {
  try {
    var FKNAMES = ['Finance-Lib-FiscalCalendar', 'Finance-Lib-NumberFormat',
                   'Finance-Lib-SchemaGuard', 'Finance-Lib-Typography'];
    var fkx = new XMLHttpRequest();
    fkx.open('GET', '/api/method/frappe.client.get_list'
      + '?doctype=Client+Script'
      + '&fields='  + encodeURIComponent(JSON.stringify(['name', 'script']))
      + '&filters=' + encodeURIComponent(JSON.stringify([['name', 'in', FKNAMES]]))
      + '&limit_page_length=0', false);
    fkx.setRequestHeader('X-Frappe-CSRF-Token', frappe.csrf_token);
    fkx.send(null);
    var fkrows = (JSON.parse(fkx.responseText) || {}).message || [];
    FKNAMES.forEach(function (n) {
      var r = fkrows.find(function (v) { return v.name === n; });
      if (r && r.script) (0, eval)(r.script);
    });
  } catch (fke) { console.error('Furnishka library load failed:', fke); }
}
var FKMISSING = ['FurnishkaFY', 'FurnishkaFmt', 'FurnishkaGuard', 'FurnishkaType']
  .filter(function (g) { return !window[g]; });
if (FKMISSING.length || window.FurnishkaFY.degraded) {
  throw new Error('Furnishka shared libraries unavailable — AP dashboard halted rather than '
    + 'render a figure against a guessed fiscal year. Missing: '
    + (FKMISSING.join(', ') || 'none') + '. '
    + (window.FurnishkaFY ? (window.FurnishkaFY.health() || '') : ''));
}

/* ============================================================================
   Furnishka · Accounts Payable — Procure to Pay
   Custom HTML Block field: script
   v29 · 06-Oct-2026 · full rebuild, replaces the whole `script` field

   Reads:  frappe.desk.query_report.run on "AP Procure to Pay Master" (the server
           enforces AP Location access on every call — the dropdown here is not
           the control), plus two reads that exist only to prove that report:
             · GL Entry  sum(credit-debit) on the payable control, at the cut-off
             · Supplier  the enabled roster, so a vendor with no activity is
                         still accounted for rather than silently absent
   Writes: ONLY the dashboard-owned DocType "AP Location Classification Override",
           and only from the Categorise AP Location panel (managers). No ERP
           transaction, master or ledger is ever written from this file.

   Every link is revertable: tabs, filters, period and drill-downs use browser
   history, so Back returns to the previous view.

   ── WHAT THE REBUILD REMOVED ──────────────────────────────────────────────
   Eleven patch layers collapsed into one generation. Specifically gone:

   1. DOM SCRAPING. v28 built the merged supplier table by reading textContent
      out of the <td>s of the two tables it then hid, and parsing the formatted
      strings back into numbers. It joined that to a ledger read taken at a
      different moment, which is why it had to warn about a "capture gap" and
      refuse to call itself a reconciliation. The merged model is now computed
      from the report rows in the same snapshot as everything else, so the two
      sides cannot disagree and the warning has no reason to exist.
   2. FIVE download helpers and THREE XLSX writers (v24 x2, v25, v27, v28)
      collapsed to one dl(), one csvExport() carrying provenance, one xlsxBook().
   3. TWO header-filter engines (the original text/number one and v28's
      band-based one) collapsed to one, serving both filterable tables, with
      the union of both feature sets: multi-select with search and select-all,
      sign bands, min/max, blank/non-blank, and sort from the same popover.
   4. MONKEY-PATCHING. paintAgeing, paintTabs, paintDrawer, drill and
      scopeFilters were each reassigned by later layers; TABS was mutated and
      re-sorted twice. All of it is now declared once, in order.
   5. Two MutationObservers, a 3-second watchdog interval, a 120ms polling
      interval and a 45-second skeleton timeout. A single-generation render
      needs none of them: paint happens when data arrives.
   6. Dead renderers whose output was permanently display:none — the Overview
      block strip, subledger-to-GL table and exception census; the original
      ageing KPIs, chart, table and pager; the vendor ledger control card;
      v21's bridge and credits; v24's four-bucket table; the advance-booked and
      advance-not-booked cards (now the "Bills + advance" and "Advance only"
      classes of the one table).

   ── ONE SOURCE ────────────────────────────────────────────────────────────
   The report carries explicit row_type and grain_key, and a measure is
   populated ONLY on the grain that owns it, so nothing is summed across mixed
   grains. Grains used here:
     PO Doc          purchase order header   ordered / received / pending / invoiced
     PO Item         purchase order line     item bifurcation
     PR Doc          goods receipt header    received / billed / unbilled
     PR Item         goods receipt line      item bifurcation
     PI Doc          purchase invoice header gross / outstanding at cut-off / age
     PI Item         purchase invoice line   net, by route
     PI Tax          tax row    THE ONLY GRAIN STATUTORY TAX IS TOTALLED FROM
     PE Allocation   payment reference       allocated cash, bank, recon state
     PE Unallocated  payment header          the ERP unallocated field
     AP Aging        open ledger item        the ageing tab's bills / advances / journals
     AP GL Control   supplier                party ledger balance at cut-off
     AP GL Journal   GL entry                opening + manual journals
     AP Unapplied    supplier                DERIVED unapplied cash (reconciling figure)
     AP Opening Open Items   bill            carried in from before the period

   ── WHY THIS DASHBOARD IS SHAPED THIS WAY ─────────────────────────────────
   Verified read-only against production; each finding is reported, not hidden.

   1. AP subledger to GL closes at 0.00:
        AP control = open items + journals not linked to an invoice − unapplied cash
      Open items are measured from ledger movement at the cut-off. They are NOT
      read from Purchase Invoice.outstanding_amount, which is a current value
      and cannot answer a historical cut-off.
   2. Item-level GST does not reconcile to posted tax (FY 2026-27: item fields
      4,30,43,289.71 against tax rows 4,09,55,239.86 — 20,88,049.85 across 217
      invoices, worst PINV-26-04280 at 1,38,801.60 of item GST on 0.00 of posted
      tax). Statutory tax therefore comes only from the PI Tax grain. Item rows
      carry taxable value for bifurcation and never a tax total.
   3. is_tax_withholding_account is unreliable — set on 175 of 1,793 real TDS
      rows. The report resolves TDS by account name and treats the flag as a
      fallback.
   4. against_voucher is not maintained on payments (2,256 of 2,370 AP payment
      GL rows carry none). Settlement lives in Payment Entry Reference. All 114
      rows that do carry it also exist as references, so there is no double count.
   5. Purchase Invoice.unallocated_amount understates unapplied cash by
      2,80,38,738.35. The AP Unapplied grain is derived and is the reconciling
      figure; PE Unallocated is kept for drill-down only.
   6. No Payment Entry anywhere references a Purchase Order. Advances are real
      but none are PO-linked, so the advance card reports unapplied supplier
      cash rather than showing a PO advance of 0.00.
   7. 10 of 2,344 supplier payments carry a clearance date. Bank reconciliation
      is reported as an exception, not as a footnote.
   8. Party ledger sign: credit minus debit. Cr POSITIVE = owed to the supplier.
      Dr NEGATIVE = net receivable. A Dr balance is NEVER netted into the
      payable total — Schedule III requires it presented as an advance to
      suppliers under Other Current Assets. The native General Ledger prints
      Dr-positive, so a payable appears there as a negative closing balance:
      same magnitude, opposite sign. Stated on the face of the table.

   ── NUMBERS ───────────────────────────────────────────────────────────────
   Indian 2-2-3 grouping with paise on every figure, chip, axis and cell:
   24,74,86,314.43. Never abbreviated. Tolerance 0.50, shown, never forced.
   Dates dd-mmm-yyyy. Fiscal year from the Fiscal Year master, never arithmetic.
   ============================================================================ */

(() => {
'use strict';

/* ══ 1 · root, constants, primitives ═════════════════════════════════════ */

const ROOT = (typeof root_element !== 'undefined' && root_element)
  ? root_element
  : document.querySelector('[data-fk-block="ap-analysis"]')?.parentNode || document;

const el = (k) => ROOT.querySelector(`[data-fk="${k}"]`);
if (!el('timeline')) return;

const REPORT  = 'AP Procure to Pay Master';
const COMPANY = frappe.defaults.get_user_default('Company')
             || (frappe.boot.sysdefaults && frappe.boot.sysdefaults.company);
const AP_ACCOUNT = 'Accounts Payable - FT';
const OVERRIDE_DT = 'AP Location Classification Override';
const INCEPTION = '2000-04-01';
const TOL = 0.50;

const esc = (s) => frappe.utils.escape_html(String(s == null ? '' : s));
const num = (v) => { const n = Number(v); return isFinite(n) ? n : 0; };
const r2  = (v) => Math.round(num(v) * 100) / 100;

/* Indian 2-2-3 grouping, always exactly two decimals, never abbreviated */
const inr = (v, dp = 2) => window.FurnishkaFmt.inr(num(v), dp);
const rs  = (v) => '₹' + inr(v);
const cnt = (v) => window.FurnishkaFmt.inr(num(v), 0);
const pct = (a, b) => (!b ? '—' : window.FurnishkaFmt.pct(a / b * 100, 1));
const dash = (v) => Math.abs(num(v)) < 0.005 ? '–' : inr(v);
const plural = (n, one, many) => num(n) === 1 ? one : (many || one + 's');
/* "1 document", "2 documents" — counts are quoted all over this page and a
   stray plural on a single row looks like a bug in the figure next to it */
const countOf = (n, one, many) => cnt(n) + ' ' + plural(n, one, many);

const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const iso   = (d) => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')
                   + '-' + String(d.getDate()).padStart(2, '0');
const parse = (s) => { const d = new Date(String(s) + 'T00:00:00'); return isNaN(d) ? null : d; };
const ddmmmyyyy = (s) => { const d = parse(s); return d
  ? String(d.getDate()).padStart(2, '0') + '-' + MON[d.getMonth()] + '-' + d.getFullYear() : '—'; };
const today = () => frappe.datetime.get_today();
const daysAgo = (from, to) => Math.round((parse(to) - parse(from)) / 864e5);
const dayBefore = (d) => frappe.datetime.add_days(d, -1);
const q = (v) => encodeURIComponent(v);

/* The fiscal year comes from the Fiscal Year MASTER via FurnishkaFY: resolution
   is by containment (year_start_date <= d <= year_end_date), never by month
   arithmetic, and there is no year literal anywhere. Fiscal quarters divide the
   FY from year_start_date, so a short or shifted FY yields short quarters, not
   a wrong April anchor. fyOf returns null outside the calendar — deliberate. */
const FY = window.FurnishkaFY;
const fyOf    = (s)  => FY.fyOf(FY.iso(parse(s)) || FY.today());
const fyStart = (fy) => FY.fyStart(fy);
const fyEnd   = (fy) => FY.fyEnd(fy);
const fyLabel = (fy) => FY.fyLabel(fy);
const fqOf    = (s)  => FY.fqOf(FY.iso(parse(s)) || FY.today());

/* ── registrations · the report classifies, this only labels ─────────────── */
const LOCS = [
  { k: 'BLR', label: 'BLR — Karnataka', state: 'Karnataka', gstin: '29AAFCF4969D1ZC' },
  { k: 'JDP', label: 'JDP — Rajasthan', state: 'Rajasthan', gstin: '08AAFCF4969D1ZG' },
  { k: 'Unassigned', label: 'Unassigned', state: '', gstin: '' }
];
const locOf = (k) => LOCS.find((x) => x.k === k) || {};
const locLabel = (k) => locOf(k).label || (k || 'Combined');
/* Karnataka, Rajasthan or Unassigned — there is no Conflict bucket (report rule) */
const locChip = (k) => '<span class="fk-chip ' + (k === 'BLR' || k === 'JDP' ? 'c-ok' : 'c-warn')
  + '">' + esc(k === 'BLR' ? '◆ BLR' : k === 'JDP' ? '◆ JDP' : '▲ ' + (k || 'Unassigned')) + '</span>';

/* Route rule (master report): Inventory P2P = any invoice line on a PO or GRN,
   or update_stock, or a line posted to a Stock Received But Not Billed / Stock
   account. Everything else is a Direct Expense/Service Bill. Debit notes stay
   inside their route and are identified by status 'Return' (= is_return 1). */
const SOURCE_CLASSES = ['Inventory P2P', 'Direct Expense/Service Bill'];
const IS_P2P    = (r) => r.source_class === 'Inventory P2P';
const IS_RETURN = (r) => r.status === 'Return';
const PE_GRAINS = ['PE Allocation', 'PE Unallocated'];
const payVal = (r) => r.row_type === 'PE Unallocated' ? num(r.advance) : num(r.paid);

const CAN_CLASSIFY = () =>
  ['System Manager', 'Accounts Manager'].some((x) => (frappe.user_roles || []).indexOf(x) >= 0)
  || ['krishna.hari@furnishka.com', 'mahesh.majali@furnishka.com'].indexOf(frappe.session.user) >= 0;

/* ── ageing buckets · six, declared once, used by every consumer ─────────── */
const BANDS = [
  { k: 'notdue', label: 'Not due',  tone: 'var(--b0)' },
  { k: 'b1',     label: '1 – 30',   tone: 'var(--b1)' },
  { k: 'b2',     label: '31 – 60',  tone: 'var(--b2)' },
  { k: 'b3',     label: '61 – 90',  tone: 'var(--b3)' },
  { k: 'b4',     label: '91 – 180', tone: 'var(--b4)' },
  { k: 'b5',     label: 'Over 180', tone: 'var(--b5)' }
];
const bandIdx = (age) => { const g = num(age);
  return g <= 0 ? 0 : g <= 30 ? 1 : g <= 60 ? 2 : g <= 90 ? 3 : g <= 180 ? 4 : 5; };

/* The server ages too, on this filter. The grain returns the three raw dates as
   well, so the pills re-age client-side and never cost a refetch — but the
   filter is still sent, so a raw run of the report matches what is on screen.
   `server` is the value the report validates against; anything else falls back
   to Overdue Date. */
const AG_BASIS = {
  bill:    { label: 'Bill date',    field: 'supplier_invoice_date', server: 'Bill Date' },
  posting: { label: 'Posting date', field: 'date',                  server: 'Posting Date' },
  due:     { label: 'Due date',     field: 'required_by',           server: 'Due Date' }
};

/* the five classes of the one supplier table */
const CLASSES = {
  B:  { label: 'Bills only',        tone: 'c-info', group: 0 },
  BA: { label: 'Bills + advance',   tone: 'c-warn', group: 0 },
  A:  { label: 'Advance only',      tone: 'c-warn', group: 1 },
  L:  { label: 'Ledger only',       tone: 'c-crit', group: 2 },
  NA: { label: 'No activity',       tone: 'c-neu',  group: 3 }
};
const GROUP_NOTE = {
  1: 'Advance only — cash paid out, no bill booked. Not aged: there is no bill to be late against.',
  2: 'Ledger only — carries a balance on ' + AP_ACCOUNT + ' but no open item in the ageing grain.',
  3: 'No activity — enabled in the supplier master, nothing on the payable ledger.'
};

/* ══ 2 · read-only query layer ═══════════════════════════════════════════ */

const CACHE = new Map();

const scopeFilters = () => ({
  supplier: S.supplier, supplier_group: S.sg, cost_center: S.cc,
  department: S.dept, recon_state: S.recon, ap_location: S.loc
});

function runReport(over, opts) {
  const sc = (opts && opts.noScope) ? {} : scopeFilters();
  const rt = (over && over.row_type) || 'PI Doc';
  /* bank recon state exists only on payment grains — applied to PO / PR / PI /
     ledger grains it blanked the page */
  if (PE_GRAINS.indexOf(rt) < 0) sc.recon_state = '';
  const f = Object.assign({
    company: COMPANY,
    from_date: S.from, to_date: S.to, as_of: S.asOf,
    aging_basis: AG_BASIS[S.agBasis].server,
    row_type: 'PI Doc', source_class: '', supplier: '', supplier_group: '',
    cost_center: '', department: '', recon_state: '', ap_location: ''
  }, sc, over || {});
  const key = JSON.stringify(f);
  if (CACHE.has(key)) return CACHE.get(key);
  const p = frappe.call({
    method: 'frappe.desk.query_report.run',
    args: { report_name: REPORT, filters: f, ignore_prepared_report: 1 }
  }).then((r) => (r.message && r.message.result) || []);
  CACHE.set(key, p);
  return p;
}
const soft = (p, fb) => p.catch((e) => { console.warn('[AP]', e); return fb; });

/* the trial-balance side. GL Entry carries no AP Location / supplier group /
   department, so when the scope uses one the tie is reported not comparable
   rather than quietly compared against the wrong population. */
function glControl() {
  const f = { account: AP_ACCOUNT, is_cancelled: 0, company: COMPANY, posting_date: ['<=', S.asOf] };
  if (S.supplier) { f.party_type = 'Supplier'; f.party = S.supplier; }
  if (S.cc) f.cost_center = S.cc;
  return frappe.call({ method: 'frappe.client.get_list', args: {
    doctype: 'GL Entry', filters: f,
    fields: ['sum(credit-debit) as v', 'count(name) as n'], limit_page_length: 0 }
  }).then((r) => { const m = (r.message || [])[0] || {}; return { v: r2(m.v), n: num(m.n) }; })
    .catch(() => null);
}
const glComparable = () => !(S.loc || S.sg || S.dept);

/* ── the two ERP controls ─────────────────────────────────────────────────
   The master report is the SOURCE, because it is the only thing that carries
   AP Location, GSTIN, route and classification. It is not allowed to be its
   own proof. These two are what the dashboard is proved against:

     1. ERPNext's own statutory "Accounts Payable" report
     2. each vendor's own ledger — GL Entry on the payable account, by party

   Both are read COMPANY-WIDE regardless of what the dashboard is scoped to.
   That is deliberate: a per-supplier comparison is then always valid, because
   both sides describe the same party at the same cut-off and no scope
   mismatch can creep in. Only the grand-total comparison depends on the
   dashboard being unscoped, and the statement says so when it is not. */
function nativeAP() {
  return frappe.call({ method: 'frappe.desk.query_report.run', args: {
    report_name: 'Accounts Payable',
    filters: { company: COMPANY, report_date: S.asOf, party_type: 'Supplier',
      ageing_based_on: 'Due Date', range1: 30, range2: 60, range3: 90, range4: 180,
      based_on_payment_terms: 0 },
    ignore_prepared_report: 1 }
  }).then((r) => {
    const rows = (r.message && r.message.result) || [];
    const m = new Map();
    let n = 0;
    rows.forEach((x) => {
      const party = x.party || x.supplier;
      if (!party) return;                       /* the report's own total rows */
      const v = num(x.outstanding);
      if (!v) return;
      const e = m.get(party) || { out: 0, n: 0, name: x.party_name || x.supplier_name || party };
      e.out = r2(e.out + v); e.n++;
      m.set(party, e); n++;
    });
    return { byParty: m, rows: n, total: r2(Array.from(m.values()).reduce((a, e) => a + e.out, 0)) };
  }).catch((e) => { console.warn('[AP] native Accounts Payable report unavailable', e); return null; });
}

function vendorLedger() {
  return frappe.call({ method: 'frappe.client.get_list', args: {
    doctype: 'GL Entry',
    filters: { account: AP_ACCOUNT, is_cancelled: 0, company: COMPANY,
      party_type: 'Supplier', posting_date: ['<=', S.asOf] },
    fields: ['party', 'sum(credit-debit) as v', 'count(name) as n'],
    group_by: 'party', limit_page_length: 0 }
  }).then((r) => {
    const m = new Map();
    ((r.message) || []).forEach((x) => { if (x.party) m.set(x.party, { bal: r2(x.v), n: num(x.n) }); });
    return { byParty: m, total: r2(Array.from(m.values()).reduce((a, e) => a + e.bal, 0)) };
  }).catch((e) => { console.warn('[AP] vendor ledger read failed', e); return null; });
}

function supplierRoster() {
  return frappe.call({ method: 'frappe.client.get_list', args: {
    doctype: 'Supplier', fields: ['name', 'supplier_name'], filters: { disabled: 0 },
    limit_page_length: 0, order_by: 'supplier_name asc' }
  }).then((r) => { const m = new Map();
    ((r.message) || []).forEach((x) => m.set(x.name, x.supplier_name || x.name)); return m; })
    .catch(() => new Map());
}

/* ERP deep links — every figure reproduces in the native report */
const glUrl = (sup) => '/app/query-report/General Ledger?company=' + q(COMPANY)
  + '&from_date=' + q(INCEPTION) + '&to_date=' + q(S.asOf)
  + '&account=' + q('["' + AP_ACCOUNT + '"]')
  + '&group_by=' + q('Group by Voucher (Consolidated)')
  + (sup ? '&party_type=Supplier&party=' + q('["' + sup + '"]') : '');
const apUrl = (sup) => '/app/query-report/Accounts Payable?company=' + q(COMPANY)
  + '&report_date=' + q(S.asOf) + '&ageing_based_on=' + q('Due Date')
  + '&range1=30&range2=60&range3=90&range4=180&based_on_payment_terms=0'
  + (sup ? '&party_type=Supplier&party=' + q('["' + sup + '"]') : '');

/* ══ 3 · state ═══════════════════════════════════════════════════════════ */

const S = {
  from: null, to: null, asOf: null, preset: 'fy', grain: 'month',
  loc: '', cc: '', supplier: '', sg: '', dept: '', src: '', recon: '',
  tab: 'overview', tlMode: 'abs', tlHidden: new Set(),
  agBasis: 'due', agLevel: 'supplier',
  grView: 'docs', coView: 'cohort', payView: 'register',
  hf: {}, sort: {},
  search: { po: '', gr: '', inv: '', tp: '', ag: '', vb: '', vc: '', cls: '' },
  page: { po: 1, gr: 1, inv: 1, pay: 1, tp: 1, ag: 1, vb: 1, vcdn: 1, vcpe: 1, vcje: 1, cls: 1, dr: 1 },
  PAGE: 40,
  dims: { ccs: new Set(), groups: new Set(), depts: new Set(), banks: new Set(), suppliers: new Map() },
  D: {}, L: {}, A: {}, lazy: {},
  clsSel: new Set(), cls: null,
  drawer: null, drawerStack: [],
  away: false, busy: false, loaded: false
};
S.sort.ag = { key: 'group', dir: -1 };
S.sort.cls = { key: 'amount', dir: -1 };
S.sort.vb = { key: 'closing', dir: -1 };

const status = (msg, cls) => { const s = el('status'); if (!s) return;
  s.className = 'fk-status' + (cls ? ' ' + cls : ''); s.textContent = msg || ''; };

/* ══ 4 · period presets ══════════════════════════════════════════════════ */

const PRESETS = [
  { k: 'today',   label: 'Today',       range: () => [today(), today()] },
  { k: 'week',    label: 'This week',   range: () => { const d = parse(today());
      const s = new Date(d); s.setDate(d.getDate() - ((d.getDay() + 6) % 7));
      return [iso(s), today()]; } },
  { k: 'month',   label: 'This month',  range: () => { const d = parse(today());
      return [iso(new Date(d.getFullYear(), d.getMonth(), 1)), today()]; } },
  { k: 'lmonth',  label: 'Last month',  range: () => { const d = parse(today());
      return [iso(new Date(d.getFullYear(), d.getMonth() - 1, 1)),
              iso(new Date(d.getFullYear(), d.getMonth(), 0))]; } },
  { k: 'quarter', label: 'This fiscal quarter', range: () => {
      const fy = fyOf(today()), qn = fqOf(today()), sm = 3 + (qn - 1) * 3;
      const sy = fy + (sm > 11 ? 1 : 0), m = sm % 12;
      const e = iso(new Date(sy, m + 3, 0));
      return [iso(new Date(sy, m, 1)), e > today() ? today() : e]; } },
  { k: 'fy',      label: fyLabel(fyOf(today())) + ' to date',
      range: () => [fyStart(fyOf(today())), today()] },
  { k: 'fyfull',  label: fyLabel(fyOf(today())) + ' (full)',
      range: () => { const fy = fyOf(today()); return [fyStart(fy), fyEnd(fy)]; } },
  { k: 'lfy',     label: fyLabel(fyOf(today()) - 1),
      range: () => { const fy = fyOf(today()) - 1; return [fyStart(fy), fyEnd(fy)]; } },
  { k: 'all',     label: 'Since inception', range: () => [FY.inceptionStart(), today()] }
];
function applyPreset(k) {
  const p = PRESETS.find((x) => x.k === k); if (!p) return;
  const [f, t] = p.range();
  S.preset = k; S.from = f; S.to = t; S.asOf = t;
  el('from').value = f; el('to').value = t; el('asofin').value = t;
  paintPresets();
}
const paintPresets = () => el('presets').innerHTML = PRESETS.map((p) =>
  `<button type="button" data-p="${p.k}" aria-pressed="${p.k === S.preset}">${esc(p.label)}</button>`).join('');

/* ══ 5 · the run · one snapshot, one model ═══════════════════════════════ */

let REQ = 0;

async function run() {
  /* A newer filter or period selection supersedes an in-flight load. REQ stops
     a stale result from painting over a newer one. */
  S.busy = true;
  const my = ++REQ;
  status('Loading the master report…', 'busy');

  try {
    /* period grains — documents posted inside the window */
    status('Loading documents posted in the period…', 'busy');
    const [poDoc, prDoc, piDoc, piTax, peAlloc, peUnalloc, opening] = await Promise.all([
      soft(runReport({ row_type: 'PO Doc' }), []),
      soft(runReport({ row_type: 'PR Doc' }), []),
      soft(runReport({ row_type: 'PI Doc' }), []),
      soft(runReport({ row_type: 'PI Tax' }), []),
      soft(runReport({ row_type: 'PE Allocation' }), []),
      soft(runReport({ row_type: 'PE Unallocated' }), []),
      soft(runReport({ row_type: 'AP Opening Open Items' }), [])
    ]);
    if (my !== REQ) return;
    S.D = { poDoc, prDoc, piDoc, piTax, peAlloc, peUnalloc, opening };

    /* ledger grains — full history up to the cut-off — plus the two independent
       proofs. One Promise.all, so every figure on the page describes the same
       instant. This is what retires the capture-gap warning: there is no gap. */
    status('Loading the ledger position at the cut-off…', 'busy');
    const LEDGER = { from_date: INCEPTION, to_date: S.asOf };
    const [openAll, unapplied, journals, control, aging, gl, roster, nat, vled] = await Promise.all([
      soft(runReport(Object.assign({ row_type: 'PI Doc' }, LEDGER)), []),
      soft(runReport(Object.assign({ row_type: 'AP Unapplied' }, LEDGER)), []),
      soft(runReport(Object.assign({ row_type: 'AP GL Journal' }, LEDGER)), []),
      soft(runReport(Object.assign({ row_type: 'AP GL Control' }, LEDGER)), []),
      soft(runReport(Object.assign({ row_type: 'AP Aging' }, LEDGER)), []),
      glComparable() ? glControl() : Promise.resolve(null),
      supplierRoster(),
      nativeAP(),
      vendorLedger()
    ]);
    if (my !== REQ) return;
    S.L = { openAll, unapplied, journals, control, aging };
    S.gl = gl;
    S.roster = roster;
    S.nat = nat;
    S.vled = vled;

    S.lazy = {};
    S.loaded = true;
    derive();
    deriveAgeing();
    paintAll();
    status('');
  } catch (e) {
    console.error('[AP]', e);
    status('Could not load. ' + (e && e.message ? e.message : 'See the browser console.'), 'err');
  } finally { if (my === REQ) S.busy = false; }
}

/* item grains and the two off-window reads are heavy, so they load the first
   time a tab needs them and are cached against their own key */
async function lazy(key, rowType, over, label) {
  if (S.lazy[key]) return S.lazy[key];
  status('Loading ' + (label || rowType) + '…', 'busy');
  const rows = await soft(runReport(Object.assign({ row_type: rowType }, over || {})), []);
  S.lazy[key] = rows;
  status('');
  return rows;
}
const openingKey = () => 'vbOpen|' + dayBefore(S.from) + '|' + JSON.stringify(scopeFilters());
const vbOpening = () => { const cut = dayBefore(S.from);
  return lazy(openingKey(), 'AP GL Control',
    { from_date: INCEPTION, to_date: cut, as_of: cut },
    'the opening balance at ' + ddmmmyyyy(cut)); };
const vcUnallocated = () => lazy('vcPE|' + S.asOf, 'PE Unallocated',
  { from_date: INCEPTION, to_date: S.asOf }, 'unapplied payment documents');

/* ══ 6 · derivations ═════════════════════════════════════════════════════ */

const sum   = (rows, f) => (rows || []).reduce((a, r) => a + num(r[f]), 0);
const where = (rows, fn) => (rows || []).filter(fn);

function groupSumBy(rows, keyFn, valFn) {
  const m = new Map();
  (rows || []).forEach((r) => {
    const k = keyFn(r);
    const e = m.get(k) || { key: k, v: 0, n: 0, rows: [] };
    e.v += valFn(r); e.n++; e.rows.push(r);
    m.set(k, e);
  });
  return Array.from(m.values()).sort((a, b) => Math.abs(b.v) - Math.abs(a.v));
}
const groupSum = (rows, key, field) => groupSumBy(rows, (r) => r[key] || '—', (r) => num(r[field]));

function bankSplit(rows, field) {
  const m = new Map();
  Array.from(S.dims.banks).sort().forEach((b) => m.set(b, { key: b, v: 0, n: 0, rows: [] }));
  (rows || []).forEach((r) => {
    const b = r.bank || 'No bank account';
    if (!m.has(b)) m.set(b, { key: b, v: 0, n: 0, rows: [] });
    const e = m.get(b); e.v += num(r[field]); e.n++; e.rows.push(r);
  });
  return Array.from(m.values());
}

function derive() {
  const D = S.D, L = S.L;

  /* ── procurement, inventory route only ──────────────────────────────── */
  const po = D.poDoc;
  D.po = {
    rows: po, n: po.length,
    value: sum(po, 'gross'), received: sum(po, 'received'),
    pending: sum(po, 'pending_receipt'), invoiced: sum(po, 'invoiced'),
    closed: where(po, (r) => r.status === 'Closed' || r.status === 'Completed'),
    open:   where(po, (r) => r.status !== 'Closed' && r.status !== 'Completed')
  };
  D.po.closedValue = sum(D.po.closed, 'gross');
  D.po.openValue   = sum(D.po.open, 'gross');
  D.po.byStatus    = groupSum(po, 'status', 'gross');

  /* ── goods receipt, net of receipt returns ──────────────────────────── */
  const pr = D.prDoc;
  D.pr = {
    rows: pr, n: pr.length,
    value: sum(pr, 'gross'), billed: sum(pr, 'invoiced'), unbilled: sum(pr, 'pending_bill'),
    fully:   where(pr, (r) => r.recon_state === 'Fully Billed'),
    open:    where(pr, (r) => r.recon_state !== 'Fully Billed'),
    returns: where(pr, IS_RETURN)
  };
  D.pr.fullyValue = sum(D.pr.fully, 'gross');

  /* ── invoices, every route ──────────────────────────────────────────── */
  const pi = D.piDoc;
  const returns = where(pi, IS_RETURN), normal = where(pi, (r) => !IS_RETURN(r));
  D.pi = {
    rows: pi, n: pi.length, returns, normal,
    gross: sum(pi, 'gross'), taxable: sum(pi, 'taxable'),
    normalV: sum(normal, 'gross'), normalN: normal.length,
    returnV: sum(returns, 'gross'), returnN: returns.length,
    byRoute: groupSum(pi, 'source_class', 'gross'),
    p2p: where(pi, IS_P2P), direct: where(pi, (r) => !IS_P2P(r))
  };
  D.pi.p2pV = sum(D.pi.p2p, 'gross');
  D.pi.directV = sum(D.pi.direct, 'gross');
  D.roundOff = sum(pi, 'round_off');

  /* ── statutory tax: the PI Tax grain and nothing else ───────────────── */
  const tx = D.piTax;
  D.tax = {
    rows: tx,
    cgst: sum(tx, 'cgst'), sgst: sum(tx, 'sgst'), igst: sum(tx, 'igst'),
    tds: sum(tx, 'tds'), other: sum(tx, 'other_charges')
  };
  D.tax.gst = D.tax.cgst + D.tax.sgst + D.tax.igst;

  /* ── settlement ─────────────────────────────────────────────────────── */
  D.allocated    = sum(D.peAlloc, 'paid');
  D.unallocField = sum(D.peUnalloc, 'advance');
  D.cashOut      = D.allocated + D.unallocField;
  D.payN = new Set(D.peAlloc.map((r) => r.pe_id).concat(D.peUnalloc.map((r) => r.pe_id))).size;
  D.bankAlloc = bankSplit(D.peAlloc, 'paid');
  D.bankAdv   = bankSplit(D.peUnalloc, 'advance');
  D.unrec     = where(D.peAlloc, (r) => r.recon_state === 'Bank Unreconciled');
  D.unrecV    = sum(D.unrec, 'paid');

  /* ── ledger position at the cut-off ─────────────────────────────────── */
  const openItems = where(L.openAll, (r) => Math.abs(num(r.outstanding)) > 0.005);
  L.open       = openItems;
  L.openTotal  = sum(openItems, 'outstanding');
  L.unappliedV = sum(L.unapplied, 'advance');
  L.jrNonInv   = where(L.journals, (r) => !r.pi_id);
  L.jrOnInv    = where(L.journals, (r) => r.pi_id);
  L.jrNonInvV  = sum(L.jrNonInv, 'outstanding');
  L.jrOnInvV   = sum(L.jrOnInv, 'outstanding');
  L.controlV   = sum(L.control, 'outstanding');
  L.ppRows     = where(L.aging, (r) => r.entry_type === 'Payment Entry' && num(r.outstanding) > 0.005);
  L.ppV        = sum(L.ppRows, 'outstanding');
  /* NOT + L.ppV. The AP Unapplied grain carries every Payment Entry target row
     with advance = -amount, so a payment sitting on the payable side arrives as
     a negative advance and '- unapplied' has already added it back. Adding it
     again counts it twice. Confirmed against the report source, 06-Oct-2026. */
  L.computed   = L.openTotal + L.jrNonInvV - L.unappliedV;
  L.variance   = L.computed - L.controlV;
  L.ok         = Math.abs(L.variance) <= TOL;
  /* the AP control carries no accounting dimension, so the tie is reported
     not applicable when the scope uses one rather than quietly compared */
  L.scopable   = !(S.cc || S.dept);
  /* the grain returns EVERY pre-period invoice with its outstanding patched from
     the payment ledger, so a settled one comes back carrying 0.00 — counting
     rows would overstate what is actually carried in */
  L.openingOpen = where(D.opening, (r) => Math.abs(num(r.outstanding)) > 0.005);
  L.openingV    = sum(L.openingOpen, 'outstanding');

  /* ── routes: payments and open items follow the invoice they belong to ── */
  S.ROUTE = new Map((L.openAll || []).map((r) => [r.grain_key, r.source_class]));
  D.p2pAlloc  = where(D.peAlloc, (r) => S.ROUTE.get(r.pi_id) === 'Inventory P2P');
  D.p2pAllocV = sum(D.p2pAlloc, 'paid');
  D.dirAllocV = sum(where(D.peAlloc, (r) => S.ROUTE.get(r.pi_id) === 'Direct Expense/Service Bill'), 'paid');
  L.openP2P  = where(openItems, IS_P2P);              L.openP2PV = sum(L.openP2P, 'outstanding');
  L.openDir  = where(openItems, (r) => !IS_P2P(r));   L.openDirV = sum(L.openDir, 'outstanding');
  /* inventory vendors: a PO, a GRN or an Inventory P2P invoice */
  D.invVendors = new Set([].concat(D.poDoc || [], D.prDoc || [], where(L.openAll, IS_P2P))
    .map((r) => r.supplier));
  L.unappliedInv  = where(L.unapplied, (r) => D.invVendors.has(r.supplier));
  L.unappliedInvV = sum(L.unappliedInv, 'advance');
  D.bankAdvInv = bankSplit(where(D.peUnalloc, (r) => D.invVendors.has(r.supplier)), 'advance');

  D.series = buildSeries();
  D.cohort = buildCohort();
  D.truePurchase = buildTruePurchase();
}

/* ── the merged ageing model ──────────────────────────────────────────────
   Computed from the AP Aging grain (every open item on the payable account at
   the cut-off), the AP GL Journal grain, the AP GL Control grain and the
   Supplier roster. All four arrive in the same Promise.all as everything else.

   The AP Aging grain decomposes exactly:
     party ledger = bills + debit notes − advances + payments-on-account + journals
   so the Delta column is a genuine cross-grain tie between AP Aging and
   AP GL Control, not the incomplete subtraction v28 performed on scraped text.

   Switching the ageing basis re-ages client-side from dates the grain already
   carries, so it never refetches. */
function deriveAgeing() {
  if (!S.L.aging) return;
  const basisField = AG_BASIS[S.agBasis].field;
  const byId = new Map();
  const base = (id, nm, src) => {
    const o = { id, nm: nm || id, cls: 'NA', src,
      b: BANDS.map(() => 0), billTot: 0, billN: 0,
      dn: 0, dnN: 0, adv: 0, advN: 0, pp: 0, jv: 0, led: null,
      bills: [], locs: new Set() };
    byId.set(id, o); return o;
  };
  (S.roster || new Map()).forEach((nm, id) => base(id, nm, 'master'));

  const agTot = { b: BANDS.map(() => 0), billTot: 0, billN: 0, dn: 0, dnN: 0,
                  adv: 0, advN: 0, pp: 0, jv: 0, led: 0, cr: 0, dr: 0, drN: 0 };
  const allBills = [];

  (S.L.aging || []).forEach((r) => {
    const id = r.supplier; if (!id) return;
    const o = byId.get(id) || base(id, r.supplier_name || id, 'aging');
    if (r.supplier_name) o.nm = r.supplier_name;
    if (r.ap_location) o.locs.add(r.ap_location);
    const v = num(r.outstanding), vt = r.entry_type;

    if (vt === 'Purchase Invoice') {
      if (Math.abs(v) < 0.005) return;
      const from = r[basisField] || r.date || S.asOf;
      const age = daysAgo(from, S.asOf);
      const i = bandIdx(age);
      const bill = { sup: id, supName: o.nm, no: r.supplier_invoice_no || r.entry_id,
        vno: r.entry_id, from, due: r.required_by || '', age, amt: v,
        band: v < 0 ? 'Debit note' : BANDS[i].label, loc: r.ap_location || 'Unassigned' };
      o.bills.push(bill); allBills.push(bill);
      if (v < 0) { o.dn += v; o.dnN++; agTot.dn += v; agTot.dnN++; }
      else { o.b[i] += v; o.billTot += v; o.billN++; agTot.b[i] += v; agTot.billTot += v; agTot.billN++; }
    } else if (vt === 'Payment Entry') {
      if (v < -0.005) { o.adv += -v; o.advN++; agTot.adv += -v; agTot.advN++; }
      else if (v > 0.005) { o.pp += v; agTot.pp += v; }
    } else if (vt === 'Journal Entry') {
      o.jv += v; agTot.jv += v;
    }
  });

  /* journals not linked to an invoice — the Overview bridge's own grain, kept
     as an explicit cross-check of the ageing grain's journal rows */
  const jrGl = new Map();
  (S.L.jrNonInv || []).forEach((r) => {
    if (!r.supplier) return;
    jrGl.set(r.supplier, num(jrGl.get(r.supplier)) + num(r.outstanding));
  });

  /* party ledger — the AP GL Control grain, one row per supplier */
  (S.L.control || []).forEach((r) => {
    const id = r.supplier; if (!id) return;
    const o = byId.get(id) || base(id, r.supplier_name || id, 'control');
    o.led = num(o.led) + num(r.outstanding);
  });

  const tieable = !(S.loc || S.src || S.cc || S.dept || S.sg || S.recon);
  const natBy = S.nat ? S.nat.byParty : null;
  const vlBy  = S.vled ? S.vled.byParty : null;
  /* a party present in a control but in no grain still has to appear, or the
     dashboard could hide a balance simply by not knowing about it */
  [natBy, vlBy].forEach((m) => { if (!m) return;
    m.forEach((_v, id) => { if (!byId.has(id)) base(id, (S.roster && S.roster.get(id)) || id, 'control'); }); });

  const rows = [];
  byId.forEach((o) => {
    o.jrGl = jrGl.has(o.id) ? r2(jrGl.get(o.id)) : null;
    o.billTot = r2(o.billTot); o.dn = r2(o.dn); o.adv = r2(o.adv);
    o.pp = r2(o.pp); o.jv = r2(o.jv);
    o.b = o.b.map(r2);
    o.net = r2(o.billTot + o.dn - o.adv + o.pp + o.jv);
    if (o.led !== null) { o.led = r2(o.led); o.delta = r2(o.led - o.net); }
    else { o.delta = null; }
    /* the two ERP controls, and the dashboard's distance from each */
    o.nat = natBy && natBy.has(o.id) ? natBy.get(o.id).out : (natBy ? 0 : null);
    o.vl  = vlBy  && vlBy.has(o.id)  ? vlBy.get(o.id).bal  : (vlBy ? 0 : null);
    o.dNat = (o.nat === null || !tieable) ? null : r2(o.net - o.nat);
    o.dVl  = (o.vl  === null || !tieable) ? null : r2(o.net - o.vl);
    o.allocatable = o.adv > 0.005 ? r2(Math.min(o.adv, Math.max(o.billTot, 0))) : null;
    o.ap_location = o.locs.size === 1 ? Array.from(o.locs)[0] : o.locs.size ? 'Mixed' : '';
    o.bills.sort((x, y) => y.age - x.age || Math.abs(y.amt) - Math.abs(x.amt));

    const hasBills = Math.abs(o.billTot) > 0.005 || Math.abs(o.dn) > 0.005;
    const hasAdv   = o.adv > 0.005;
    const hasLed   = (o.led !== null && Math.abs(o.led) > 0.005)
                  || (o.vl !== null && Math.abs(o.vl) > 0.005);
    o.cls = hasBills ? (hasAdv ? 'BA' : 'B') : hasAdv ? 'A' : hasLed ? 'L' : 'NA';
    o.action = agAction(o);
    rows.push(o);
  });

  rows.forEach((o) => {
    if (o.led === null) return;
    agTot.led += o.led;
    if (o.led > 0.5) agTot.cr += o.led;
    else if (o.led < -0.5) { agTot.dr += o.led; agTot.drN++; }
  });

  /* the advance split is the Schedule III distinction, not two populations:
     one supplier carries one advance, and which line it sits on depends only
     on whether a bill has been booked against it */
  const advHeld = r2(rows.filter((o) => o.cls === 'BA').reduce((a, o) => a + o.adv, 0));
  const advOnly = r2(rows.filter((o) => o.cls === 'A').reduce((a, o) => a + o.adv, 0));
  const bucketSum = r2(agTot.b.reduce((a, v) => a + v, 0));
  const net = r2(agTot.billTot + agTot.dn - agTot.adv + agTot.pp + agTot.jv);
  const ledTot = r2(agTot.led);
  const glTot = S.gl ? S.gl.v : null;

  const natTot = S.nat ? S.nat.total : null;
  const vlTot  = S.vled ? S.vled.total : null;
  const offNat = rows.filter((o) => o.dNat !== null && Math.abs(o.dNat) > TOL).length;
  const offVl  = rows.filter((o) => o.dVl  !== null && Math.abs(o.dVl)  > TOL).length;

  S.A = {
    rows, bills: allBills, bands: agTot.b.map(r2),
    nat: natTot, natRows: S.nat ? S.nat.rows : 0, vl: vlTot,
    offNat, offVl,
    /* Both controls hold the WHOLE supplier. Any filter that drops rows inside a
       supplier — AP Location (the report drops rows after classifying them),
       route, cost centre, department, group, bank state — makes a per-supplier
       comparison meaningless, because the two sides describe different
       populations. A supplier filter is fine: it selects which vendors to show,
       and each one shown still carries all of its rows. */
    tieable: !(S.loc || S.src || S.cc || S.dept || S.sg || S.recon),
    unscoped: !(S.loc || S.sg || S.dept || S.cc || S.supplier || S.src || S.recon),
    billTot: r2(agTot.billTot), billN: agTot.billN,
    dn: r2(agTot.dn), dnN: agTot.dnN,
    adv: r2(agTot.adv), advN: agTot.advN, advHeld, advOnly,
    pp: r2(agTot.pp), jv: r2(agTot.jv), jrGl: r2(S.L.jrNonInvV),
    led: ledTot, cr: r2(agTot.cr), dr: r2(agTot.dr), drN: agTot.drN,
    bucketSum, net, gl: glTot, glN: S.gl ? S.gl.n : 0,
    bandN: BANDS.map((b, i) => allBills.filter((x) => x.amt > 0 && bandIdx(x.age) === i).length),
    off: rows.filter((o) => o.delta !== null && Math.abs(o.delta) > TOL).length,
    nL: rows.filter((o) => o.cls === 'L').length,
    nNA: rows.filter((o) => o.cls === 'NA').length,
    basis: AG_BASIS[S.agBasis].label
  };
}

/* the one action a row earns, from its own shape */
function agAction(o) {
  /* a disagreement with either ERP control outranks everything else: until it
     is explained, the row's own figures cannot be acted on */
  if ((o.dNat !== null && Math.abs(o.dNat) > TOL)
      || (o.dVl !== null && Math.abs(o.dVl) > TOL)) return 'Investigate';
  if (o.adv > 0.005 && o.billTot > 0.005) return 'Reconcile';
  if (o.adv > 0.005) return 'Obtain bill';
  if (o.cls === 'L' || (o.delta !== null && Math.abs(o.delta) > TOL)) return 'Review';
  return '';
}

/* ── timeline and cohort ────────────────────────────────────────────────── */

const SERIES = [
  { k: 'ordered',  label: 'Ordered',  colour: 'var(--s1)' },
  { k: 'received', label: 'Received', colour: 'var(--s3)' },
  { k: 'invoiced', label: 'Invoiced', colour: 'var(--s2)' },
  { k: 'paid',     label: 'Paid',     colour: 'var(--s4)' }
];
function bucketKey(s) {
  const d = parse(s); if (!d) return null;
  if (S.grain === 'day') return iso(d);
  if (S.grain === 'week') { const x = new Date(d); x.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return iso(x); }
  return iso(new Date(d.getFullYear(), d.getMonth(), 1));
}
function bucketLabel(k) {
  const d = parse(k); if (!d) return k;
  if (S.grain === 'month') return MON[d.getMonth()] + " '" + String(d.getFullYear()).slice(2);
  return String(d.getDate()).padStart(2, '0') + ' ' + MON[d.getMonth()];
}
function buildSeries() {
  const D = S.D, keys = new Set(), acc = new Map();
  const add = (dt, f, v) => {
    const k = bucketKey(dt); if (!k) return;
    keys.add(k);
    let r = acc.get(k);
    if (!r) { r = { k }; SERIES.forEach((s) => r[s.k] = 0); acc.set(k, r); }
    r[f] += num(v);
  };
  (D.poDoc || []).forEach((r) => add(r.date, 'ordered', r.gross));
  (D.prDoc || []).forEach((r) => add(r.date, 'received', r.gross));
  where(D.piDoc, IS_P2P).forEach((r) => add(r.date, 'invoiced', r.gross));   /* inventory chain only */
  (D.p2pAlloc || []).forEach((r) => add(r.date, 'paid', r.paid));

  const out = Array.from(keys).sort().map((k) => acc.get(k));
  if (S.tlMode === 'cum') {
    const runTot = {}; SERIES.forEach((s) => runTot[s.k] = 0);
    out.forEach((r) => SERIES.forEach((s) => { runTot[s.k] += r[s.k]; r[s.k] = runTot[s.k]; }));
  }
  return out;
}

/* PO cohort: select the orders RELEASED in each bucket, then follow those exact
   orders through to the cut-off — receipts and invoices that happened later are
   still counted against the month the order was released in. */
function buildCohort() {
  const m = new Map();
  (S.D.poDoc || []).forEach((r) => {
    const k = bucketKey(r.date); if (!k) return;
    let e = m.get(k);
    if (!e) { e = { k, ordered: 0, received: 0, pendingReceipt: 0, invoiced: 0,
                    pendingBill: 0, paid: 0, n: 0, rows: [] }; m.set(k, e); }
    e.ordered        += num(r.ordered);
    e.received       += num(r.received);
    e.pendingReceipt += num(r.pending_receipt);
    e.invoiced       += num(r.invoiced);
    e.pendingBill    += num(r.pending_bill);
    e.n++; e.rows.push(r);
  });
  return Array.from(m.values()).sort((a, b) => a.k < b.k ? -1 : 1);
}

/* Supplier true purchase — invoice-led, reconciles to the invoice register */
function buildTruePurchase() {
  const m = new Map();
  const touch = (sup, name) => {
    let e = m.get(sup);
    if (!e) { e = { supplier: sup, name: name || sup, p2p: 0, directExp: 0, debitNotes: 0,
                    net: 0, paid: 0, outstanding: 0, gst: 0, tds: 0, invRows: [] };
              m.set(sup, e); }
    return e;
  };
  (S.D.piDoc || []).forEach((r) => {
    const e = touch(r.supplier, r.supplier_name), g = num(r.gross);
    if (IS_P2P(r)) e.p2p += g; else e.directExp += g;
    if (IS_RETURN(r)) e.debitNotes += g;   /* memo — already inside its route */
    e.net += g;
    e.gst += num(r.cgst) + num(r.sgst) + num(r.igst);
    e.tds += num(r.tds);
    e.invRows.push(r);
  });
  (S.D.peAlloc || []).forEach((r) => { touch(r.supplier, r.supplier_name).paid += num(r.paid); });
  (S.L.open || []).forEach((r) => { touch(r.supplier, r.supplier_name).outstanding += num(r.outstanding); });
  return Array.from(m.values()).sort((a, b) => Math.abs(b.net) - Math.abs(a.net));
}

/* ── vendor balances · opening + movement = closing ─────────────────────── */
function splitByType(rows) {
  const m = new Map();
  (rows || []).forEach((r) => {
    if (!r.supplier) return;
    let e = m.get(r.supplier);
    if (!e) { e = { total: 0, PI: 0, RET: 0, JE: 0, PE: 0, OTH: 0 }; m.set(r.supplier, e); }
    const v = num(r.outstanding), t = r.entry_type;
    e.total += v;
    if (t === 'Purchase Invoice') { if (IS_RETURN(r)) e.RET += v; else e.PI += v; }
    else if (t === 'Journal Entry') e.JE += v;
    else if (t === 'Payment Entry') e.PE += v;
    else e.OTH += v;
  });
  return m;
}
function vbModel() {
  const L = S.L, zero = { total: 0, PI: 0, RET: 0, JE: 0, PE: 0, OTH: 0 };
  const now = splitByType(L.control), op = splitByType(S.lazy[openingKey()] || []);
  const names = new Map();
  (L.control || []).forEach((r) => { if (r.supplier) names.set(r.supplier, r.supplier_name || r.supplier); });
  (S.lazy[openingKey()] || []).forEach((r) => {
    if (r.supplier && !names.has(r.supplier)) names.set(r.supplier, r.supplier_name || r.supplier); });
  const unap = new Map();
  (L.unapplied || []).forEach((r) => unap.set(r.supplier, num(unap.get(r.supplier)) + num(r.advance)));
  const glRows = new Map();
  (L.control || []).forEach((r) => glRows.set(r.supplier, num(glRows.get(r.supplier)) + 1));

  const out = [];
  names.forEach((nm, sup) => {
    const n = now.get(sup) || zero, o = op.get(sup) || zero;
    const e = { supplier: sup, name: nm,
      opening: r2(o.total), openJE: r2(o.JE), openPI: r2(o.PI + o.RET),
      openPE: r2(o.PE), openOther: r2(o.OTH),
      billed: r2(n.PI - o.PI), dn: r2(n.RET - o.RET), jr: r2(n.JE - o.JE),
      pay: r2(n.PE - o.PE), other: r2(n.OTH - o.OTH),
      closing: r2(n.total), unap: r2(unap.get(sup) || 0), n: num(glRows.get(sup)) };
    e.movement = r2(e.billed + e.dn + e.jr + e.pay + e.other);
    e.built = r2(e.opening + e.movement);
    e.variance = r2(e.built - e.closing);
    e.ties = Math.abs(e.variance) <= TOL;
    out.push(e);
  });
  return out.sort((a, b) => Math.abs(b.closing) - Math.abs(a.closing));
}

/* ── vendor credits · the three doors a credit arrives through ──────────── */
const vcReturns = () => where(S.L.openAll, (r) => IS_RETURN(r) && Math.abs(num(r.outstanding)) > 0.005)
  .map((r) => {
    const raised = Math.abs(num(r.gross)), avail = Math.abs(num(r.outstanding));
    return Object.assign({}, r, { vc_raised: raised, vc_avail: avail,
      vc_applied: r2(raised - avail), vc_age: daysAgo(r.date, S.asOf),
      /* a credit raised in paise but carried as a round rupee is a migration
         artefact, not an ageing problem — flagged as its own thing */
      vc_round: Math.abs(raised - avail) < 0.005
        && Math.abs(avail - Math.round(avail)) < 0.005
        && Math.abs(raised - Math.round(raised)) > 0.004 });
  }).sort((a, b) => b.vc_avail - a.vc_avail);

function vcModel() {
  const oldest = new Map();
  (S.L.open || []).forEach((r) => {
    if (IS_RETURN(r)) return;
    const c = oldest.get(r.supplier);
    if (!c || String(r.date) < String(c.date)) oldest.set(r.supplier, { date: r.date, id: r.grain_key });
  });
  const dn = vcReturns();
  const pe = (S.lazy['vcPE|' + S.asOf] || []).filter((r) => num(r.advance) > 0.005)
    .map((r) => Object.assign({}, r, { vc_age: daysAgo(r.date, S.asOf),
      vc_hint: oldest.get(r.supplier) || null }))
    .sort((a, b) => num(b.advance) - num(a.advance));
  const je = where(S.L.jrNonInv, (r) => num(r.outstanding) > 0.005)
    .sort((a, b) => num(b.outstanding) - num(a.outstanding));
  return { dn, pe, je,
    dnTot: r2(dn.reduce((a, r) => a + r.vc_avail, 0)),
    peTot: r2(pe.reduce((a, r) => a + num(r.advance), 0)),
    jeTot: r2(je.reduce((a, r) => a + num(r.outstanding), 0)) };
}

/* ══ 7 · one header-filter engine, Excel-style ═══════════════════════════
   Serves every filterable table on the page. The union of what the two engines
   this replaces could do: multi-select with search and select-all for text,
   sign bands plus min/max for amounts, blank / non-blank for both, and sort
   from the same popover. A column opts out with nofilter. */

const hfOf = (scope) => (S.hf[scope] = S.hf[scope] || {});
const hfCount = (scope) => Object.keys(S.hf[scope] || {}).length;

function hfText(c, r) {
  if (!c) return '';
  const v = c.plain ? c.plain(r) : r[c.k];
  return v === null || v === undefined || v === '' ? '(blank)' : String(v);
}
const hfNum = (c, r) => c.raw ? c.raw(r) : r[c.k];

function hfPass(rows, scope, cols) {
  const F = S.hf[scope] || {};
  const keys = Object.keys(F);
  if (!keys.length) return rows;
  return rows.filter((r) => keys.every((k) => {
    const f = F[k], c = cols.find((x) => x.k === k);
    if (!c) return true;
    if (f.type === 'set') return f.values.has(hfText(c, r));
    const raw = hfNum(c, r);
    const blank = raw === null || raw === undefined || raw === '';
    const v = num(raw);
    if (f.bands) {
      const band = blank ? 'blank' : v > 0.005 ? 'pos' : v < -0.005 ? 'neg' : 'zero';
      if (!f.bands[band]) return false;
    }
    if (f.min !== null && f.min !== undefined && !(!blank && v >= f.min)) return false;
    if (f.max !== null && f.max !== undefined && !(!blank && v <= f.max)) return false;
    return true;
  }));
}

function hfSort(rows, scope, cols) {
  const s = S.sort[scope];
  if (!s || !s.key) return rows;
  const c = (cols || []).find((x) => x.k === s.key);
  const d = s.dir < 0 ? -1 : 1;
  return rows.slice().sort((a, b) => {
    if (c && c.num) return (num(hfNum(c, a)) - num(hfNum(c, b))) * d;
    const av = c ? hfText(c, a) : String(a[s.key] == null ? '' : a[s.key]);
    const bv = c ? hfText(c, b) : String(b[s.key] == null ? '' : b[s.key]);
    return av.localeCompare(bv, undefined, { numeric: true }) * d;
  });
}

const closeHf = () => ROOT.querySelectorAll('.fk-hf').forEach((x) => x.remove());
let HF_OPENED = 0;
const OPT_CAP = 400;

function hfOpen(scope, key, cols, rows, th, repaint) {
  closeHf();
  const c = cols.find((x) => x.k === key); if (!c) return;
  const cur = (S.hf[scope] || {})[key];
  const pop = document.createElement('div');
  pop.className = 'fk-hf';
  pop.setAttribute('role', 'dialog');
  pop.setAttribute('aria-label', 'Filter ' + c.label);

  let body;
  if (c.num) {
    const bands = [['pos', 'Positive'], ['neg', 'Negative'], ['zero', 'Zero'], ['blank', 'Blank']];
    body = `<div>${bands.map((b) =>
      `<label class="hf-b"><input type="checkbox" class="hf-bd" value="${b[0]}"` +
      `${!cur || !cur.bands || cur.bands[b[0]] ? ' checked' : ''}> ${b[1]}</label>`).join('')}</div>`
      + `<label class="hf-l">Between<span class="hf-nums">` +
        `<input type="number" step="0.01" data-hf="min" placeholder="min" value="${cur && cur.min != null ? esc(cur.min) : ''}">` +
        `<input type="number" step="0.01" data-hf="max" placeholder="max" value="${cur && cur.max != null ? esc(cur.max) : ''}">` +
      `</span></label>`;
  } else {
    const m = new Map();
    rows.forEach((r) => { const v = hfText(c, r); m.set(v, num(m.get(v)) + 1); });
    const vals = Array.from(m.keys()).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const shown = vals.slice(0, OPT_CAP);
    body = `<input type="search" data-hf="q" placeholder="Search values…" aria-label="Search values">`
      + `<div class="hf-sel"><button type="button" class="fk-btn fk-btn-ghost fk-xs" data-hf="all">Select all</button>`
      + `<button type="button" class="fk-btn fk-btn-ghost fk-xs" data-hf="none">Clear</button></div>`
      + `<div class="hf-list">${shown.map((v) =>
          `<label data-v="${esc(v.toLowerCase())}"><input type="checkbox" class="hf-cb" value="${esc(v)}"` +
          `${!cur || !cur.values || cur.values.has(v) ? ' checked' : ''}>` +
          `<span>${esc(v)}</span><span class="n">${cnt(m.get(v))}</span></label>`).join('')}</div>`
      + (vals.length > shown.length
          ? `<div class="hf-cap">${cnt(vals.length)} values — showing the first ${cnt(OPT_CAP)}; refine the search</div>` : '');
  }

  pop.innerHTML =
    `<div class="hf-h"><b>${esc(c.label)}</b>` +
    `<button type="button" class="fk-btn fk-btn-ghost fk-xs" data-hf="x" aria-label="Close">✕</button></div>` +
    `<div class="hf-sort"><button type="button" class="fk-btn fk-xs" data-hf="asc">▲ Sort ascending</button>` +
    `<button type="button" class="fk-btn fk-xs" data-hf="desc">▼ Sort descending</button></div>` +
    body +
    `<div class="hf-f"><button type="button" class="fk-btn fk-btn-ghost fk-xs" data-hf="reset">Remove filter</button>` +
    `<button type="button" class="fk-btn fk-btn-primary fk-xs" data-hf="apply">Apply</button></div>`;

  ROOT.querySelector('[data-fk-block="ap-analysis"]').appendChild(pop);
  th.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  const rc = th.getBoundingClientRect();
  const ph = Math.min(pop.offsetHeight || 340, window.innerHeight - 16);
  pop.style.maxHeight = (window.innerHeight - 16) + 'px';
  let top = rc.bottom + 4;
  if (top + ph > window.innerHeight - 8) top = Math.max(8, rc.top - ph - 4);
  pop.style.top = top + 'px';
  pop.style.left = Math.max(8, Math.min(window.innerWidth - 308, rc.left)) + 'px';
  /* the scrollIntoView above fires a late scroll event — do not let it close this */
  HF_OPENED = Date.now();

  const done = () => { closeHf(); repaint(); };
  const qi = pop.querySelector('[data-hf="q"]');
  if (qi) {
    qi.focus();
    qi.addEventListener('input', () => {
      const t = qi.value.trim().toLowerCase();
      pop.querySelectorAll('.hf-list label').forEach((l) => {
        l.hidden = !!t && l.dataset.v.indexOf(t) < 0; });
    });
  }
  pop.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); pop.querySelector('[data-hf="apply"]').click(); } });

  pop.addEventListener('click', (e) => {
    const b = e.target.closest('[data-hf]');
    if (!b || b.tagName === 'INPUT') return;
    const a = b.dataset.hf;
    if (a === 'x') return closeHf();
    if (a === 'asc' || a === 'desc') { S.sort[scope] = { key, dir: a === 'asc' ? 1 : -1 }; return done(); }
    if (a === 'all' || a === 'none') {
      pop.querySelectorAll('.hf-list label').forEach((l) => {
        if (!l.hidden) l.querySelector('input').checked = (a === 'all'); });
      return;
    }
    const F = hfOf(scope);
    if (a === 'reset') { delete F[key]; return done(); }
    if (a !== 'apply') return;
    if (c.num) {
      const on = {}; let nb = 0, tb = 0;
      pop.querySelectorAll('.hf-bd').forEach((cb) => { tb++; if (cb.checked) { on[cb.value] = 1; nb++; } });
      const mn = pop.querySelector('[data-hf="min"]').value.replace(/[^0-9.\-]/g, '');
      const mx = pop.querySelector('[data-hf="max"]').value.replace(/[^0-9.\-]/g, '');
      const f = { type: 'num', bands: nb === tb ? null : on,
        min: mn === '' ? null : parseFloat(mn), max: mx === '' ? null : parseFloat(mx) };
      if (!f.bands && f.min === null && f.max === null) delete F[key]; else F[key] = f;
    } else {
      /* only what is visible AND ticked: a search narrows the choice, as in Excel */
      const labels = Array.from(pop.querySelectorAll('.hf-list label'));
      const on = labels.filter((l) => !l.hidden && l.querySelector('input').checked)
        .map((l) => l.querySelector('input').value);
      if (on.length === labels.length && !labels.some((l) => l.hidden)) delete F[key];
      else F[key] = { type: 'set', values: new Set(on) };
    }
    done();
  });
}

/* ══ 8 · rendering primitives ════════════════════════════════════════════ */

const chip = (txt, cls) => `<span class="fk-chip ${cls}">${esc(txt)}</span>`;
const docLink = (dt, dn) => dn
  ? `<a class="fk-lnk" href="/app/${String(dt).toLowerCase().replace(/ /g, '-')}/${q(dn)}" target="_blank" rel="noopener">${esc(dn)} ↗</a>`
  : '—';
const entryLink = (r) => docLink(r.entry_type, r.entry_id);
const supLink = (r) => r.supplier
  ? `<a class="fk-lnk" href="/app/supplier/${q(r.supplier)}" target="_blank" rel="noopener">${esc(r.supplier_name || r.supplier)} ↗</a>`
  : '—';
/* every amount that represents a party position opens that party's ledger */
const glLink = (sup, text) => sup
  ? `<button type="button" class="fk-numbtn" data-apl="${esc(sup)}">${esc(text)}</button>`
  : esc(text || '');
const erpNum = (v, href, tip) => Math.abs(num(v)) < 0.005
  ? '<span class="zero">–</span>'
  : `<a class="fk-numbtn" href="${href}" target="_blank" rel="noopener" title="${esc(tip)}">${esc(inr(v))}</a>`;

/* KPI cards are buttons — every one opens the exact rows behind its number.
   The chip keeps only its first clause on screen and carries the rest as a
   title, so a card is one line high whatever it has to say. */
let KPI_SEQ = 0;
const KPI_REG = new Map();
function kpi(host, items) {
  const h = el(host); if (!h) return;
  h.innerHTML = items.map((i) => {
    const parts = String(i.c || '').split(' · ').filter((x) => /\d/.test(x.replace(/<[^>]+>/g, '')));
    const id = 'k' + (++KPI_SEQ);
    KPI_REG.set(id, i);
    const tag = i.drill || i.go ? 'button' : 'div';
    return `<${tag} class="fk-kpi" data-kpi="${id}" style="--tone:${i.tone || 'var(--s1)'}"` +
      (tag === 'button' ? ` type="button" aria-label="${esc(i.k)} — open detail"` : '') + '>' +
      `<div class="k"><span>${esc(i.k)}</span></div>` +
      `<div class="v${num(i.raw) <= -0.005 ? ' neg' : ''}">${i.v}</div>` +
      `<div class="c" title="${esc(parts.join(' · '))}">${esc(parts[0] || '')}</div></${tag}>`;
  }).join('');
}

const emptySvg = (m) =>
  `<text class="tick" x="50%" y="50%" text-anchor="middle" style="font-size:12px">${esc(m)}</text>`;

function barChart(target, rows, opt) {
  opt = opt || {};
  const svg = el(target); if (!svg) return;
  if (!rows.length) { svg.innerHTML = emptySvg(opt.empty || 'Nothing to show'); return; }
  const W = opt.W || 620, rowH = 32, p = { l: opt.l || 170, r: 190, t: 14, b: 10 };
  const H = p.t + p.b + rows.length * rowH;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  const max = Math.max.apply(null, rows.map((r) => Math.abs(r.v))) || 1;
  const bw = W - p.l - p.r;
  svg.innerHTML = rows.map((r, i) => {
    const y = p.t + i * rowH, w = Math.max(2, Math.abs(r.v) / max * bw);
    return `<g class="${r.key != null ? 'hit' : ''}"${r.key != null ? ` data-bar="${esc(String(r.key))}"` : ''}>`
      + `<text class="tick" text-anchor="end" x="${p.l - 10}" y="${y + 16}" style="fill:var(--ink);font-weight:600">${esc(r.label)}</text>`
      + `<rect x="${p.l}" y="${y + 5}" width="${w}" height="16" rx="3" fill="${r.colour || 'var(--s1)'}">`
      + `<title>${esc(r.label)}: ${esc(rs(r.v))}</title></rect>`
      + `<text class="dlabel" x="${p.l + w + 8}" y="${y + 17}">${esc(rs(r.v))}</text>`
      + (r.sub ? `<text class="dsub" text-anchor="end" x="${p.l - 10}" y="${y + 27}">${esc(r.sub)}</text>` : '')
      + '</g>';
  }).join('');
}

function table(target, cols, rows, opt) {
  opt = opt || {};
  const t = el(target); if (!t) return;
  const F = opt.hf ? (S.hf[opt.hf] || {}) : null;
  const sortKey = opt.hf && S.sort[opt.hf] ? S.sort[opt.hf] : null;
  const head = '<thead><tr>' + cols.map((c) =>
    `<th class="${c.num ? 'num ' : ''}${c.div ? 'div ' : ''}${F && F[c.k] ? 'hf-on' : ''}" data-k="${esc(c.k)}">` +
    esc(c.label) +
    (sortKey && sortKey.key === c.k ? `<span class="ar">${sortKey.dir > 0 ? '▲' : '▼'}</span>` : '') +
    (F && !c.nofilter ? `<button type="button" class="hf-btn" data-hfk="${esc(c.k)}" aria-label="Filter and sort ${esc(c.label)}">${F[c.k] ? '⏷' : '▾'}</button>` : '') +
    '</th>').join('') + '</tr></thead>';

  let body = '';
  if (!rows.length) {
    body = `<tbody><tr><td class="empty" colspan="${cols.length}">${esc(opt.empty || 'Nothing matches the current scope.')}</td></tr></tbody>`;
  } else {
    body = '<tbody>' + rows.map((r) => {
      let pre = '';
      if (opt.section) { const s = opt.section(r);
        if (s) pre = `<tr class="sect"><td colspan="${cols.length}">${esc(s)}</td></tr>`; }
      return pre + `<tr${r.__key ? ` data-row="${esc(r.__key)}"` : ''}>` + cols.map((c) => {
        const v = c.cell ? c.cell(r) : esc(r[c.k]);
        const raw = c.num ? num(hfNum(c, r)) : 0;
        return `<td class="${c.num ? 'num' : 'txt'}${c.div ? ' div' : ''}${c.num && raw < -0.005 ? ' neg' : ''}">${v}</td>`;
      }).join('') + '</tr>';
    }).join('') + '</tbody>';
  }

  let foot = '';
  if (opt.foot) {
    foot += '<tr>' + cols.map((c) =>
      `<td class="${c.num ? 'num' : 'txt'}${c.div ? ' div' : ''}">${opt.foot[c.k] != null ? opt.foot[c.k] : ''}</td>`).join('') + '</tr>';
  }
  (opt.recon || []).forEach((line) => {
    foot += `<tr><td class="recon" colspan="${cols.length}">${line}</td></tr>`;
  });
  t.innerHTML = head + body + (foot ? '<tfoot>' + foot + '</tfoot>' : '');
}

function footFor(cols, rows, label) {
  const f = {};
  cols.forEach((c) => { if (c.num && c.raw) f[c.k] = rs(rows.reduce((a, r) => a + num(c.raw(r)), 0)); });
  if (cols[0]) f[cols[0].k] = label || ('Total — ' + countOf(rows.length, 'row'));
  return f;
}

/* one reconciliation line, rendered the same way everywhere it appears */
function reconLine(label, got, expect, note) {
  if (expect === null || expect === undefined) {
    return `<span>${label} · <b>${esc(rs(got))}</b>${note ? ' · ' + note : ''}</span>`;
  }
  const d = r2(num(got) - num(expect)), ok = Math.abs(d) <= TOL;
  return `<span class="${ok ? 'good' : 'bad'}">${ok ? '✔' : '✖'}</span> ${label} · ` +
    `<b>${esc(rs(got))}</b> vs <b>${esc(rs(expect))}</b> · Δ <b>${esc(rs(d))}</b>` +
    (note ? ' · ' + note : '');
}

function pager(target, total, page, go) {
  const host = el(target); if (!host) return;
  const pages = Math.max(1, Math.ceil(total / S.PAGE));
  if (page > pages) page = pages;
  host.innerHTML = `<span>${cnt(total)} rows · page ${page} of ${pages}</span>`
    + `<button type="button" data-go="${page - 1}" ${page <= 1 ? 'disabled' : ''}>‹ Prev</button>`
    + `<button type="button" data-go="${page + 1}" ${page >= pages ? 'disabled' : ''}>Next ›</button>`;
  host.onclick = (e) => {
    const b = e.target.closest('[data-go]'); if (!b || b.disabled) return;
    const n = parseInt(b.dataset.go, 10); if (n >= 1 && n <= pages) go(n);
  };
}
const slice = (rows, page) => rows.slice((page - 1) * S.PAGE, (page - 1) * S.PAGE + S.PAGE);
const matches = (q0, ...fields) => !q0 || fields.map((f) => String(f == null ? '' : f))
  .join(' ').toLowerCase().indexOf(q0) >= 0;

/* ══ 9 · column definitions ══════════════════════════════════════════════ */

const money = (k, label, div) => ({ k, label, num: true, div: !!div,
  cell: (r) => esc(inr(r[k])), raw: (r) => r[k] });
const moneyDash = (k, label, div) => ({ k, label, num: true, div: !!div,
  cell: (r) => Math.abs(num(r[k])) < 0.005 ? '<span class="zero">' + dash(r[k]) + '</span>' : esc(inr(r[k])),
  raw: (r) => r[k] });
const dateCol = (k, label) => ({ k, label, cell: (r) => esc(ddmmmyyyy(r[k])), plain: (r) => r[k] || '' });

function deliveryChip(r) {
  const req = r.required_by;
  if (!req) return chip('◆ No required-by date', 'c-neu');
  const late = daysAgo(req, S.asOf);
  if (num(r.pending_receipt) <= 0.005) return chip('✔ Delivered', 'c-ok');
  if (late <= 0) return chip('◆ Not yet due', 'c-neu');
  if (late <= 7) return chip('▲ ' + cnt(late) + ' days late', 'c-warn');
  return chip('✖ ' + cnt(late) + ' days late', 'c-crit');
}

const C = {
  date:  dateCol('date', 'Date'),
  post:  dateCol('date', 'Posting date'),
  req:   dateCol('required_by', 'Required by'),
  due:   dateCol('required_by', 'Due date'),
  sup:   { k: 'supplier_name', label: 'Vendor', cell: supLink,
           plain: (r) => r.supplier_name || r.supplier || '(blank)' },
  status:{ k: 'status', label: 'Status' },
  src:   { k: 'source_class', label: 'Route' },
  entry: { k: 'entry_id', label: 'Entry ID', cell: entryLink },
  etype: { k: 'entry_type', label: 'Entry type' },
  flag:  { k: 'quality_flag', label: 'Quality flag',
           cell: (r) => r.quality_flag ? chip('▲ ' + r.quality_flag, 'c-warn') : '' },
  loc:   { k: 'ap_location', label: 'AP Location',
           cell: (r) => r.ap_location === 'Mixed' ? chip('◆ Mixed', 'c-warn') : locChip(r.ap_location),
           plain: (r) => r.ap_location || 'Unassigned' },
  state: { k: 'registration_state', label: 'Registration state',
           cell: (r) => esc(r.registration_state || '—') },
  gstin: { k: 'resolved_gstin', label: 'Resolved GSTIN', cell: (r) => esc(r.resolved_gstin || '—') },
  cstat: { k: 'classification_status', label: 'Classification',
           cell: (r) => esc(r.classification_status || '—') },
  csrc:  { k: 'classification_source', label: 'Classification source',
           cell: (r) => esc(r.classification_source || '—') },
  creason:{ k: 'classification_reason', label: 'Classification evidence',
           cell: (r) => `<small>${esc(r.classification_reason || '')}</small>` },
  by:    { k: 'created_by', label: 'Created by', cell: (r) => esc(r.created_by || '—') },
  bank:  { k: 'bank', label: 'Bank account', cell: (r) => esc(r.bank || '—') },
  mode:  { k: 'payment_mode', label: 'Mode', cell: (r) => esc(r.payment_mode || '—') },
  reconState: { k: 'recon_state', label: 'Bank reconciliation',
           cell: (r) => chip(r.recon_state === 'Bank Reconciled' ? '✔ Reconciled' : '▲ Unreconciled',
                             r.recon_state === 'Bank Reconciled' ? 'c-ok' : 'c-warn'),
           plain: (r) => r.recon_state || '(blank)' }
};

const COLS = {};
const pick = (cols, keys) => keys.map((k) => cols.find((c) => c.k === k)).filter(Boolean);

COLS.po = [C.date, { k: 'grain_key', label: 'PO ID', cell: (r) => docLink('Purchase Order', r.grain_key) },
  C.sup, C.req, C.status, money('taxable', 'Taxable'), money('gross', 'Total incl GST'),
  money('received', 'Received'), money('pending_receipt', 'Yet to receive'), money('invoiced', 'Invoiced')];

COLS.poFull = [
  { k: 'grain_key', label: 'PO ID',
    cell: (r) => `<button type="button" class="fk-exp" data-po="${esc(r.grain_key)}" aria-label="Expand to items">▸</button>` + docLink('Purchase Order', r.grain_key) },
  C.date, C.sup, C.req, C.status,
  money('taxable', 'Taxable'), money('cgst', 'CGST'), money('sgst', 'SGST'), money('igst', 'IGST'),
  money('gross', 'Total incl GST'), money('advance', 'Advance'),
  money('received', 'Received'), money('pending_receipt', 'Yet to receive'), money('invoiced', 'Invoiced'),
  { k: 'age_days', label: 'Days since PO', num: true,
    cell: (r) => cnt(daysAgo(r.date, S.asOf)), raw: (r) => daysAgo(r.date, S.asOf) },
  { k: 'delivery', label: 'Delivery status', cell: deliveryChip,
    plain: (r) => num(r.pending_receipt) <= 0.005 ? 'Delivered' : 'Pending' },
  C.loc, C.state, C.cstat, C.by, C.etype, C.entry];

COLS.pr = [C.date, { k: 'grain_key', label: 'PR / GRN ID', cell: (r) => docLink('Purchase Receipt', r.grain_key) },
  C.sup, money('gross', 'Receipt value'), money('invoiced', 'Billed'), money('pending_bill', 'Unbilled')];

COLS.prFull = [
  { k: 'grain_key', label: 'PR / GRN ID',
    cell: (r) => `<button type="button" class="fk-exp" data-pr="${esc(r.grain_key)}" aria-label="Expand to items">▸</button>` + docLink('Purchase Receipt', r.grain_key) },
  dateCol('date', 'Incoming date'), C.sup,
  /* the vendor invoice number is resolved from the invoice that carries the
     receipt, so a receipt that has not been billed has none and says so */
  { k: 'supplier_invoice_no', label: 'Vendor invoice no',
    cell: (r) => r.supplier_invoice_no ? esc(r.supplier_invoice_no)
      : num(r.invoiced) > 0.005 ? chip('▲ Vendor invoice no. missing', 'c-warn')
      : chip('◆ Not billed yet', 'c-neu'),
    plain: (r) => r.supplier_invoice_no || '(blank)' },
  { k: 'warehouse', label: 'Warehouse / unit', cell: (r) => esc(r.warehouse || '—') },
  C.loc, C.state,
  money('taxable', 'Taxable'), money('cgst', 'CGST'), money('sgst', 'SGST'), money('igst', 'IGST'),
  money('gross', 'Total incl GST'), money('invoiced', 'Billed'), money('pending_bill', 'Unbilled'),
  { k: 'grn2inv', label: 'GRN → invoice days', num: true,
    raw: (r) => r.supplier_invoice_date ? daysAgo(r.date, r.supplier_invoice_date) : null,
    cell: (r) => { if (!r.supplier_invoice_date) return '—';
      const n = daysAgo(r.date, r.supplier_invoice_date);
      return n < 0 ? chip('▲ Invoice precedes GRN by ' + cnt(Math.abs(n)) + ' day' + (Math.abs(n) === 1 ? '' : 's'), 'c-warn') : cnt(n); } },
  { k: 'recon_state', label: 'State',
    cell: (r) => chip(r.recon_state === 'Fully Billed' ? '✔ Fully billed' : '▲ Unbilled',
                      r.recon_state === 'Fully Billed' ? 'c-ok' : 'c-warn'),
    plain: (r) => r.recon_state || '(blank)' },
  C.src, C.etype, C.entry];

COLS.grnVendor = [
  { k: 'name', label: 'Vendor',
    cell: (r) => `<a class="fk-lnk" href="/app/supplier/${q(r.supplier)}" target="_blank" rel="noopener">${esc(r.name)} ↗</a>`,
    plain: (r) => r.name },
  money('po', 'PO value'), money('grn', 'GRN done'), money('pending', 'Pending receipt'),
  money('billed', 'Billed'), money('unbilled', 'Unbilled'), money('ret', 'Return value'),
  dateCol('earliest', 'Earliest pending required'),
  { k: 'oldest', label: 'Oldest unbilled (days)', num: true,
    cell: (r) => r.oldest ? cnt(r.oldest) : '—', raw: (r) => r.oldest }];

COLS.piFull = [
  { k: 'grain_key', label: 'PI ID',
    cell: (r) => `<button type="button" class="fk-exp" data-pi="${esc(r.grain_key)}" aria-label="Expand to items and tax">▸</button>` + docLink('Purchase Invoice', r.grain_key) },
  C.loc, C.state, C.gstin, C.cstat, C.csrc, C.post, C.sup,
  { k: 'supplier_gstin', label: 'Supplier GSTIN', cell: (r) => esc(r.supplier_gstin || '—') },
  { k: 'supplier_invoice_no', label: 'Supplier bill no',
    cell: (r) => r.supplier_invoice_no ? esc(r.supplier_invoice_no) : chip('✖ Missing', 'c-crit'),
    plain: (r) => r.supplier_invoice_no || '(blank)' },
  dateCol('supplier_invoice_date', 'Bill date'),
  { k: 'doc_kind', label: 'Document',
    cell: (r) => IS_RETURN(r) ? chip('↩ Debit note / return', 'c-warn') : 'Invoice',
    plain: (r) => IS_RETURN(r) ? 'Debit note / return' : 'Invoice' },
  C.src,
  { k: 'place_of_supply', label: 'Place of supply', cell: (r) => esc(r.place_of_supply || '—') },
  { k: 'reverse_charge', label: 'Reverse charge',
    cell: (r) => num(r.reverse_charge) ? 'Yes' : 'No', plain: (r) => num(r.reverse_charge) ? 'Yes' : 'No' },
  money('taxable', 'Taxable'), money('cgst', 'CGST'), money('sgst', 'SGST'), money('igst', 'IGST'),
  money('tds', 'TDS'), money('other_charges', 'Other'), money('round_off', 'Round off'),
  money('gross', 'Net payable'), money('paid', 'Paid'), money('outstanding', 'Outstanding'),
  C.due,
  { k: 'age_days', label: 'Days overdue', num: true,
    cell: (r) => num(r.age_days) > 0 ? cnt(r.age_days) : '—', raw: (r) => r.age_days },
  { k: 'po_id', label: 'PO', cell: (r) => esc(r.po_id || '—') },
  { k: 'pr_id', label: 'PR', cell: (r) => esc(r.pr_id || '—') },
  { k: 'company_address', label: 'Company address', cell: (r) => esc(r.company_address || '—') },
  { k: 'leaf_cost_centre', label: 'Leaf cost centre',
    cell: (r) => esc(r.leaf_cost_centre || r.cost_center || '—') },
  { k: 'resolved_location_cost_centre', label: 'Resolved location cost centre',
    cell: (r) => esc(r.resolved_location_cost_centre || '—') },
  { k: 'department', label: 'Department', cell: (r) => esc(r.department || '—') },
  { k: 'division', label: 'Division', cell: (r) => esc(r.division || '—') },
  { k: 'project', label: 'Project', cell: (r) => esc(r.project || '—') },
  C.by, C.creason, C.flag, C.etype, C.entry];
/* on screen: the columns people read. The CSV carries every column above. */
COLS.piView = pick(COLS.piFull, ['grain_key', 'ap_location', 'date', 'supplier_name',
  'supplier_invoice_no', 'source_class', 'taxable', 'cgst', 'sgst', 'igst', 'tds',
  'gross', 'paid', 'outstanding', 'age_days', 'classification_status']);

COLS.tax = [C.date, { k: 'pi_id', label: 'PI ID', cell: (r) => docLink('Purchase Invoice', r.pi_id) },
  C.sup, { k: 'item_name', label: 'Account head', cell: (r) => esc(r.item_name || '—') },
  money('cgst', 'CGST'), money('sgst', 'SGST'), money('igst', 'IGST'),
  money('tds', 'TDS'), money('other_charges', 'Other')];

COLS.pay = [C.date, { k: 'pe_id', label: 'Payment Entry', cell: (r) => docLink('Payment Entry', r.pe_id) },
  C.sup, money('paid', 'Allocated'), money('advance', 'Unallocated'),
  { k: 'supplier_invoice_no', label: 'Reference type', cell: (r) => esc(r.supplier_invoice_no || '—') },
  { k: 'item_name', label: 'Reference ID', cell: (r) => esc(r.item_name || '—') },
  C.bank, C.reconState];
COLS.payFull = COLS.pay.concat([C.mode,
  { k: 'reference_no', label: 'Reference / UTR', cell: (r) => esc(r.reference_no || '—') },
  C.status, C.loc, C.state, C.gstin, C.cstat,
  { k: 'cost_center', label: 'Cost centre', cell: (r) => esc(r.cost_center || '—') },
  { k: 'department', label: 'Department', cell: (r) => esc(r.department || '—') },
  C.by, C.entry]);
COLS.payAdv = [C.date, { k: 'pe_id', label: 'Payment Entry', cell: (r) => docLink('Payment Entry', r.pe_id) },
  C.sup, money('paid', 'Paid'), money('advance', 'Unallocated'), C.bank, C.mode, C.reconState];

COLS.payReg = [
  { k: 'pe_id', label: 'Payment Entry', cell: (r) => docLink('Payment Entry', r.pe_id) },
  C.post, C.sup, C.bank, C.mode,
  { k: 'reference_no', label: 'Reference / UTR', cell: (r) => esc(r.reference_no || '—') },
  dateCol('reference_date', 'Reference date'),
  money('paid', 'Paid'), money('allocated', 'Allocated'), money('unallocated', 'Unallocated'),
  { k: 'linked_n', label: 'Linked invoices', num: true, cell: (r) => cnt(r.linked_n), raw: (r) => r.linked_n },
  { k: 'linked_ids', label: 'Linked invoice IDs', cell: (r) => `<small>${esc(r.linked_ids || '—')}</small>` },
  C.loc,
  { k: 'split_text', label: 'Split by allocation', cell: (r) => `<small>${esc(r.split_text || '—')}</small>` },
  C.state, C.gstin, C.reconState, dateCol('clearance_date', 'Clearance date'), C.by, C.flag];
COLS.payView = pick(COLS.payReg, ['pe_id', 'date', 'supplier_name', 'bank', 'payment_mode',
  'reference_no', 'paid', 'allocated', 'unallocated', 'linked_n', 'ap_location',
  'recon_state', 'quality_flag']);

COLS.unapplied = [C.sup,
  { k: 'pe_id', label: 'Payment Entry', cell: (r) => docLink('Payment Entry', r.pe_id) },
  C.loc, money('advance', 'Unapplied cash'),
  { k: 'quality_flag', label: 'Basis', cell: (r) => esc(r.quality_flag || '') }];

COLS.openItems = [dateCol('supplier_invoice_date', 'Bill date'), C.post,
  { k: 'grain_key', label: 'PI ID', cell: (r) => docLink('Purchase Invoice', r.grain_key) },
  C.sup,
  { k: 'supplier_invoice_no', label: 'Supplier bill no', cell: (r) => esc(r.supplier_invoice_no || '—') },
  C.due, money('gross', 'Net payable'),
  { k: 'outstanding', label: 'Outstanding', num: true, raw: (r) => r.outstanding,
    cell: (r) => glLink(r.supplier, inr(r.outstanding)) },
  { k: 'age_days', label: 'Days outstanding', num: true, cell: (r) => cnt(r.age_days), raw: (r) => r.age_days },
  C.loc, C.state, C.cstat, C.src, C.entry];

COLS.cohort = [
  { k: 'k', label: 'Cohort', cell: (r) => esc(bucketLabel(r.k)), plain: (r) => bucketLabel(r.k) },
  { k: 'n', label: 'Orders', num: true, cell: (r) => r.n ? cnt(r.n) : '—', raw: (r) => r.n },
  money('ordered', 'Ordered'), money('received', 'Received'),
  money('pendingReceipt', 'Pending receipt'), money('invoiced', 'Invoiced'),
  money('pendingBill', 'Pending invoice'), money('paid', 'Paid')];

COLS.truePurchase = [
  { k: 'name', label: 'Vendor',
    cell: (r) => `<a class="fk-lnk" href="/app/supplier/${q(r.supplier)}" target="_blank" rel="noopener">${esc(r.name)} ↗</a>`,
    plain: (r) => r.name },
  money('p2p', 'Inventory P2P'), money('directExp', 'Direct expense / service'),
  money('debitNotes', 'of which debit notes'),
  money('net', 'Net booked'), money('paid', 'Paid'), money('outstanding', 'Outstanding'),
  money('gst', 'Input GST'), money('tds', 'TDS')];

COLS.journal = [C.date, C.sup, { k: 'entry_id', label: 'Voucher', cell: entryLink },
  { k: 'source_class', label: 'Kind' }, money('outstanding', 'Credit − debit'),
  C.loc, C.cstat, C.creason];
COLS.control = [C.sup, { k: 'entry_id', label: 'Document', cell: entryLink },
  { k: 'status', label: 'Part' }, money('outstanding', 'Credit − debit'), C.loc, C.cstat];

COLS.itemPO = [{ k: 'item', label: 'Item', cell: (r) => esc(r.item || '') },
  { k: 'item_name', label: 'Description', cell: (r) => esc(r.item_name || '') },
  { k: 'qty', label: 'Ordered qty', num: true, cell: (r) => inr(r.qty, 3), raw: (r) => r.qty },
  money('ordered', 'Ordered ₹'), money('received', 'Received ₹'),
  money('pending_receipt', 'Pending ₹'), money('invoiced', 'Invoiced ₹'),
  C.req, { k: 'warehouse', label: 'Warehouse', cell: (r) => esc(r.warehouse || '—') }, C.flag];
COLS.itemPR = [{ k: 'item', label: 'Item', cell: (r) => esc(r.item || '') },
  { k: 'item_name', label: 'Description', cell: (r) => esc(r.item_name || '') },
  { k: 'qty', label: 'Qty', num: true, cell: (r) => inr(r.qty, 3), raw: (r) => r.qty },
  money('received', 'Received ₹'), money('invoiced', 'Billed ₹'), money('pending_bill', 'Unbilled ₹'),
  { k: 'po_id', label: 'PO', cell: (r) => docLink('Purchase Order', r.po_id) },
  { k: 'warehouse', label: 'Warehouse', cell: (r) => esc(r.warehouse || '—') }, C.flag];
COLS.itemPI = [{ k: 'item', label: 'Item / account head', cell: (r) => esc(r.item || r.item_name || '') },
  { k: 'item_name', label: 'Description', cell: (r) => esc(r.item_name || '') },
  { k: 'qty', label: 'Qty', num: true, cell: (r) => num(r.qty) ? inr(r.qty, 3) : '—', raw: (r) => r.qty },
  money('invoiced', 'Net ₹'), money('cgst', 'CGST'), money('sgst', 'SGST'),
  money('igst', 'IGST'), money('tds', 'TDS'), money('other_charges', 'Other'),
  { k: 'po_id', label: 'PO', cell: (r) => docLink('Purchase Order', r.po_id) },
  { k: 'pr_id', label: 'PR', cell: (r) => docLink('Purchase Receipt', r.pr_id) }, C.src];

/* ── the merged supplier table ─────────────────────────────────────────────
   One row per enabled supplier. The identity it proves, left to right:
     bills + debit notes − advance + payments on account + journals = party ledger
   Delta is that identity's residue, so a non-zero Delta means the AP Aging
   grain and the AP GL Control grain disagree for that supplier — which is the
   exception worth seeing, not a rounding artefact of an incomplete formula. */
const vendorCell = {
  k: 'nm', label: 'Vendor',
  cell: (r) => `<button type="button" class="fk-exp" data-agsup="${esc(r.id)}" aria-label="Expand to bills">▸</button>`
    + `<a class="fk-lnk" href="/app/supplier/${q(r.id)}" target="_blank" rel="noopener" title="${esc(r.id)}">${esc(r.nm)} ↗</a>`,
  plain: (r) => r.nm
};
COLS.agSupplier = [vendorCell,
  { k: 'cls', label: 'Class', cell: (r) => chip(CLASSES[r.cls].label, CLASSES[r.cls].tone),
    plain: (r) => CLASSES[r.cls].label }]
  .concat(BANDS.map((b, i) => ({ k: b.k, label: b.label, num: true,
    raw: (r) => r.b[i],
    cell: (r) => erpNum(r.b[i], apUrl(r.id), 'Open in the native Accounts Payable report') })))
  .concat([
    { k: 'billTot', label: 'Bills total', num: true, raw: (r) => r.billTot,
      cell: (r) => erpNum(r.billTot, apUrl(r.id), 'Open in the native Accounts Payable report') },
    moneyDash('dn', 'Debit notes'),
    moneyDash('adv', 'Advance held', true),
    moneyDash('allocatable', 'Allocatable'),
    moneyDash('pp', 'Payments on account'),
    moneyDash('jv', 'Open journals'),
    /* the proof block: what this page says, then the two ERP numbers it has to
       agree with, each with its own distance. A dash means agreement within
       the tolerance; a figure means a real disagreement worth opening. */
    { k: 'net', label: 'This page', num: true, div: true, raw: (r) => r.net,
      cell: (r) => glLink(r.id, inr(r.net)) },
    { k: 'nat', label: 'Native AP report', num: true, raw: (r) => r.nat,
      cell: (r) => r.nat === null ? '<span class="zero">n/a</span>'
        : erpNum(r.nat, apUrl(r.id), "Open this vendor in ERPNext's own Accounts Payable report") },
    { k: 'dNat', label: 'Δ vs native AP', num: true, raw: (r) => r.dNat,
      cell: (r) => r.dNat === null ? '<span class="zero">n/a</span>'
        : Math.abs(r.dNat) <= TOL ? '<span class="zero">–</span>'
        : `<a class="fk-numbtn" href="${apUrl(r.id)}" target="_blank" rel="noopener">${esc(inr(r.dNat))}</a>` },
    { k: 'vl', label: 'Vendor ledger', num: true, div: true, raw: (r) => r.vl,
      cell: (r) => r.vl === null ? '<span class="zero">n/a</span>'
        : glLink(r.id, inr(r.vl)) },
    { k: 'dVl', label: 'Δ vs ledger', num: true, raw: (r) => r.dVl,
      cell: (r) => r.dVl === null ? '<span class="zero">n/a</span>'
        : Math.abs(r.dVl) <= TOL ? '<span class="zero">–</span>'
        : `<button type="button" class="fk-numbtn" data-apl="${esc(r.id)}">${esc(inr(r.dVl))}</button>` },
    { k: 'action', label: 'Action', nofilter: false, plain: (r) => r.action || '(none)',
      cell: (r) => r.action
        ? `<button type="button" class="fk-act" data-agact="${r.action}" data-agid="${esc(r.id)}">${esc(r.action)}</button>` : '' }
  ]);
/* the CSV carries the roster identifiers and the ERP links too */
COLS.agSupplierFull = COLS.agSupplier.concat([
  { k: 'id', label: 'Supplier ID', plain: (r) => r.id },
  { k: 'led', label: 'Master report party balance', num: true, raw: (r) => r.led },
  { k: 'delta', label: 'Δ master report vs this page', num: true, raw: (r) => r.delta },
  { k: 'jrGl', label: 'Journals not linked to an invoice', num: true, raw: (r) => r.jrGl },
  { k: 'erp_gl', label: 'ERP General Ledger', plain: (r) => location.origin + glUrl(r.id) },
  { k: 'erp_ap', label: 'ERP Accounts Payable', plain: (r) => location.origin + apUrl(r.id) }
]);

COLS.agBill = [
  { k: 'supName', label: 'Vendor',
    cell: (r) => `<a class="fk-lnk" href="/app/supplier/${q(r.sup)}" target="_blank" rel="noopener">${esc(r.supName)} ↗</a>`,
    plain: (r) => r.supName },
  { k: 'no', label: 'Bill no', cell: (r) => esc(r.no || r.vno), plain: (r) => r.no || r.vno },
  { k: 'vno', label: 'Voucher', cell: (r) => docLink('Purchase Invoice', r.vno) },
  dateCol('from', 'Aged from'), dateCol('due', 'Due date'),
  { k: 'age', label: 'Age', num: true, cell: (r) => cnt(r.age), raw: (r) => r.age },
  { k: 'band', label: 'Bucket',
    cell: (r) => chip(r.band, r.band === 'Debit note' ? 'c-neu' : r.band === 'Not due' ? 'c-ok'
      : (r.band === 'Over 180' || r.band === '91 – 180') ? 'c-crit' : 'c-warn'),
    plain: (r) => r.band },
  { k: 'amt', label: 'Balance', num: true, raw: (r) => r.amt,
    cell: (r) => glLink(r.sup, inr(r.amt)) },
  { k: 'loc', label: 'AP Location', cell: (r) => locChip(r.loc), plain: (r) => r.loc }];

/* ── vendor balances ──────────────────────────────────────────────────────── */
const vbVendor = { k: 'name', label: 'Vendor',
  cell: (r) => `<button type="button" class="fk-exp" data-vbsup="${esc(r.supplier)}" aria-label="Expand to open invoices">▸</button>`
    + `<a class="fk-lnk" href="/app/supplier/${q(r.supplier)}" target="_blank" rel="noopener">${esc(r.name)} ↗</a>`,
  plain: (r) => r.name };
COLS.vb = [vbVendor, money('opening', 'Opening'), money('billed', 'Billed'),
  money('dn', 'Debit notes'), money('jr', 'Journals'), money('pay', 'Payments'), money('other', 'Other'),
  { k: 'closing', label: 'Closing', num: true, div: true, raw: (r) => r.closing,
    cell: (r) => glLink(r.supplier, inr(r.closing)) },
  { k: 'variance', label: 'Tie', num: true, raw: (r) => r.variance,
    cell: (r) => chip((r.ties ? '✔ ' : '✖ ') + inr(r.variance), r.ties ? 'c-ok' : 'c-crit'),
    plain: (r) => r.ties ? 'Ties' : 'Off' }];
COLS.vbOpening = [vbVendor, money('openJE', 'Migration JV'), money('openPI', 'Pre-period bills'),
  money('openPE', 'Pre-period payments'), money('openOther', 'Other'),
  money('opening', 'Opening total'), money('closing', 'Closing')];
COLS.vbFull = COLS.vb.concat([money('openJE', 'Migration JV'), money('openPI', 'Pre-period bills'),
  money('openPE', 'Pre-period payments'), money('openOther', 'Opening other'),
  money('unap', 'Unapplied cash'), money('movement', 'Period movement'),
  { k: 'n', label: 'GL rows', num: true, cell: (r) => cnt(r.n), raw: (r) => r.n }]);

/* ── vendor credits · one table per door ──────────────────────────────────── */
const door = (cls, txt, label) => ({ k: 'door', label: 'Door', nofilter: true,
  cell: () => `<span class="fk-door ${cls}">${txt}</span>`, plain: () => label });
COLS.vcDn = [door('dn', 'DN', 'Purchase return'),
  { k: 'grain_key', label: 'Document', cell: (r) => docLink('Purchase Invoice', r.grain_key) },
  C.sup, C.date, money('vc_raised', 'Credit raised'), money('vc_applied', 'Applied'),
  { k: 'vc_avail', label: 'Available', num: true, raw: (r) => r.vc_avail,
    cell: (r) => `<b>${esc(inr(r.vc_avail))}</b>` },
  { k: 'vc_age', label: 'Age', num: true, cell: (r) => cnt(r.vc_age), raw: (r) => r.vc_age },
  { k: 'vc_state', label: 'Flag', nofilter: true,
    cell: (r) => r.vc_round ? chip('▲ paise vs rupee', 'c-crit')
      : r.vc_age > 90 ? chip('✖ ageing', 'c-crit') : chip('▲ open', 'c-warn'),
    plain: (r) => (r.vc_age > 90 ? 'ageing' : 'open') + (r.vc_round ? ' | paise vs rupee' : '') },
  C.loc];
COLS.vcPe = [door('pe', 'PE', 'Unapplied payment'),
  { k: 'pe_id', label: 'Payment Entry', cell: (r) => docLink('Payment Entry', r.pe_id) },
  C.sup, C.date, money('paid', 'Paid'),
  { k: 'advance', label: 'Unapplied', num: true, raw: (r) => r.advance,
    cell: (r) => `<b>${esc(inr(r.advance))}</b>` },
  { k: 'vc_age', label: 'Age', num: true, cell: (r) => cnt(r.vc_age), raw: (r) => r.vc_age },
  C.bank,
  { k: 'vc_hint', label: 'Oldest open bill', nofilter: true,
    cell: (r) => r.vc_hint ? docLink('Purchase Invoice', r.vc_hint.id) : '<span class="zero">–</span>',
    plain: (r) => r.vc_hint ? r.vc_hint.id : '' },
  C.loc];
COLS.vcJe = [door('je', 'JE', 'Journal credit'),
  { k: 'entry_id', label: 'Voucher', cell: entryLink },
  C.sup, C.date, money('outstanding', 'Credit'), C.loc];

/* ── Categorise AP Location ───────────────────────────────────────────────── */
COLS.cls = [
  { k: 'sel', label: '✓', nofilter: true,
    cell: (r) => `<input type="checkbox" data-cls="row:${esc(r.key)}"${S.clsSel.has(r.key) ? ' checked' : ''} aria-label="Select ${esc(r.dn)}">` },
  { k: 'dn', label: 'Source document', cell: (r) => docLink(r.dt, r.dn), plain: (r) => r.dn || '' },
  { k: 'dt', label: 'Type', plain: (r) => r.dt || '' },
  { k: 'supplier_name', label: 'Supplier', cell: (r) => esc(r.supplier_name || r.supplier || '—'),
    plain: (r) => r.supplier_name || r.supplier || '(blank)' },
  { k: 'date', label: 'Posting date', cell: (r) => esc(ddmmmyyyy(r.date)), plain: (r) => r.date || '(blank)' },
  { k: 'amount', label: 'Amount', num: true, cell: (r) => esc(inr(r.amount)), raw: (r) => r.amount },
  { k: 'company_gstin', label: 'Company GSTIN', cell: (r) => esc(r.company_gstin || '—'),
    plain: (r) => r.company_gstin || '(blank)' },
  { k: 'leaf', label: 'Leaf cost centre', cell: (r) => esc(r.leaf || '—'),
    plain: (r) => r.leaf || '(blank)' },
  { k: 'auto', label: 'Automatic result', cell: (r) => locChip(r.auto), plain: (r) => r.auto || '' },
  { k: 'status', label: 'Classification', cell: (r) => `<small>${esc(r.status)}</small>`,
    plain: (r) => r.status || '' },
  { k: 'owner', label: 'Created by', cell: (r) => esc(r.owner || '—'), plain: (r) => r.owner || '(blank)' },
  { k: 'prop', label: 'Proposed AP Location', nofilter: true,
    cell: (r) => `<select data-cls="prop:${esc(r.key)}"><option value="">—</option>` +
      ['BLR', 'JDP'].map((l) => `<option value="${l}"${S.cls && S.cls.proposed[r.key] === l ? ' selected' : ''}>${esc(locLabel(l))}</option>`).join('') +
      '</select>' }];

/* ══ 10 · drill-down drawer ══════════════════════════════════════════════ */

function drill(title, rows, cols, field, expect, valFn, opts) {
  const got = field ? rows.reduce((a, r) => a + num(r[field]), 0)
    : valFn ? rows.reduce((a, r) => a + num(valFn(r)), 0) : null;
  if (S.drawer) S.drawerStack.push(Object.assign({}, S.drawer, { page: S.page.dr }));
  S.drawer = Object.assign({ title, rows, cols, field, expect, got, valFn }, opts || {});
  S.page.dr = 1;
  paintDrawer();
  const d = el('drawer');
  d.hidden = false;
  fitDrawer();
  history.pushState({ fk: 'drawer', depth: S.drawerStack.length + 1, v: navSnap() }, '');
  const c = el('dr-close'); if (c) c.focus();
}

function fitDrawer() {
  const p = el('drawer') && el('drawer').querySelector('.fk-drawer-panel');
  if (!p) return;
  const vh = window.innerHeight || 0;
  if (vh > 0) p.style.maxHeight = Math.max(240, Math.round(0.78 * vh)) + 'px';
}

function drawerBack(fromPop) {
  if (!S.drawer) return;
  if (S.drawerStack.length) {
    S.drawer = S.drawerStack.pop(); S.page.dr = S.drawer.page || 1; paintDrawer();
  } else { el('drawer').hidden = true; S.drawer = null; }
  if (!fromPop) { NAV_SKIP = true; history.back(); }
}
function closeDrawer() {
  const depth = S.drawer ? S.drawerStack.length + 1 : 0;
  el('drawer').hidden = true; S.drawer = null; S.drawerStack = [];
  if (depth) { NAV_SKIP = true; history.go(-depth); }
}

function paintDrawer() {
  const d = S.drawer; if (!d) return;
  el('dr-title').textContent = d.title;
  el('dr-kicker').textContent = countOf(d.rows.length, 'row')
    + (S.drawerStack.length ? ' · level ' + (S.drawerStack.length + 1) : '');
  const back = el('dr-back');
  back.hidden = !S.drawerStack.length;
  back.textContent = S.drawerStack.length
    ? '← Back to ' + S.drawerStack[S.drawerStack.length - 1].title : '← Back';
  el('dr-gl').innerHTML = d.supplier ? glLink(d.supplier, 'Supplier ledger ↗') : '';

  const proof = el('dr-proof');
  if (d.proof) { proof.className = 'fk-drawer-proof' + (d.proofBad ? ' bad' : ''); proof.innerHTML = d.proof; }
  else if (d.got == null || d.expect == null) {
    proof.className = 'fk-drawer-proof';
    proof.innerHTML = `<b>${cnt(d.rows.length)}</b> rows.`;
  } else {
    const diff = r2(d.got - d.expect), ok = Math.abs(diff) <= TOL;
    proof.className = 'fk-drawer-proof' + (ok ? '' : ' bad');
    proof.innerHTML = `${ok ? '✔' : '✖'} Rows <b>${esc(rs(d.got))}</b> · ` +
      `card <b>${esc(rs(d.expect))}</b> · Δ <b>${esc(rs(diff))}</b>`;
  }
  table('dr-table', d.cols, slice(d.rows, S.page.dr), {
    foot: footFor(d.cols, d.rows), empty: 'No rows.' });
  pager('dr-pager', d.rows.length, S.page.dr, (n) => { S.page.dr = n; paintDrawer(); });
}

/* ── the supplier ledger · the drill-down behind every party figure ───────
   Built from GL Entry on the payable account and tied to ERPNext's own
   get_balance_on, so the drawer proves itself against the ERP rather than
   against another view of this dashboard. */
async function openApLedger(supplier, label) {
  if (!supplier) return;
  const nm = label || supplierName(supplier) || supplier;
  const title = 'Supplier ledger · ' + nm;
  const cols = [
    dateCol('date', 'Date'),
    { k: 'vno', label: 'Voucher',
      cell: (r) => r.vno ? `<a class="fk-lnk" href="/app/${esc(r.vslug)}/${q(r.vno)}" target="_blank" rel="noopener">${esc(r.vno)} ↗</a>` : '–' },
    { k: 'part', label: 'Particulars' },
    moneyDash('inv', 'Invoice value'), moneyDash('tds', 'TDS'), moneyDash('net', 'Net payable'),
    moneyDash('paid', 'Paid'), money('bal', 'Balance')];

  if (S.drawer) S.drawerStack.push(Object.assign({}, S.drawer, { page: S.page.dr }));
  S.drawer = { title, cols, rows: [], supplier, proof: 'Reading the ledger…' };
  S.page.dr = 1;
  paintDrawer();
  el('drawer').hidden = false;
  fitDrawer();
  history.pushState({ fk: 'drawer', depth: S.drawerStack.length + 1, v: navSnap() }, '');
  const c = el('dr-close'); if (c) c.focus();

  const slug = (t) => String(t || '').toLowerCase().replace(/\s+/g, '-');
  let g = [];
  try {
    const res = await frappe.call({ method: 'frappe.client.get_list', args: {
      doctype: 'GL Entry',
      filters: { account: AP_ACCOUNT, party: supplier, is_cancelled: 0, posting_date: ['<=', S.asOf] },
      fields: ['posting_date', 'voucher_type', 'voucher_no', 'debit', 'credit'],
      limit_page_length: 0, order_by: 'posting_date asc, creation asc' } });
    g = (res && res.message) || [];
  } catch (e) { console.warn('[AP] supplier ledger', e); }

  const byV = new Map();
  g.forEach((x) => {
    const k = x.voucher_type + '|' + x.voucher_no;
    if (!byV.has(k)) byV.set(k, []);
    byV.get(k).push(x);
  });
  let bal = 0;
  const rows = [];
  byV.forEach((set) => {
    const f = set[0];
    const cr = set.reduce((s, x) => s + num(x.credit), 0);
    const dr = set.reduce((s, x) => s + num(x.debit), 0);
    const isPI = f.voucher_type === 'Purchase Invoice';
    let inv = 0, tds = 0, net = 0, paid = 0;
    if (isPI) { inv = cr; if (dr) tds = -dr; net = cr - dr; }
    else { if (cr) { inv = cr; net = cr; } if (dr) paid = dr; }
    bal += cr - dr;
    rows.push({ date: f.posting_date, vno: f.voucher_no, vslug: slug(f.voucher_type),
      part: f.voucher_type + (isPI && dr ? ' · net of TDS' : ''),
      inv, tds, net, paid, bal: r2(bal) });
  });

  let erpBal = r2(bal);
  try {
    const b = await frappe.call({ method: 'erpnext.accounts.utils.get_balance_on', args: {
      account: AP_ACCOUNT, party_type: 'Supplier', party: supplier, date: S.asOf, company: COMPANY } });
    /* get_balance_on returns the Dr-positive figure; this page is Cr-positive */
    erpBal = r2(-num(b && b.message));
  } catch (e) { console.warn('[AP] get_balance_on', e); }

  const d = r2(bal - erpBal), ok = Math.abs(d) <= TOL;
  S.drawer = { title, cols, rows, supplier,
    proofBad: !ok,
    proof: `${ok ? '✔' : '✖'} Ledger closing <b>${esc(rs(bal))}</b> · ` +
      `ERPNext party balance on ${esc(ddmmmyyyy(S.asOf))} <b>${esc(rs(erpBal))}</b> · ` +
      `Δ <b>${esc(rs(d))}</b> · tolerance ${esc(inr(TOL))}<br>` +
      `Account <b>${esc(AP_ACCOUNT)}</b> · cut-off <b>${esc(ddmmmyyyy(S.asOf))}</b> · ` +
      `${esc(cnt(rows.length))} vouchers · Cr positive = owed` };
  paintDrawer();
}
function supplierName(sup) {
  if (S.roster && S.roster.has(sup)) return S.roster.get(sup);
  if (S.dims.suppliers.has(sup)) return S.dims.suppliers.get(sup);
  return '';
}

/* ══ 11 · row expansion — item and tax bifurcation in place ══════════════ */

/* exact row lookup — never a suffix match, because one document id can be a
   suffix of another (PINV-26-0412 inside PINV-26-04129) */
function findRow(tableKey, fullKey) {
  const t = el(tableKey); if (!t) return null;
  const rows = t.querySelectorAll('tr[data-row]');
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].getAttribute('data-row') === fullKey) return rows[i];
  }
  return null;
}
function collapseKids(tr) {
  let n = tr.nextElementSibling;
  if (!n || !n.classList.contains('kid')) return false;
  while (n && n.classList.contains('kid')) { const x = n.nextElementSibling; n.remove(); n = x; }
  tr.classList.remove('exp');
  return true;
}
function renderKids(tr, cols, kids, note) {
  tr.classList.add('exp');
  const host = document.createElement('tr');
  host.className = 'kid';
  host.innerHTML = `<td class="txt" colspan="${tr.children.length}"><table class="fk-tbl"></table></td>`;
  tr.after(host);
  const inner = host.querySelector('table');
  if (!kids.length) { inner.outerHTML = esc(note || 'No lines readable for this document.'); return; }
  inner.innerHTML = '<thead><tr>' + cols.map((c) =>
      `<th class="${c.num ? 'num' : ''}">${esc(c.label)}</th>`).join('') + '</tr></thead>'
    + '<tbody>' + kids.map((r) => '<tr>' + cols.map((c) =>
        `<td class="${c.num ? 'num' : 'txt'}">${c.cell ? c.cell(r) : esc(r[c.k])}</td>`).join('') + '</tr>').join('') + '</tbody>'
    + '<tfoot><tr>' + cols.map((c, i) => {
        if (i === 0) return `<td class="txt">${cnt(kids.length)} lines${note ? ' · ' + esc(note) : ''}</td>`;
        if (c.num && c.raw) return `<td class="num">${esc(rs(kids.reduce((a, r) => a + num(c.raw(r)), 0)))}</td>`;
        return '<td></td>';
      }).join('') + '</tr></tfoot>';
}
async function expandDoc(tableKey, prefix, id, cacheKey, rowType, cols, matchField, note) {
  const tr = findRow(tableKey, prefix + id);
  if (!tr || collapseKids(tr)) return;
  const rows = await lazy(cacheKey, rowType, {}, rowType + ' lines');
  renderKids(tr, cols, rows.filter((r) => String(r[matchField]) === String(id)), note);
}

/* ══ 12 · exports · one writer each for CSV and XLSX ═════════════════════ */

function dl(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.style.display = 'none';
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}

const cellValue = (c, r) => c.raw ? c.raw(r) : c.plain ? c.plain(r) : r[c.k];

/* Every CSV carries its own provenance, so a file that leaves this page can
   still say what it is, what it was filtered to, and when it was taken. */
function csvExport(name, cols, rows, extra) {
  const qq = (v) => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  const head = [
    ['Furnishka Tech Private Limited — Accounts Payable'],
    ['Report', name],
    ['Period', ddmmmyyyy(S.from) + ' to ' + ddmmmyyyy(S.to)],
    ['Ledger cut-off', ddmmmyyyy(S.asOf)],
    ['Ageing basis', AG_BASIS[S.agBasis].label],
    ['AP Location / GSTIN', S.loc
      ? locLabel(S.loc) + (locOf(S.loc).gstin ? ' · ' + locOf(S.loc).gstin : '')
      : 'Combined (BLR + JDP + Unassigned)'],
    ['Scope', [S.cc || 'All cost centres', S.supplier, S.sg, S.dept, S.src, S.recon]
      .filter(Boolean).join(' · ')],
    ['Source', 'Query Report: AP Procure to Pay Master (server-side, read-only)'],
    ['Payable account', AP_ACCOUNT],
    ['Sign convention', 'Credit minus debit. Cr positive = owed to the supplier. '
      + 'Dr negative = net receivable, reclassified per Schedule III and never netted.'],
    ['Run at', frappe.datetime.now_datetime()],
    ['Run by', frappe.session.user],
    ['Numbers', 'Indian format, paise retained, not abbreviated'],
    []
  ].map((r) => r.map(qq).join(',')).join('\n');

  const body = [cols.map((c) => qq(c.label)).join(',')]
    .concat(rows.map((r) => cols.map((c) => {
      const v = cellValue(c, r);
      return typeof v === 'number' ? String(r2(v)) : qq(v);   /* numbers stay numeric cells */
    }).join(','))).join('\n');

  const tail = (extra && extra.length)
    ? '\n\n' + extra.map((r) => r.map((v) =>
        typeof v === 'number' ? String(r2(v)) : qq(v)).join(',')).join('\n')
    : '';

  dl(new Blob(['﻿' + head + '\n' + body + tail], { type: 'text/csv;charset=utf-8;' }),
    'Furnishka_AP_' + name.replace(/[^\w]+/g, '_') + '_' + S.asOf + '.csv');
}

/* ── one XLSX writer ──────────────────────────────────────────────────────
   Replaces three near-identical copies (v24 single-sheet, v24 multi-sheet,
   v27 multi-sheet). Styles: 1 bold, 2 amount, 3 bold amount, 4 header,
   5 title, 6 integer. A sheet is { name, rows, widths, freeze, af, sel },
   a row is an array of [value, style] pairs or nulls.
   Number format is the Indian 2-2-3 grouping, so a figure opened in Excel
   reads the way it reads on this page. */
/* A1-style column name, declared at this level so a caller can size an
   autofilter range to the columns it actually wrote instead of guessing. */
function xlsxCol(n) {
  let s = ''; n++;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}
function xlsxBook(sheets, name) {
  const enc = new TextEncoder();
  const TBL = (() => { const a = new Uint32Array(256);
    for (let i = 0; i < 256; i++) { let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      a[i] = c >>> 0; } return a; })();
  const crc = (u) => { let c = 0xFFFFFFFF;
    for (let i = 0; i < u.length; i++) c = TBL[(c ^ u[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0; };
  const xx = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const colName = xlsxCol;
  const STYLE = { n: 2, N: 3, i: 6, h: 4, b: 1, T: 5 };

  const sheetXml = (sh) => {
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<sheetViews><sheetView workbookViewId="0"' + (sh.sel ? ' tabSelected="1"' : '') + '>'
      + (sh.freeze ? `<pane ySplit="${sh.freeze}" topLeftCell="A${sh.freeze + 1}" activePane="bottomLeft" state="frozen"/>` : '')
      + '</sheetView></sheetViews>';
    if (sh.widths) x += '<cols>' + sh.widths.map((w, i) =>
      `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>';
    x += '<sheetData>';
    (sh.rows || []).forEach((row, ri) => {
      const r = ri + 1;
      let cs = '';
      (row || []).forEach((cv, ci) => {
        if (!cv) return;
        const v = cv[0], k = cv[1];
        if (v === null || v === undefined || v === '') return;
        const ref = colName(ci) + r;
        if (k === 'n' || k === 'N' || k === 'i') cs += `<c r="${ref}" s="${STYLE[k]}"><v>${k === 'i' ? num(v) : r2(v)}</v></c>`;
        else if (STYLE[k]) cs += `<c r="${ref}" s="${STYLE[k]}" t="inlineStr"><is><t>${xx(v)}</t></is></c>`;
        else cs += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xx(v)}</t></is></c>`;
      });
      if (cs) x += `<row r="${r}">${cs}</row>`;
    });
    x += '</sheetData>';
    if (sh.af) x += `<autoFilter ref="${sh.af}"/>`;
    return x + '</worksheet>';
  };

  const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<numFmts count="1"><numFmt numFmtId="164" formatCode="[&gt;=100000]##\\,##\\,##0.00;[&lt;=-100000]\\-##\\,##\\,##0.00;##,##0.00"/></numFmts>'
    + '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="14"/><name val="Calibri"/></font></fonts>'
    + '<fills count="3"><fill><patternFill patternType="none"/></fill>'
    + '<fill><patternFill patternType="gray125"/></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FFEEF0F3"/><bgColor indexed="64"/></patternFill></fill></fills>'
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

  let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
  let wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
  let rel = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  let parts = [];
  sheets.forEach((sh, i) => {
    ct += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    wb += `<sheet name="${xx(sh.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
    rel += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
    parts.push(['xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(sh)]);
  });
  ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
  wb += '</sheets></workbook>';
  rel += `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;

  const files = [
    ['[Content_Types].xml', ct],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', wb],
    ['xl/_rels/workbook.xml.rels', rel],
    ['xl/styles.xml', styles]
  ].concat(parts);

  const u16 = (n) => [n & 255, (n >> 8) & 255];
  const u32 = (n) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
  const nw = new Date();
  const dt = ((nw.getHours() << 11) | (nw.getMinutes() << 5) | (nw.getSeconds() >> 1)) & 0xFFFF;
  const dd = (((nw.getFullYear() - 1980) << 9) | ((nw.getMonth() + 1) << 5) | nw.getDate()) & 0xFFFF;
  const chunks = [], central = [];
  let off = 0;
  files.forEach((f) => {
    const nb = enc.encode(f[0]), db = enc.encode(f[1]), c = crc(db);
    const lh = [].concat([80, 75, 3, 4], u16(20), u16(0), u16(0), u16(dt), u16(dd),
      u32(c), u32(db.length), u32(db.length), u16(nb.length), u16(0));
    chunks.push(new Uint8Array(lh), nb, db);
    central.push({ nb, c, s: db.length, o: off });
    off += lh.length + nb.length + db.length;
  });
  const cd = [];
  central.forEach((e) => {
    cd.push(new Uint8Array([].concat([80, 75, 1, 2], u16(20), u16(20), u16(0), u16(0),
      u16(dt), u16(dd), u32(e.c), u32(e.s), u32(e.s), u16(e.nb.length),
      u16(0), u16(0), u16(0), u16(0), u32(0), u32(e.o))), e.nb);
  });
  const cdLen = cd.reduce((a, x) => a + x.length, 0);
  const eocd = new Uint8Array([].concat([80, 75, 5, 6], u16(0), u16(0),
    u16(central.length), u16(central.length), u32(cdLen), u32(off), u16(0)));
  dl(new Blob(chunks.concat(cd, [eocd]),
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name + '.xlsx');
}

/* ══ 13 · painting ═══════════════════════════════════════════════════════ */

function paintAll() {
  paintScope(); paintTabs(); paintTimeline();
  paintOverview(); paintPO(); paintGRN(); paintInvoices();
  paintAgeing(); paintVCred();
  if (S.tab === 'vbal') ensureVBal();
  if (S.tab === 'vcred') ensureVCred();
  if (S.tab === 'cls' && !S.cls) clsOpen();
  el('asof').innerHTML = 'Period to ' + esc(ddmmmyyyy(S.to))
    + '<br>ledger as on ' + esc(ddmmmyyyy(S.asOf))
    + '<br>run ' + esc(frappe.datetime.now_datetime().slice(0, 16));
}

function paintScope() {
  const L = S.L;
  const scope = [S.loc ? locLabel(S.loc) : '', S.cc].filter(Boolean).join(' · ');
  el('scope').innerHTML =
    `<b>${esc(ddmmmyyyy(S.from))} – ${esc(ddmmmyyyy(S.to))}</b> · ledger cut-off <b>${esc(ddmmmyyyy(S.asOf))}</b> · `
    + (scope ? `scope <b>${esc(scope)}</b> · ` : 'AP location <b>Combined</b> · ')
    + (S.supplier ? `supplier <b>${esc(supplierName(S.supplier) || S.supplier)}</b> · ` : '')
    + (S.sg ? `group <b>${esc(S.sg)}</b> · ` : '')
    + (S.dept ? `department <b>${esc(S.dept)}</b> · ` : '')
    + (S.src ? `route <b>${esc(S.src)}</b> · ` : '')
    + (S.recon ? `bank state <b>${esc(S.recon)}</b> · ` : '')
    + (!L.scopable ? '<span>Ledger tie not applicable at cost-centre or department scope</span>'
       : L.ok ? `<span class="ok">✔ Ledger tie Δ ${esc(rs(L.variance))}</span>`
       : `<span class="bad">✖ Ledger Δ ${esc(rs(L.variance))}</span>`);
  el('foot-scope').textContent =
    `${ddmmmyyyy(S.from)} – ${ddmmmyyyy(S.to)} · ${frappe.session.user}`;
}

const TABS = [
  { k: 'overview', label: 'Overview',          val: () => rs(S.L.openTotal) },
  { k: 'po',       label: 'Purchase Orders',   val: () => rs(S.D.po.value) },
  { k: 'grn',      label: 'Goods Receipt',     val: () => rs(S.D.pr.value) },
  { k: 'inv',      label: 'Invoices & Payments', val: () => rs(S.D.pi.gross) },
  { k: 'vcred',    label: 'Vendor Credits',
    val: () => rs(r2(vcReturns().reduce((a, r) => a + r.vc_avail, 0) + num(S.L.unappliedV))) },
  { k: 'cls',      label: 'Categorise AP Location',
    val: () => cnt(S.clsPending || 0) + ' to classify', show: CAN_CLASSIFY },
  { k: 'ageing',   label: 'AP Aging',          val: () => rs(r2(S.A.billTot + S.A.dn)) },
  { k: 'vbal',     label: 'Vendor Balances',   val: () => rs(S.L.controlV) }
];
function paintTabs() {
  if (S.tab === 'cls' && !CAN_CLASSIFY()) S.tab = 'overview';
  el('tabs').innerHTML = TABS.filter((t) => !t.show || t.show()).map((t) =>
    '<button type="button" role="tab" data-t="' + t.k + '" aria-selected="' + (t.k === S.tab) + '">'
    + '<span>' + esc(t.label) + '</span><span class="fk-tabval">'
    + esc(S.loaded ? t.val() : '—') + '</span></button>').join('');
  TABS.forEach((t) => { const p = el('panel-' + t.k); if (p) p.hidden = (t.k !== S.tab); });
}

/* ── timeline ───────────────────────────────────────────────────────────── */
function paintTimeline() {
  const rows = S.D.series || [];
  el('tl-legend').innerHTML = SERIES.map((s) =>
    `<span data-s="${s.k}" role="button" tabindex="0" aria-pressed="${!S.tlHidden.has(s.k)}">`
    + `<i style="background:${s.colour}"></i>${esc(s.label)}</span>`).join('')
    + `<span style="margin-left:auto;color:var(--muted);cursor:default">`
    + `${S.tlMode === 'cum' ? 'Cumulative' : 'Per ' + S.grain}</span>`;

  const svg = el('timeline');
  if (!rows.length) { svg.innerHTML = emptySvg('No documents in this period'); el('tl-foot').innerHTML = ''; return; }
  const W = 1200, H = 300, p = { l: 128, r: 22, t: 18, b: 46 };
  const iw = W - p.l - p.r, ih = H - p.t - p.b;
  const live = SERIES.filter((s) => !S.tlHidden.has(s.k));
  let max = 0; rows.forEach((r) => live.forEach((s) => { if (r[s.k] > max) max = r[s.k]; }));
  max = max * 1.14 || 1;
  const x = (i) => rows.length === 1 ? p.l + iw / 2 : p.l + i * iw / (rows.length - 1);
  const y = (v) => p.t + ih - (v / max) * ih;
  let h = '';
  for (let i = 0; i <= 4; i++) { const v = max * i / 4, yy = y(v);
    h += `<line class="gl" x1="${p.l}" x2="${W - p.r}" y1="${yy}" y2="${yy}"/>`;
    h += `<text class="tick" text-anchor="end" x="${p.l - 9}" y="${yy + 3.5}">${esc(rs(v))}</text>`; }
  h += `<line class="axis" x1="${p.l}" x2="${W - p.r}" y1="${y(0)}" y2="${y(0)}"/>`;
  const step = Math.ceil(rows.length / 14);
  rows.forEach((r, i) => { if (i % step === 0 || i === rows.length - 1)
    h += `<text class="tick" text-anchor="middle" x="${x(i)}" y="${p.t + ih + 18}">${esc(bucketLabel(r.k))}</text>`; });
  live.forEach((s) => {
    h += `<polyline points="${rows.map((r, i) => x(i).toFixed(1) + ',' + y(r[s.k]).toFixed(1)).join(' ')}" `
       + `fill="none" stroke="${s.colour}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    if (rows.length <= 40) rows.forEach((r, i) => {
      h += `<circle class="hit" data-stage="${s.k}" data-bucket="${esc(r.k)}" cx="${x(i)}" cy="${y(r[s.k])}" `
         + `r="3.4" fill="${s.colour}" stroke="#fff" stroke-width="1.4">`
         + `<title>${esc(s.label)} · ${esc(bucketLabel(r.k))}: ${esc(rs(r[s.k]))}</title></circle>`; });
  });
  svg.innerHTML = h;

  const tot = {};
  SERIES.forEach((s) => tot[s.k] = rows.reduce((a, r) => a + r[s.k], 0));
  if (S.tlMode === 'cum' && rows.length) SERIES.forEach((s) => tot[s.k] = rows[rows.length - 1][s.k]);
  el('tl-foot').innerHTML = SERIES.map((s) =>
    `<span><b>${rs(tot[s.k])}</b> ${esc(s.label.toLowerCase())}</span>`).join('');
}

/* ══ 14 · Overview ═══════════════════════════════════════════════════════ */

let ROUTE_REG = [];

function paintOverview() {
  const D = S.D, L = S.L;
  const p2pRet = where(D.pi.p2p, IS_RETURN), dirRet = where(D.pi.direct, IS_RETURN);
  const over90 = r2(S.A.bands ? S.A.bands[4] + S.A.bands[5] : 0);
  kpi('ov-kpis', [
    { k: 'Inventory P2P invoices', v: rs(D.pi.p2pV), raw: D.pi.p2pV, tone: 'var(--s1)',
      c: countOf(D.pi.p2p.length - p2pRet.length, 'invoice') + ' · ' + countOf(p2pRet.length, 'debit note') + ' · '
        + pct(Math.abs(D.pi.p2pV), Math.abs(D.pi.gross)) + ' of booked · open ' + rs(L.openP2PV),
      drill: () => drill('Inventory P2P invoices', D.pi.p2p, COLS.piFull, 'gross', D.pi.p2pV) },
    { k: 'Direct expense / service bills', v: rs(D.pi.directV), raw: D.pi.directV, tone: 'var(--s2)',
      c: countOf(D.pi.direct.length - dirRet.length, 'invoice') + ' · ' + countOf(dirRet.length, 'debit note') + ' · '
        + pct(Math.abs(D.pi.directV), Math.abs(D.pi.gross)) + ' of booked · open ' + rs(L.openDirV),
      drill: () => drill('Direct expense / service bills', D.pi.direct, COLS.piFull, 'gross', D.pi.directV) },
    { k: 'AP outstanding at cut-off', v: rs(L.openTotal), raw: L.openTotal, tone: 'var(--s5)',
      c: countOf(L.open.length, 'open invoice') + ' · P2P ' + rs(L.openP2PV) + ' · direct ' + rs(L.openDirV)
        + ' · over 90 days ' + rs(over90),
      drill: () => drill('AP outstanding at cut-off', L.open, COLS.openItems, 'outstanding', L.openTotal) },
    { k: 'Supplier payments (period)', v: rs(D.cashOut), raw: D.cashOut, tone: 'var(--s4)',
      c: countOf(D.payN, 'payment entry') + ' · to P2P invoices ' + rs(D.p2pAllocV)
        + ' · to direct bills ' + rs(D.dirAllocV) + ' · unallocated ' + rs(D.unallocField),
      drill: () => drill('Supplier payments', D.peAlloc.concat(D.peUnalloc), COLS.pay, null, D.cashOut, payVal) },
    { k: 'Unbilled GRN exposure', v: rs(D.pr.unbilled), raw: D.pr.unbilled, tone: 'var(--s2)',
      c: countOf(D.pr.open.length, 'receipt') + ' not fully billed · PO pending receipt ' + rs(D.po.pending),
      drill: () => drill('Unbilled receipts', D.pr.open, COLS.prFull, 'pending_bill', D.pr.unbilled) },
    { k: 'Opening open items', v: rs(L.openingV), raw: L.openingV, tone: 'var(--muted)',
      c: countOf(L.openingOpen.length, 'bill') + ' carried in, still open at ' + ddmmmyyyy(S.asOf)
        + ' · of ' + countOf(D.opening.length, 'invoice') + ' raised before the period',
      drill: () => drill('Opening open items', L.openingOpen, COLS.openItems, 'outstanding', L.openingV) },
    { k: 'Unapplied supplier cash', v: rs(L.unappliedV), raw: L.unappliedV, tone: 'var(--s4)',
      c: countOf(L.unapplied.length, 'payment entry') + ' · '
        + cnt(new Set(L.unapplied.map((r) => r.supplier)).size) + ' vendors · click to reconcile',
      go: () => openReconciliation(S.supplier) },
    { k: 'AP control tie', v: L.scopable ? rs(L.variance) : 'n/a', raw: L.scopable ? L.variance : 0,
      tone: !L.scopable ? 'var(--muted)' : L.ok ? 'var(--s3)' : 'var(--s5)',
      c: !L.scopable ? '' : (L.ok ? '✔ within ₹0.50 tolerance' : '✖ outside ₹0.50 tolerance')
        + ' · open items ' + rs(L.openTotal) + ' + journals ' + rs(L.jrNonInvV)
        + ' − unapplied ' + rs(L.unappliedV) + ' · control ' + rs(L.controlV),
      drill: () => drillControlTie() }
  ]);

  /* two booking routes, never mixed */
  const routes = SOURCE_CLASSES.map((c) => {
    const g = D.pi.byRoute.find((x) => x.key === c);
    return { label: c, v: g ? g.v : 0, n: g ? g.n : 0, rows: g ? g.rows : [],
      colour: c === 'Inventory P2P' ? 'var(--s1)' : 'var(--s2)',
      icon: c === 'Inventory P2P' ? '✔' : '◆' };
  });
  ROUTE_REG = routes;
  const rmax = Math.max.apply(null, routes.map((r) => Math.abs(r.v))) || 1;
  el('ov-routes').innerHTML = routes.map((r, i) =>
    `<div class="fk-outrow hit" data-route="${i}" role="button" tabindex="0">`
    + `<span class="nm"><span>${r.icon}</span>${esc(r.label)}</span>`
    + `<span class="track"><i style="width:${(Math.abs(r.v) / rmax * 100).toFixed(1)}%;background:${r.colour}"></i></span>`
    + `<span class="amt">${rs(r.v)}<small>${cnt(r.n)} documents · ${pct(Math.abs(r.v), Math.abs(D.pi.gross))}</small></span></div>`).join('')
    + `<div class="fk-outrow"><span class="nm" style="font-weight:700">Net booked purchase</span>`
    + `<span class="track"></span><span class="amt">${rs(D.pi.gross)}</span></div>`;

  paintLocationRecon();
}

/* the subledger-to-GL bridge, as a drill rather than a permanently hidden card */
function drillControlTie() {
  const L = S.L;
  const rows = [
    { l: 'Open items at cut-off', v: L.openTotal, g: 'PI Doc',
      note: countOf(L.open.length, 'invoice') + ' with a non-zero ledger balance' },
    { l: '+ Journals not linked to an invoice', v: L.jrNonInvV, g: 'AP GL Journal',
      note: cnt(L.jrNonInv.length) + ' opening and manual entries on the AP control' },
    { l: '− Unapplied cash', v: -L.unappliedV, g: 'AP Unapplied',
      note: 'derived per supplier; the ERP unallocated field is lower' },
    { l: 'Memo · payments carrying a payable balance', v: L.ppV, g: 'AP Aging',
      note: countOf(L.ppRows.length, 'payment') + ' on the payable side — already inside the '
        + 'unapplied line above as a negative advance, so it is NOT added again' },
    { l: '= Computed AP control', v: L.computed, g: '—', note: 'what the subledger says' },
    { l: 'AP control account balance', v: L.controlV, g: 'AP GL Control',
      note: cnt(L.control.length) + ' supplier ledger rows at the cut-off' },
    { l: 'Variance', v: L.variance, g: '—',
      note: L.ok ? 'within the ₹0.50 tolerance' : 'OUTSIDE the ₹0.50 tolerance' },
    { l: 'Memo · journals already inside open items', v: L.jrOnInvV, g: 'AP GL Journal',
      note: 'invoice-linked — excluded above so they are not counted twice' }
  ];
  drill('AP subledger to general ledger', rows, [
    { k: 'l', label: 'Line' },
    { k: 'v', label: '₹', num: true, raw: (r) => r.v, cell: (r) => esc(inr(r.v)) },
    { k: 'g', label: 'Grain' }, { k: 'note', label: 'Reading' }
  ], null, null, null, {
    proofBad: !L.ok,
    proof: (L.ok ? '✔' : '✖') + ' Open items measured from ledger movement at '
      + esc(ddmmmyyyy(S.asOf)) + ', not from the current outstanding field, so a past cut-off '
      + 'returns the position as it stood then. Variance <b>' + esc(rs(L.variance))
      + '</b> against a tolerance of <b>' + esc(inr(TOL)) + '</b>.' });
}

/* AP subledger to ledger, proved separately for every registration. Every grain
   is classified by the document that owns it, so the lines add back to Combined
   with no voucher counted twice. */
let LOCREC = [];
function paintLocationRecon() {
  const L = S.L;
  const lines = LOCS.map((l) => l.k).concat(['__all']).map((k, i) => {
    const f = (r) => k === '__all' || (r.ap_location || 'Unassigned') === k;
    const open = where(L.open, f), jr = where(L.jrNonInv, f),
          un = where(L.unapplied, f), ct = where(L.control, f);
    const O = sum(open, 'outstanding'), J = sum(jr, 'outstanding'),
          U = sum(un, 'advance'), Cv = sum(ct, 'outstanding');
    return { i, k, label: k === '__all' ? 'Combined' : locLabel(k),
      gstin: locOf(k).gstin || '', O, J, U, C: Cv, V: r2(O + J - U - Cv),
      open, jr, un, ct };
  });
  LOCREC = lines;
  const cell = (r, part, v) =>
    `<button type="button" class="fk-numbtn" data-locrec="${r.i}:${part}">${esc(inr(v))}</button>`;
  table('ov-locrec', [
    { k: 'label', label: 'AP location / GSTIN',
      cell: (r) => `<b>${esc(r.label)}</b>${r.gstin ? '<br><small>' + esc(r.gstin) + '</small>' : ''}` },
    { k: 'O', label: 'Open items', num: true, raw: (r) => r.O, cell: (r) => cell(r, 'O', r.O) },
    { k: 'J', label: '+ Journals', num: true, raw: (r) => r.J, cell: (r) => cell(r, 'J', r.J) },
    { k: 'U', label: '− Unapplied cash', num: true, raw: (r) => r.U, cell: (r) => cell(r, 'U', r.U) },
    { k: 'C', label: 'AP control', num: true, div: true, raw: (r) => r.C, cell: (r) => cell(r, 'C', r.C) },
    { k: 'V', label: 'Variance', num: true, raw: (r) => r.V,
      cell: (r) => !S.L.scopable ? chip('n/a at this scope', 'c-neu')
        : chip((Math.abs(r.V) <= TOL ? '✔ ' : '✖ ') + inr(r.V),
               Math.abs(r.V) <= TOL ? 'c-ok' : 'c-crit') }
  ], lines);

  const disputed = [].concat(S.D.piDoc || [], S.D.poDoc || [], S.D.prDoc || [], S.L.openAll || [])
    .filter((r) => r.conflict_details)
    .map((r) => r.classification_doctype + '|' + r.classification_docname)
    .filter((v, i, a) => a.indexOf(v) === i).length;
  el('ov-locnote').innerHTML = 'Allocations follow the invoice they settle; unallocated cash follows '
    + 'the payment; journals against an invoice follow the invoice. Click any figure for the rows behind it.'
    + (disputed ? ' · <b>' + cnt(disputed) + '</b> documents carry a GST registration that disagrees '
      + 'with their cost centre — listed on the Categorise AP Location tab.' : '');
}

/* ══ 15 · Purchase Orders ════════════════════════════════════════════════ */

function paintPO() {
  const D = S.D;
  kpi('po-kpis', [
    { k: 'Purchase order total value', v: rs(D.po.value), raw: D.po.value, tone: 'var(--s1)',
      c: countOf(D.po.n, 'submitted order') + ' in the period',
      drill: () => drill('Purchase order total value', D.po.rows, COLS.poFull, 'gross', D.po.value) },
    /* no Payment Entry in this ERP references a Purchase Order, so this card
       reports the real advance position instead of a PO advance of 0.00 */
    { k: 'Advances to inventory vendors', v: rs(S.L.unappliedInvV), raw: S.L.unappliedInvV, tone: 'var(--s4)',
      c: countOf(S.L.unappliedInv.length, 'payment entry') + ' · of ' + rs(S.L.unappliedV)
        + ' unapplied across all vendors · none is PO-linked',
      drill: () => drill('Advances to inventory vendors', S.L.unappliedInv, COLS.unapplied, 'advance', S.L.unappliedInvV) },
    { k: 'Closed purchase order value', v: rs(D.po.closedValue), raw: D.po.closedValue, tone: 'var(--s3)',
      c: cnt(D.po.closed.length) + ' completed or manually closed',
      drill: () => drill('Closed purchase orders', D.po.closed, COLS.poFull, 'gross', D.po.closedValue) },
    { k: 'Open purchase order value', v: rs(D.po.openValue), raw: D.po.openValue, tone: 'var(--s2)',
      c: 'received ' + rs(D.po.received) + ' · yet to receive ' + rs(D.po.pending),
      drill: () => drill('Open purchase orders', D.po.open, COLS.poFull, 'gross', D.po.openValue) }
  ]);

  barChart('po-status', D.po.byStatus.map((g) => ({
    key: 'status:' + g.key, label: g.key, v: g.v, sub: countOf(g.n, 'order'),
    colour: /hold|stopped|closed/i.test(g.key) ? 'var(--s5)' : 'var(--s1)'
  })), { empty: 'No purchase orders in this period' });

  barChart('po-bank', D.bankAdvInv.filter((b) => b.v !== 0 || b.n).map((b) => ({
    key: 'bankadv:' + b.key, label: b.key, v: b.v, sub: countOf(b.n, 'payment entry'),
    colour: /card/i.test(b.key) ? 'var(--s5)' : 'var(--s4)'
  })), { empty: 'No supplier advances in this period' });

  const q0 = S.search.po.trim().toLowerCase();
  const rows = q0 ? D.po.rows.filter((r) => matches(q0, r.grain_key, r.supplier_name)) : D.po.rows;
  rows.forEach((r) => r.__key = 'po:' + r.grain_key);
  table('po-table', COLS.poFull, slice(rows, S.page.po), {
    foot: footFor(COLS.poFull, rows), empty: 'No purchase orders in this scope.' });
  pager('po-pager', rows.length, S.page.po, (n) => { S.page.po = n; paintPO(); });
  paintCohort();
}

function paintCohort() {
  const D = S.D;
  const rows = S.coView === 'cohort' ? D.cohort : (D.series || []).map((r) => ({
    k: r.k, ordered: r.ordered, received: r.received, invoiced: r.invoiced,
    pendingReceipt: 0, pendingBill: 0, n: 0, paid: r.paid, rows: [] }));
  const svg = el('po-cohort');
  if (!rows.length) { svg.innerHTML = emptySvg('Nothing in this period'); table('co-table', COLS.cohort, []); return; }

  const W = 1200, H = 300, p = { l: 128, r: 22, t: 18, b: 46 };
  const iw = W - p.l - p.r, ih = H - p.t - p.b;
  const keys = S.coView === 'cohort'
    ? [['ordered', 'Ordered', 'var(--s1)'], ['received', 'Received', 'var(--s3)'],
       ['invoiced', 'Invoiced', 'var(--s2)'], ['pendingReceipt', 'Pending receipt', 'var(--s5)']]
    : [['ordered', 'Ordered', 'var(--s1)'], ['received', 'Received', 'var(--s3)'],
       ['invoiced', 'Invoiced', 'var(--s2)'], ['paid', 'Paid', 'var(--s4)']];
  let max = 0; rows.forEach((r) => keys.forEach((k) => { max = Math.max(max, num(r[k[0]])); }));
  max = max * 1.14 || 1;
  const n = rows.length, slot = iw / n, bw = Math.min(16, slot / (keys.length + 1));
  const cx = (i) => p.l + slot * i + slot / 2, y = (v) => p.t + ih - (v / max) * ih;
  let h = '';
  for (let i = 0; i <= 4; i++) { const v = max * i / 4, yy = y(v);
    h += `<line class="gl" x1="${p.l}" x2="${W - p.r}" y1="${yy}" y2="${yy}"/>`;
    h += `<text class="tick" text-anchor="end" x="${p.l - 9}" y="${yy + 3.5}">${esc(rs(v))}</text>`; }
  h += `<line class="axis" x1="${p.l}" x2="${W - p.r}" y1="${y(0)}" y2="${y(0)}"/>`;
  const step = Math.ceil(n / 14);
  rows.forEach((r, i) => {
    keys.forEach((k, j) => {
      const v = num(r[k[0]]);
      const x0 = cx(i) - (keys.length * bw) / 2 + j * bw;
      h += `<rect class="hit" data-cohort="${esc(r.k)}" x="${x0}" y="${y(v)}" width="${Math.max(1, bw - 1.5)}" `
         + `height="${Math.max(1, y(0) - y(v))}" rx="2" fill="${k[2]}" opacity=".9">`
         + `<title>${esc(k[1])} · ${esc(bucketLabel(r.k))}: ${esc(rs(v))}</title></rect>`;
    });
    if (i % step === 0 || i === n - 1)
      h += `<text class="tick" text-anchor="middle" x="${cx(i)}" y="${p.t + ih + 18}">${esc(bucketLabel(r.k))}</text>`;
  });
  keys.forEach((k, j) => {
    h += `<rect x="${p.l + j * 148}" y="${p.t - 12}" width="11" height="4" rx="2" fill="${k[2]}"/>`;
    h += `<text class="tick" x="${p.l + j * 148 + 16}" y="${p.t - 8}" style="font-weight:700">${esc(k[1])}</text>`;
  });
  svg.innerHTML = h;

  const cols = S.coView === 'cohort' ? COLS.cohort.filter((c) => c.k !== 'paid') : COLS.cohort;
  table('co-table', cols, rows.map((r) => Object.assign({ __key: 'co:' + r.k }, r)), {
    foot: footFor(cols, rows, 'Total') });
}

/* ══ 16 · Goods Receipt ══════════════════════════════════════════════════ */

function paintGRN() {
  const D = S.D;
  kpi('gr-kpis', [
    { k: 'Purchase receipt / GRN total', v: rs(D.pr.value), raw: D.pr.value, tone: 'var(--s3)',
      c: countOf(D.pr.n, 'receipt') + ' · net of ' + countOf(D.pr.returns.length, 'return'),
      drill: () => drill('GRN total value', D.pr.rows, COLS.prFull, 'gross', D.pr.value) },
    { k: 'Fully billed GRN value', v: rs(D.pr.fullyValue), raw: D.pr.fullyValue, tone: 'var(--s1)',
      c: countOf(D.pr.fully.length, 'receipt') + ' fully invoiced at the cut-off',
      drill: () => drill('Fully billed GRN', D.pr.fully, COLS.prFull, 'gross', D.pr.fullyValue) },
    /* a submitted receipt has already recorded the receipt, so the unbilled
       value — not an "open receipt" count — is the honest exposure */
    { k: 'Unbilled GRN value', v: rs(D.pr.unbilled), raw: D.pr.unbilled, tone: 'var(--s5)',
      c: 'of ' + rs(D.pr.value) + ' received · billed ' + rs(D.pr.billed),
      drill: () => drill('Unbilled GRN', D.pr.open, COLS.prFull, 'pending_bill', D.pr.unbilled) },
    { k: 'Pending receipt against PO', v: rs(D.po.pending), raw: D.po.pending, tone: 'var(--s2)',
      c: 'ordered ' + rs(D.po.value) + ' · received ' + rs(D.po.received),
      drill: () => drill('Pending receipt against PO',
        where(D.po.rows, (r) => num(r.pending_receipt) > 0.005), COLS.poFull, 'pending_receipt', D.po.pending) }
  ]);

  /* received against invoiced */
  const rows = D.series || [];
  const svg = el('gr-trend');
  if (!rows.length) svg.innerHTML = emptySvg('No documents in this period');
  else {
    const W = 1200, H = 260, p = { l: 128, r: 22, t: 22, b: 42 };
    const iw = W - p.l - p.r, ih = H - p.t - p.b;
    let max = 0; rows.forEach((r) => { max = Math.max(max, r.received, r.invoiced); });
    max = max * 1.14 || 1;
    const n = rows.length, slot = iw / n, bw = Math.min(24, slot * 0.34);
    const cx = (i) => p.l + slot * i + slot / 2, y = (v) => p.t + ih - (v / max) * ih;
    let h = '';
    for (let i = 0; i <= 4; i++) { const v = max * i / 4, yy = y(v);
      h += `<line class="gl" x1="${p.l}" x2="${W - p.r}" y1="${yy}" y2="${yy}"/>`;
      h += `<text class="tick" text-anchor="end" x="${p.l - 9}" y="${yy + 3.5}">${esc(rs(v))}</text>`; }
    h += `<line class="axis" x1="${p.l}" x2="${W - p.r}" y1="${y(0)}" y2="${y(0)}"/>`;
    const step = Math.ceil(n / 14);
    rows.forEach((r, i) => {
      h += `<rect x="${cx(i) - bw}" y="${y(r.received)}" width="${bw}" height="${Math.max(1, y(0) - y(r.received))}" `
         + `fill="var(--s3)" opacity=".85"><title>Received ${esc(bucketLabel(r.k))}: ${esc(rs(r.received))}</title></rect>`;
      h += `<rect x="${cx(i)}" y="${y(r.invoiced)}" width="${bw}" height="${Math.max(1, y(0) - y(r.invoiced))}" `
         + `fill="var(--s2)" opacity=".85"><title>Invoiced ${esc(bucketLabel(r.k))}: ${esc(rs(r.invoiced))}</title></rect>`;
      if (i % step === 0 || i === n - 1)
        h += `<text class="tick" text-anchor="middle" x="${cx(i)}" y="${p.t + ih + 17}">${esc(bucketLabel(r.k))}</text>`;
    });
    h += `<text class="tick" x="${p.l}" y="${p.t - 6}" style="font-weight:700;fill:var(--s3)">■ Received</text>`;
    h += `<text class="tick" x="${p.l + 90}" y="${p.t - 6}" style="font-weight:700;fill:var(--s2)">■ Invoiced</text>`;
    svg.innerHTML = h;
  }

  const q0 = S.search.gr.trim().toLowerCase();
  if (S.grView === 'vend') {
    let vend = buildGrnVendors();
    if (q0) vend = vend.filter((v) => matches(q0, v.supplier, v.name));
    vend.forEach((v) => v.__key = 'grv:' + v.supplier);
    table('gr-table', COLS.grnVendor, slice(vend, S.page.gr), {
      foot: footFor(COLS.grnVendor, vend, 'Total — ' + countOf(vend.length, 'vendor')),
      empty: 'No vendors in this scope.' });
    pager('gr-pager', vend.length, S.page.gr, (n) => { S.page.gr = n; paintGRN(); });
  } else {
    let rws = D.pr.rows;
    if (q0) rws = rws.filter((r) => matches(q0, r.grain_key, r.supplier_name, r.supplier_invoice_no));
    rws.forEach((r) => r.__key = 'pr:' + r.grain_key);
    table('gr-table', COLS.prFull, slice(rws, S.page.gr), {
      foot: footFor(COLS.prFull, rws), empty: 'No goods receipts in this scope.' });
    pager('gr-pager', rws.length, S.page.gr, (n) => { S.page.gr = n; paintGRN(); });
  }
}

function buildGrnVendors() {
  const m = new Map();
  const touch = (sup, name) => {
    let e = m.get(sup);
    if (!e) { e = { supplier: sup, name: name || sup, po: 0, grn: 0, pending: 0,
                    billed: 0, unbilled: 0, ret: 0, earliest: null, oldest: 0 };
              m.set(sup, e); }
    return e;
  };
  (S.D.poDoc || []).forEach((r) => {
    const e = touch(r.supplier, r.supplier_name);
    e.po += num(r.ordered); e.pending += num(r.pending_receipt);
    if (num(r.pending_receipt) > 0.005 && r.required_by
        && (!e.earliest || r.required_by < e.earliest)) e.earliest = r.required_by;
  });
  (S.D.prDoc || []).forEach((r) => {
    const e = touch(r.supplier, r.supplier_name);
    e.grn += num(r.gross); e.billed += num(r.invoiced); e.unbilled += num(r.pending_bill);
    if (IS_RETURN(r)) e.ret += num(r.gross);
    if (num(r.pending_bill) > 0.005 && r.date) {
      const d = daysAgo(r.date, S.asOf); if (d > e.oldest) e.oldest = d;
    }
  });
  return Array.from(m.values()).sort((a, b) => Math.abs(b.grn) - Math.abs(a.grn));
}

/* ══ 17 · Invoices & Payments ════════════════════════════════════════════ */

function paintInvoices() {
  const D = S.D, L = S.L;
  kpi('in-kpis', [
    { k: 'Purchase invoice total', v: rs(D.pi.normalV), raw: D.pi.normalV, tone: 'var(--s2)',
      c: cnt(D.pi.normalN) + ' invoices, before debit notes',
      drill: () => drill('Purchase invoice total', D.pi.normal, COLS.piFull, 'gross', D.pi.normalV) },
    { k: 'Debit notes / purchase returns', v: rs(D.pi.returnV), raw: D.pi.returnV, tone: 'var(--s5)',
      c: cnt(D.pi.returnN) + ' documents, carried negative throughout',
      drill: () => drill('Debit notes and purchase returns', D.pi.returns, COLS.piFull, 'gross', D.pi.returnV) },
    { k: 'Net booked purchase', v: rs(D.pi.gross), raw: D.pi.gross, tone: 'var(--s1)',
      c: countOf(D.pi.n, 'document') + ' · invoices net of returns, all routes',
      drill: () => drill('Net booked purchase', D.pi.rows, COLS.piFull, 'gross', D.pi.gross) },
    { k: 'Taxable value', v: rs(D.pi.taxable), raw: D.pi.taxable, tone: 'var(--s3)',
      c: 'net of tax, at document grain · ' + countOf(D.pi.n, 'document'),
      drill: () => drill('Taxable value', D.pi.rows, COLS.piFull, 'taxable', D.pi.taxable) },
    /* item-level GST does not tie to posted tax, so every statutory figure
       below comes from the PI Tax grain and never from an item field */
    { k: 'GST total', v: rs(D.tax.gst), raw: D.tax.gst, tone: 'var(--s3)',
      c: 'CGST ' + rs(D.tax.cgst) + ' · SGST ' + rs(D.tax.sgst) + ' · IGST ' + rs(D.tax.igst),
      drill: () => drill('GST total', where(D.piTax, (r) => num(r.cgst) || num(r.sgst) || num(r.igst)),
        COLS.tax, null, D.tax.gst, (r) => num(r.cgst) + num(r.sgst) + num(r.igst)) },
    { k: 'TDS total', v: rs(D.tax.tds), raw: D.tax.tds, tone: 'var(--s2)',
      c: cnt(where(D.piTax, (r) => num(r.tds)).length) + ' withholding rows, resolved by account name',
      drill: () => drill('TDS total', where(D.piTax, (r) => num(r.tds)), COLS.tax, 'tds', D.tax.tds) },
    { k: 'Other charges & round-off', v: rs(D.tax.other + D.roundOff), raw: D.tax.other + D.roundOff,
      tone: 'var(--muted)',
      c: 'other tax rows ' + rs(D.tax.other) + ' · round-off ' + rs(D.roundOff),
      drill: () => drill('Other charges', where(D.piTax, (r) => num(r.other_charges)),
        COLS.tax, 'other_charges', D.tax.other) },
    { k: 'Paid against period invoices', v: rs(D.allocated), raw: D.allocated, tone: 'var(--s4)',
      c: countOf(D.peAlloc.length, 'allocation row'),
      drill: () => drill('Allocated payments', D.peAlloc, COLS.pay, 'paid', D.allocated) },
    { k: 'Outstanding at cut-off', v: rs(L.openTotal), raw: L.openTotal, tone: 'var(--s5)',
      c: countOf(L.open.length, 'open invoice') + ' at ' + ddmmmyyyy(S.asOf),
      drill: () => drill('Outstanding at cut-off', L.open, COLS.openItems, 'outstanding', L.openTotal) },
    { k: 'Direct expense / service bills', v: rs(D.pi.directV), raw: D.pi.directV, tone: 'var(--s2)',
      c: pct(Math.abs(D.pi.directV), Math.abs(D.pi.gross)) + ' of booked purchase · no PO, no GRN',
      drill: () => drill('Direct expense / service bills', D.pi.direct, COLS.piFull, 'gross', D.pi.directV) }
  ]);

  const q0 = S.search.inv.trim().toLowerCase();
  let rows = D.pi.rows;
  if (S.src) rows = rows.filter((r) => r.source_class === S.src);
  if (q0) rows = rows.filter((r) => matches(q0, r.grain_key, r.supplier_name, r.supplier_invoice_no));
  rows.forEach((r) => r.__key = 'pi:' + r.grain_key);
  table('in-table', COLS.piView, slice(rows, S.page.inv), {
    foot: footFor(COLS.piView, rows), empty: 'No invoices in this scope.' });
  pager('in-pager', rows.length, S.page.inv, (n) => { S.page.inv = n; paintInvoices(); });

  paintPayments();
  paintTruePurchase();
}

let BANK_REG = [];

function paintPayments() {
  const D = S.D, L = S.L;
  kpi('pa-kpis', [
    { k: 'Total released (period)', v: rs(D.cashOut), raw: D.cashOut, tone: 'var(--s4)',
      c: countOf(D.payN, 'payment entry') + ' · allocated ' + rs(D.allocated)
        + ' · unallocated ' + rs(D.unallocField),
      drill: () => drill('Total released', D.peAlloc.concat(D.peUnalloc), COLS.pay, null, D.cashOut, payVal) },
    { k: 'Allocated to invoices', v: rs(D.allocated), raw: D.allocated, tone: 'var(--s1)',
      c: countOf(D.peAlloc.length, 'Payment Entry Reference row'),
      drill: () => drill('Allocated to invoices', D.peAlloc, COLS.pay, 'paid', D.allocated) },
    { k: 'Unallocated advances', v: rs(D.unallocField), raw: D.unallocField, tone: 'var(--s2)',
      c: countOf(D.peUnalloc.length, 'entry') + ' · the ERP unallocated field, which understates the derived figure',
      drill: () => drill('Unallocated advances', D.peUnalloc, COLS.payAdv, 'advance', D.unallocField) },
    /* only a handful of supplier payments have ever been bank-matched, so this
       is close to the whole payment run: an exception to act on, not a ratio */
    { k: 'Bank-unreconciled allocated cash', v: rs(D.unrecV), raw: D.unrecV, tone: 'var(--s5)',
      c: pct(D.unrecV, D.allocated) + ' of allocated cash carries no clearance date',
      drill: () => drill('Bank-unreconciled', D.unrec, COLS.pay, 'paid', D.unrecV) },
    { k: 'Unapplied cash at cut-off', v: rs(L.unappliedV), raw: L.unappliedV, tone: 'var(--s4)',
      c: countOf(L.unapplied.length, 'entry') + ' · derived, the reconciling figure',
      drill: () => drill('Unapplied cash at cut-off', L.unapplied, COLS.unapplied, 'advance', L.unappliedV) }
  ]);

  /* instrument labels come from the Bank Account master, built server-side.
     Account numbers are never shown; a corporate card is kept separate from
     the bank account it shares a name with, because merging them would hide a
     card-funded payment channel. */
  const split = D.bankAlloc.map((b) => {
    const adv = D.bankAdv.find((a) => a.key === b.key) || { v: 0, n: 0, rows: [] };
    return { key: b.key, alloc: b.v, allocN: b.n, adv: adv.v, advN: adv.n,
      total: b.v + adv.v, rows: b.rows.concat(adv.rows) };
  });
  BANK_REG = split;
  const max = Math.max.apply(null, split.map((s) => Math.abs(s.total))) || 1;
  el('pa-instr').innerHTML =
    '<div class="fk-irow hd"><span>Company instrument</span><span>Share</span><span>Value ₹</span><span>Entries</span></div>'
    + split.map((s, i) => {
        const dead = s.total === 0 && !s.allocN && !s.advN;
        const tone = /card/i.test(s.key) ? 'var(--s5)' : 'var(--s4)';
        return `<div class="fk-irow" data-bank="${i}" role="button" tabindex="0">`
          + `<span class="nm"><span class="dot" style="background:${dead ? 'var(--muted)' : tone}"></span>${esc(s.key)}</span>`
          + `<span class="bar"><i style="width:${(Math.abs(s.total) / max * 100).toFixed(1)}%;background:${tone}"></i></span>`
          + `<span class="n">${esc(rs(s.total))}</span><span class="n">${esc(cnt(s.allocN + s.advN))}</span></div>`;
      }).join('')
    + `<div class="fk-irow hd"><span>Total</span><span></span>`
    + `<span class="n">${esc(rs(split.reduce((a, s) => a + s.total, 0)))}</span>`
    + `<span class="n">${esc(cnt(D.payN))}</span></div>`;

  el('pa-reg').setAttribute('aria-pressed', String(S.payView === 'register'));
  el('pa-det').setAttribute('aria-pressed', String(S.payView !== 'register'));
  let rows, cols, full;
  if (S.payView === 'register') {
    rows = buildPaymentRegister(); cols = COLS.payView; full = COLS.payReg;
    rows.forEach((r) => r.__key = 'per:' + r.pe_id);
  } else {
    rows = D.peAlloc.concat(D.peUnalloc); cols = COLS.pay; full = COLS.payFull;
    rows.forEach((r) => r.__key = 'pe:' + r.grain_key);
  }
  S.payRows = rows; S.payCols = full;
  table('pa-table', cols, slice(rows, S.page.pay), {
    foot: footFor(cols, rows), empty: 'No payments in this scope.' });
  pager('pa-pager', rows.length, S.page.pay, (n) => { S.page.pay = n; paintPayments(); });
}

/* one row per Payment Entry. A payment settling BLR and JDP invoices is split
   by allocation and flagged Mixed; its unallocated balance keeps its own
   classification rather than being forced onto one registration. */
function buildPaymentRegister() {
  const m = new Map();
  const touch = (r) => {
    let e = m.get(r.pe_id);
    if (!e) { e = { pe_id: r.pe_id, date: r.date, supplier: r.supplier, supplier_name: r.supplier_name,
      bank: r.bank, payment_mode: r.payment_mode, reference_no: r.reference_no,
      reference_date: r.reference_date, clearance_date: r.clearance_date,
      recon_state: r.recon_state, created_by: r.created_by, status: r.status,
      allocated: 0, unallocated: 0, paid: 0, inv: new Set(), split: {}, flags: new Set(), rows: [] };
      m.set(r.pe_id, e); }
    return e;
  };
  const tally = (r, field, into) => {
    const e = touch(r), v = num(r[field]);
    e[into] += v; e.paid += v; e.rows.push(r);
    const k = r.ap_location || 'Unassigned';
    e.split[k] = num(e.split[k]) + v;
    String(r.quality_flag || '').split(' · ').forEach((x) => { if (x) e.flags.add(x); });
    return e;
  };
  (S.D.peAlloc || []).forEach((r) => { const e = tally(r, 'paid', 'allocated'); if (r.pi_id) e.inv.add(r.pi_id); });
  (S.D.peUnalloc || []).forEach((r) => tally(r, 'advance', 'unallocated'));

  return Array.from(m.values()).map((e) => {
    const ks = Object.keys(e.split).filter((k) => Math.abs(e.split[k]) >= 0.005);
    e.ap_location = ks.length === 1 ? ks[0] : ks.length ? 'Mixed' : 'Unassigned';
    e.split_text = ks.map((k) => k + ' ' + inr(e.split[k])).join(' · ');
    e.linked_n = e.inv.size;
    e.linked_ids = Array.from(e.inv).join(', ');
    e.registration_state = ks.map((k) => locOf(k).state).filter(Boolean).join(' + ');
    e.resolved_gstin = ks.map((k) => locOf(k).gstin).filter(Boolean).join(' + ');
    e.quality_flag = Array.from(e.flags).join(' · ');
    return e;
  }).sort((a, b) => a.date < b.date ? 1 : -1);
}

function paintTruePurchase() {
  const q0 = S.search.tp.trim().toLowerCase();
  let rows = S.D.truePurchase;
  if (q0) rows = rows.filter((r) => matches(q0, r.supplier, r.name));
  rows.forEach((r) => r.__key = 'tp:' + r.supplier);
  S.tpRows = rows;
  table('tp-table', COLS.truePurchase, slice(rows, S.page.tp), {
    foot: footFor(COLS.truePurchase, rows, 'Total — ' + countOf(rows.length, 'vendor')),
    recon: [reconLine('Net booked, this table vs the invoice register above',
      rows.reduce((a, r) => a + num(r.net), 0), q0 ? null : S.D.pi.gross)],
    empty: 'No vendors in this scope.' });
  pager('tp-pager', rows.length, S.page.tp, (n) => { S.page.tp = n; paintTruePurchase(); });
}

/* ══ 18 · AP Aging ═══════════════════════════════════════════════════════
   Three visible cards and nothing hidden behind them. The statement bridges
   aged bills to the net payable per books and on to the trial balance; the
   strip gives the six-bucket shape; the table is one row per enabled supplier.

   The arithmetic is stated on the face so it can be checked by eye:
     bills + debit notes − advance + payments on account + journals = party ledger
     party ledger (subledger)                                       = GL control
   Switching the basis re-ages client-side and never refetches. */

function paintAgeing() {
  if (!S.A.rows) return;
  ['ag-bill', 'ag-post', 'ag-due'].forEach((k) => el(k).setAttribute('aria-pressed',
    String(S.agBasis === (k === 'ag-bill' ? 'bill' : k === 'ag-post' ? 'posting' : 'due'))));
  el('ag-lvl-sup').setAttribute('aria-pressed', String(S.agLevel === 'supplier'));
  el('ag-lvl-inv').setAttribute('aria-pressed', String(S.agLevel !== 'supplier'));

  paintAgStatement();
  paintAgBuckets();
  paintAgTable();
}

function paintAgStatement() {
  const A = S.A;
  const scope = [S.loc ? 'location ' + locLabel(S.loc) : '',
    S.supplier ? 'supplier ' + (supplierName(S.supplier) || S.supplier) : '',
    S.sg ? 'group ' + S.sg : '', S.cc ? 'cost centre ' + S.cc : '',
    S.dept ? 'department ' + S.dept : ''].filter(Boolean);
  el('ag-stmt-scope').innerHTML = `<b>${esc(AP_ACCOUNT)}</b> as at <b>${esc(ddmmmyyyy(S.asOf))}</b> · `
    + `aged on ${esc(A.basis.toLowerCase())} · ₹ with paise · `
    + (scope.length ? '<b>scoped to ' + esc(scope.join(' and ')) + '</b>' : '<b>company-wide</b>');

  const line = (label, v, memo, cls) =>
    `<div class="${cls || ''}">${label}</div><div class="n ${cls || ''}">${esc(inr(v))}</div>`
    + `<div class="m">${esc(memo || '')}</div>`;
  const tie = (label, got, expect, memo) => {
    if (expect === null || expect === undefined) {
      return `<div>${label}</div><div class="n">${esc(inr(got))}</div><div class="m">${esc(memo || '')}</div>`;
    }
    const d = r2(got - expect), ok = Math.abs(d) <= TOL;
    return `<div>${label}</div><div class="n">${esc(inr(expect))}</div>`
      + `<div class="m ${ok ? 'good' : 'bad'}">${ok ? '✔' : '✖'} Δ ${esc(inr(d))}</div>`;
  };

  el('ag-stmt').innerHTML =
      line('Bills outstanding, aged below', A.billTot, countOf(A.billN, 'bill'))
    + line('Debit notes, their own class', A.dn, countOf(A.dnN, 'document'))
    + line('Less: advance paid, bill booked but not allocated', -A.advHeld,
        (() => { const c = A.rows.filter((o) => o.cls === 'BA').length; return countOf(c, 'supplier'); })(), 'less')
    + line('Less: advance paid, no bill booked', -A.advOnly,
        (() => { const c = A.rows.filter((o) => o.cls === 'A').length; return countOf(c, 'supplier'); })(), 'less')
    + line('Add: payments carrying a payable balance', A.pp, '')
    + line('Add: open journal items', A.jv, '')
    + '<div class="rule"></div>'
    + line('Net payable per books', A.net, '', 'tot')
    /* SELF-TEST, not a control. The master report answers AP Aging and
       AP GL Control from one pass over tabPayment Ledger Entry with no filter
       between them, so the two are the same rows. This line therefore proves
       the decomposition ABOVE is arithmetically complete — a non-zero here is a
       bug on this page, never a finding in the data. It is the line that would
       have caught v28 omitting debit notes and payments on account. */
    + tie('Self-test: the decomposition above vs the same grain, unsplit', A.net, A.led,
        'same rows — must be 0.00')
    + '<div class="rule"></div>'

    /* THE CONTROLS. Different code paths over the same facts: the native report
       is ERPNext's own implementation, and tabGL Entry is a different table
       from the payment ledger this page is built on. */
    + (A.nat === null
        ? '<div>ERPNext Accounts Payable report</div><div class="n">n/a</div>'
          + '<div class="m">the native report did not answer — see the browser console</div>'
        : A.unscoped
          ? tie('ERPNext Accounts Payable report', A.net, A.nat, countOf(A.natRows, 'report row'))
          : '<div>ERPNext Accounts Payable report</div><div class="n">' + esc(inr(A.nat)) + '</div>'
            + '<div class="m">company-wide figure' + (A.tieable
                ? ' · filtered to suppliers here, so each row below still ties'
                : ' · this page is filtered inside suppliers, so neither total nor row ties') + '</div>')
    + (A.vl === null
        ? '<div>Vendor ledgers, ' + esc(AP_ACCOUNT) + '</div><div class="n">n/a</div>'
          + '<div class="m">the GL read did not answer — see the browser console</div>'
        : A.unscoped
          ? tie('Vendor ledgers, ' + AP_ACCOUNT, A.net, A.vl, 'Cr ' + inr(A.cr) + ' · Dr ' + inr(A.dr))
          : '<div>Vendor ledgers, ' + esc(AP_ACCOUNT) + '</div><div class="n">' + esc(inr(A.vl)) + '</div>'
            + '<div class="m">company-wide figure' + (A.tieable
                ? ' · per-supplier ties below are unaffected'
                : ' · per-supplier ties stood down at this scope') + '</div>')
    + (A.tieable
        ? `<div>Suppliers disagreeing with ERPNext&rsquo;s own AP report</div>`
          + `<div class="n ${A.offNat ? 'bad' : 'good'}">${cnt(A.offNat)}</div>`
          + `<div class="m">of ${cnt(A.rows.length)} · tolerance ${esc(inr(TOL))}</div>`
          + `<div>Suppliers disagreeing with their own ledger</div>`
          + `<div class="n ${A.offVl ? 'bad' : 'good'}">${cnt(A.offVl)}</div>`
          + `<div class="m">of ${cnt(A.rows.length)} · tolerance ${esc(inr(TOL))}</div>`
        : `<div>Per-supplier ties</div><div class="n">stood down</div>`
          + `<div class="m">a filter is dropping rows inside suppliers, so the two sides `
          + `describe different populations — clear it to prove the book</div>`)
    + (A.gl === null
        ? '<div>Cross-check: GL control total</div><div class="n">n/a</div>'
          + '<div class="m">not comparable at this scope — GL Entry carries no AP Location, '
          + 'supplier group or department</div>'
        : tie('Cross-check: GL control total', A.vl === null ? A.led : A.vl, A.gl, countOf(A.glN, 'GL row')))
    + tie('Cross-check: journals not linked to an invoice', A.jv, A.jrGl, 'AP GL Journal grain')
    + '<div class="rule"></div>'
    + `<div>Schedule III reclass — Dr balances on trade payables</div>`
    + `<div class="n less">${esc(inr(A.dr))}</div>`
    + `<div class="m">${cnt(A.drN)} ${plural(A.drN, 'supplier')} · present under Other Current Assets, never netted</div>`;
}

function paintAgBuckets() {
  const A = S.A;
  const tot = A.billTot || 1;
  const max = Math.max.apply(null, A.bands.map((v) => Math.abs(v)).concat([1]));
  el('ag-buckets').innerHTML = BANDS.map((b, i) => {
    const v = A.bands[i], w = Math.max(0.3, Math.abs(v) / max * 100);
    return `<button type="button" class="fk-bucket" data-band="${i}">`
      + `<span class="bl">${esc(b.label)}<span>${pct(v, tot)}</span></span>`
      + `<span class="bt"><span class="bf" style="width:${w.toFixed(1)}%;background:${b.tone}"></span></span>`
      + `<span class="bv">${esc(inr(v))} · ${esc(cnt(A.bandN[i]))} ${plural(A.bandN[i], 'bill')}</span></button>`;
  }).join('')
    + `<div class="fk-note">${reconLine('Σ buckets + debit notes vs bills outstanding',
         r2(A.bucketSum + A.dn), A.billTot + A.dn)}</div>`;
}

/* the supplier-level population in display order, ignoring the free-text box so
   an export carries the filtered table rather than whatever is typed */
function agVisible0() {
  const rows = hfPass(S.A.rows, 'ag', COLS.agSupplier);
  return S.sort.ag.key === 'group' ? agGroupSort(rows) : hfSort(rows, 'ag', COLS.agSupplier);
}
/* 'group' is not a column, so it needs its own comparator: the five classes in
   order, and inside each the figure that class is actually about. Without this
   the classes interleave and the section headings fire on the wrong rows. */
function agGroupSort(rows) {
  const weight = (o) => CLASSES[o.cls].group === 0 ? o.billTot
    : CLASSES[o.cls].group === 1 ? o.adv : Math.abs(num(o.led));
  return rows.slice().sort((a, b) => {
    const ga = CLASSES[a.cls].group, gb = CLASSES[b.cls].group;
    if (ga !== gb) return ga - gb;
    const d = weight(b) - weight(a);
    return d !== 0 ? d : String(a.nm).localeCompare(String(b.nm));
  });
}
function agVisible() {
  const q0 = S.search.ag.trim().toLowerCase();
  if (S.agLevel === 'supplier') {
    let rows = S.A.rows;
    if (q0) rows = rows.filter((r) => matches(q0, r.id, r.nm));
    rows = hfPass(rows, 'ag', COLS.agSupplier);
    return S.sort.ag.key === 'group' ? agGroupSort(rows) : hfSort(rows, 'ag', COLS.agSupplier);
  }
  let rows = S.A.bills;
  if (q0) rows = rows.filter((r) => matches(q0, r.sup, r.supName, r.no, r.vno));
  rows = hfPass(rows, 'agb', COLS.agBill);
  return hfSort(rows, 'agb', COLS.agBill);
}

function paintAgTable() {
  const A = S.A;
  const supLevel = S.agLevel === 'supplier';
  const cols = supLevel ? COLS.agSupplier : COLS.agBill;
  const scope = supLevel ? 'ag' : 'agb';
  const rows = agVisible();
  S.agRows = rows;

  const n = hfCount(scope);
  el('ag-hfnote').innerHTML = n
    ? `<b>${cnt(n)} column filter${n > 1 ? 's' : ''}</b> `
      + `<button type="button" class="fk-btn fk-btn-ghost fk-xs" data-hfclear="${scope}">Clear</button>`
    : '';

  if (supLevel) {
    const T = { billTot: 0, dn: 0, adv: 0, allocatable: 0, pp: 0, jv: 0,
                led: 0, net: 0, delta: 0, nat: 0, vl: 0, dNat: 0, dVl: 0 };
    const bandT = BANDS.map(() => 0);
    rows.forEach((r) => {
      BANDS.forEach((b, i) => bandT[i] += r.b[i]);
      Object.keys(T).forEach((k) => T[k] += num(r[k]));
    });
    const foot = { nm: (rows.length === A.rows.length ? 'TOTAL' : 'FILTERED')
      + ' · ' + cnt(rows.length) + ' of ' + countOf(A.rows.length, 'supplier') };
    BANDS.forEach((b, i) => foot[b.k] = rs(bandT[i]));
    ['billTot', 'dn', 'adv', 'allocatable', 'pp', 'jv', 'net', 'nat', 'vl']
      .forEach((k) => foot[k] = rs(T[k]));
    rows.forEach((r) => r.__key = 'ag:' + r.id);

    const shown = slice(rows, S.page.ag);
    let lastGroup = -1;
    const section = (r) => {
      if (S.sort.ag.key !== 'group') return null;
      const g = CLASSES[r.cls].group;
      if (g === lastGroup) return null;
      lastGroup = g;
      return GROUP_NOTE[g] || null;
    };
    table('ag-table', cols, shown, {
      hf: scope, foot, section,
      empty: 'No supplier matches the current scope.',
      recon: [
        /* self-test of this page's own arithmetic, over the rows shown */
        reconLine('Self-test: bills + debit notes − advance + payments on account + journals',
          r2(T.billTot + T.dn - T.adv + T.pp + T.jv), T.net, 'same grain — must be 0.00'),
        !A.tieable
          ? '<span>Per-supplier ties stood down: a filter is dropping rows inside suppliers, '
            + 'so this page and the ERP hold different populations.</span>'
          : A.nat === null ? '<span>ERPNext Accounts Payable report did not answer.</span>'
          : reconLine('This page vs ERPNext Accounts Payable, over the rows shown', T.net, T.nat,
              countOf(rows.filter((r) => r.dNat !== null && Math.abs(r.dNat) > TOL).length, 'supplier') + ' off'),
        !A.tieable ? ''
          : A.vl === null ? '<span>Vendor ledger read did not answer.</span>'
          : reconLine('This page vs the vendors&rsquo; own ledgers, over the rows shown', T.net, T.vl,
              countOf(rows.filter((r) => r.dVl !== null && Math.abs(r.dVl) > TOL).length, 'supplier') + ' off')
      ].filter(Boolean) });
    pager('ag-pager', rows.length, S.page.ag, (n2) => { S.page.ag = n2; paintAgTable(); });
  } else {
    rows.forEach((r) => r.__key = 'agb:' + r.vno);
    table('ag-table', cols, slice(rows, S.page.ag), {
      hf: scope,
      foot: footFor(cols, rows, 'Total · ' + cnt(rows.length) + ' of ' + countOf(A.bills.length, 'bill')),
      empty: 'No bill matches the current scope.',
      recon: [reconLine('This table vs bills outstanding plus debit notes',
        rows.reduce((a, r) => a + num(r.amt), 0),
        rows.length === A.bills.length ? r2(A.billTot + A.dn) : null)] });
    pager('ag-pager', rows.length, S.page.ag, (n2) => { S.page.ag = n2; paintAgTable(); });
  }

  el('ag-note').innerHTML =
    'One row per ' + (supLevel ? 'enabled supplier — every vendor in the master, so a vendor with no '
      + 'payable activity is accounted for rather than silently absent' : 'open bill')
    + '. Party ledger is credit minus debit on ' + esc(AP_ACCOUNT)
    + ': <b>Cr positive = owed</b> to the supplier, <b>Dr negative = net receivable</b>. '
    + 'There is no net payable column — Dr balances are reclassified under Schedule III, not netted. '
    + 'Click a figure to open the native ERP report with the same filter; the General Ledger prints '
    + 'Dr-positive, so a payable reads negative there — same magnitude, opposite sign. '
    + countOf(S.A.nL, 'supplier') + ' sit on the ledger only, ' + cnt(S.A.nNA) + ' have no payable activity.'
    + '<br><b>The two Δ columns are the point of this table.</b> <i>This page</i> is built from the '
    + 'master report, which answers from the payment ledger. <i>Native AP report</i> is ERPNext&rsquo;s '
    + 'own Accounts Payable — a different implementation of the same rules. <i>Vendor ledger</i> is '
    + 'credit minus debit on ' + esc(AP_ACCOUNT) + ' for that party, read straight from the general '
    + 'ledger, which is a different table from the payment ledger and can drift from it silently. '
    + 'Both controls are read company-wide, so a supplier row ties whatever vendors the page is '
    + 'showing. A dash means agreement inside ' + esc(inr(TOL)) + '; a figure is a real disagreement '
    + 'and the row is marked Investigate until it is explained.'
    + (S.A.tieable ? '' : ' <b>Both Δ columns are stood down right now</b> — the active filter drops '
      + 'rows inside suppliers, so this page and the ERP are not looking at the same population.');
}

/* ══ 19 · Vendor Credits ═════════════════════════════════════════════════
   ERPNext has no Vendor Credit doctype, so no single door is the answer. A
   credit arrives through exactly three, and each keeps its own table rather
   than being summed into a figure no ledger would recognise. */

function ensureVCred() {
  if (S.lazy['vcPE|' + S.asOf]) return;
  vcUnallocated().then(() => { if (S.tab === 'vcred') paintVCred(); });
}

function paintVCred() {
  if (!S.L.openAll) return;
  const L = S.L, M = vcModel();
  const q0 = S.search.vc.trim().toLowerCase();
  const keep = (r) => matches(q0, r.supplier_name, r.supplier, r.grain_key, r.pe_id, r.entry_id);
  const dn = M.dn.filter(keep), pe = M.pe.filter(keep), je = M.je.filter(keep);
  S.vcView = { dn, pe, je };

  kpi('vc-kpis', [
    { k: 'Door 1 · return credit open', v: rs(M.dnTot), raw: M.dnTot, tone: 'var(--s5)',
      c: countOf(M.dn.length, 'debit note') + ' with credit still unapplied',
      drill: () => drill('Return credit open', M.dn, COLS.vcDn, 'vc_avail', M.dnTot) },
    { k: 'Door 2 · unapplied cash (derived)', v: rs(L.unappliedV), raw: L.unappliedV, tone: 'var(--s4)',
      c: countOf(L.unapplied.length, 'payment entry') + ' · the reconciling figure',
      drill: () => drill('Unapplied cash', L.unapplied, COLS.unapplied, 'advance', L.unappliedV) },
    { k: 'Door 2 · unapplied cash (ERP field)', v: rs(M.peTot), raw: M.peTot, tone: 'var(--s4)',
      c: countOf(M.pe.length, 'payment entry') + ' · understates the derived figure by '
        + rs(r2(L.unappliedV - M.peTot)),
      drill: () => drill('Unapplied payment documents', M.pe, COLS.vcPe, 'advance', M.peTot) },
    { k: 'Door 3 · journal credits', v: rs(M.jeTot), raw: M.jeTot, tone: 'var(--warn)',
      c: countOf(M.je.length, 'GL row') + ' on the payable account',
      drill: () => drill('Journal credits', M.je, COLS.vcJe, 'outstanding', M.jeTot) },
    { k: 'Credit outside the bill book', v: rs(r2(M.dnTot + L.unappliedV)), raw: r2(M.dnTot + L.unappliedV),
      tone: 'var(--s1)',
      c: 'returns ' + rs(M.dnTot) + ' + unapplied ' + rs(L.unappliedV) }
  ]);

  const paint = (tbl, pg, cols, rows, all, key, total, label) => {
    const col = cols.find((c) => c.k === total);
    rows.forEach((r) => r.__key = key + ':' + (r.grain_key || r.pe_id || r.entry_id));
    table(tbl, cols, slice(rows, S.page[pg]), {
      foot: footFor(cols, rows, 'Total · ' + cnt(rows.length) + ' of ' + countOf(all.length, 'document')),
      recon: [reconLine(label, rows.reduce((a, r) => a + num(cellValue(col, r)), 0),
        rows.length === all.length ? null : all.reduce((a, r) => a + num(cellValue(col, r)), 0),
        rows.length === all.length ? 'whole door' : 'filtered view against the whole door')],
      empty: 'Nothing in this door at the current scope.' });
    pager(pg === 'vcdn' ? 'vc-dn-pager' : pg === 'vcpe' ? 'vc-pe-pager' : 'vc-je-pager',
      rows.length, S.page[pg], (n) => { S.page[pg] = n; paintVCred(); });
  };
  paint('vc-dn-table', 'vcdn', COLS.vcDn, dn, M.dn, 'vcdn', 'vc_avail', 'Credit available in this door');
  if (S.lazy['vcPE|' + S.asOf]) {
    paint('vc-pe-table', 'vcpe', COLS.vcPe, pe, M.pe, 'vcpe', 'advance', 'Unapplied in this door');
  } else {
    el('vc-pe-table').innerHTML = '<tbody><tr><td class="empty">Loading…</td></tr></tbody>';
    el('vc-pe-pager').innerHTML = '';
  }
  paint('vc-je-table', 'vcje', COLS.vcJe, je, M.je, 'vcje', 'outstanding', 'Journal credit in this door');
}

/* ══ 20 · Vendor Balances ════════════════════════════════════════════════ */

function ensureVBal() {
  if (S.lazy[openingKey()]) { paintVBal(); return; }
  vbOpening().then(() => { if (S.tab === 'vbal') paintVBal(); });
}

function paintVBal() {
  if (!S.loaded || !S.lazy[openingKey()]) return;
  const L = S.L, M = vbModel();
  const q0 = S.search.vb.trim().toLowerCase();
  let rows = q0 ? M.filter((r) => matches(q0, r.supplier, r.name)) : M;
  rows = hfSort(rows, 'vb', COLS.vb);
  rows.forEach((r) => r.__key = 'vb:' + r.supplier);
  S.vbRows = rows;

  const T = (arr, k) => r2(arr.reduce((a, r) => a + num(r[k]), 0));
  const live = M.filter((r) => Math.abs(r.closing) > 0.005);
  const cr = live.filter((r) => r.closing > 0), dr = live.filter((r) => r.closing < 0);
  const closed = T(M, 'closing'), opened = T(M, 'opening'), moved = T(M, 'movement');
  const off = M.filter((r) => !r.ties).length;

  kpi('vb-kpis', [
    { k: 'Opening · ' + ddmmmyyyy(dayBefore(S.from)), v: rs(opened), raw: opened, tone: 'var(--muted)',
      c: cnt(M.filter((r) => Math.abs(r.opening) > 0.005).length) + ' vendors · migration JV '
        + rs(T(M, 'openJE')) + ' · pre-period bills ' + rs(T(M, 'openPI'))
        + ' · pre-period payments ' + rs(T(M, 'openPE')),
      drill: () => drill('Opening balance', M.filter((r) => Math.abs(r.opening) > 0.005),
        COLS.vbOpening, 'opening', opened) },
    { k: 'Period movement', v: rs(moved), raw: moved, tone: 'var(--s1)',
      c: 'billed ' + rs(T(M, 'billed')) + ' · debit notes ' + rs(T(M, 'dn'))
        + ' · journals ' + rs(T(M, 'jr')) + ' · payments ' + rs(T(M, 'pay')),
      drill: () => drill('Period movement', M.filter((r) => Math.abs(r.movement) > 0.005),
        COLS.vbFull, 'movement', moved) },
    { k: 'Closing · ' + ddmmmyyyy(S.asOf), v: rs(closed), raw: closed,
      tone: Math.abs(closed - num(L.controlV)) <= TOL ? 'var(--s3)' : 'var(--s5)',
      c: 'AP control ' + rs(L.controlV) + ' · Δ ' + rs(r2(closed - num(L.controlV))),
      drill: () => drill('Closing balance', live, COLS.vbFull, 'closing', closed) },
    { k: 'Payable side', v: rs(T(cr, 'closing')), raw: T(cr, 'closing'), tone: 'var(--s5)',
      c: countOf(cr.length, 'vendor') + ' sit credit',
      drill: () => drill('Payable side', cr, COLS.vb, 'closing', T(cr, 'closing')) },
    { k: 'Advance side', v: rs(T(dr, 'closing')), raw: T(dr, 'closing'), tone: 'var(--s3)',
      c: countOf(dr.length, 'vendor') + ' sit debit · reclassified, never netted',
      drill: () => drill('Advance side', dr, COLS.vb, 'closing', T(dr, 'closing')) },
    { k: 'Unapplied cash', v: rs(L.unappliedV), raw: L.unappliedV, tone: 'var(--s4)',
      c: countOf(L.unapplied.length, 'payment entry') + ' · '
        + cnt(new Set(L.unapplied.map((r) => r.supplier)).size) + ' vendors',
      drill: () => drill('Unapplied cash', L.unapplied, COLS.unapplied, 'advance', L.unappliedV) }
  ]);

  const tied = Math.abs(closed - num(L.controlV)) <= TOL;
  const tie = el('vb-tie');
  tie.className = 'fk-flag ' + (tied ? 'fk-flag-ok' : '');
  tie.innerHTML = (tied ? '✔ ' : '✖ ')
    + `Opening <b>${esc(rs(opened))}</b> + movement <b>${esc(rs(moved))}</b> = closing <b>${esc(rs(closed))}</b>`
    + ` · AP control <b>${esc(rs(L.controlV))}</b> · Δ <b>${esc(rs(r2(closed - num(L.controlV))))}</b>`
    + ` · tolerance ${esc(inr(TOL))} · rows off: <b>${cnt(off)}</b>`
    + ` · opening taken at <b>${esc(ddmmmyyyy(dayBefore(S.from)))}</b>, the day before the period starts`;

  table('vb-table', COLS.vb, slice(rows, S.page.vb), {
    hf: 'vb',
    foot: footFor(COLS.vb, rows, 'Total · ' + cnt(rows.length) + ' of ' + countOf(M.length, 'vendor')),
    recon: [reconLine('Opening + movement vs closing, summed across this table',
      r2(T(rows, 'opening') + T(rows, 'movement')), T(rows, 'closing'))],
    empty: 'No vendor has ledger activity in this scope.' });
  pager('vb-pager', rows.length, S.page.vb, (n) => { S.page.vb = n; paintVBal(); });
}

/* ══ 21 · Categorise AP Location ═════════════════════════════════════════
   queue = Unassigned, and pattern-inferred on request: a manual classification
   applies. flags = already classified by the GST registration on the document,
   but the cost centre disagrees — that is fixed in ERP, not here. Nothing in
   this panel ever writes to a source document. */

function clsRows(kind) {
  const all = [].concat(S.D.poDoc || [], S.D.prDoc || [], S.D.piDoc || [],
    S.D.peAlloc || [], S.D.peUnalloc || [], S.L.openAll || [], S.L.journals || [],
    S.L.unapplied || [], S.L.control || []);
  const m = new Map();
  all.forEach((r) => {
    const st = r.classification_status || '', auto = r.automatic_location || '';
    const inferred = st.indexOf('Inferred') === 0, flag = r.conflict_details || '';
    if (kind === 'flags' ? !flag : !(auto === 'Unassigned' || inferred)) return;
    if (!r.classification_doctype || !r.classification_docname) return;
    const key = r.classification_doctype + '|' + r.classification_docname;
    let e = m.get(key);
    if (!e) { e = { key, dt: r.classification_doctype, dn: r.classification_docname,
      supplier: r.supplier, supplier_name: r.supplier_name, date: r.date, amount: 0,
      company_gstin: r.company_gstin, company_address: r.company_address,
      leaf: r.leaf_cost_centre, anchors: r.resolved_location_cost_centre,
      owner: r.classification_owner, auto, status: st, reason: r.classification_reason,
      flag, override: r.manual_override_reference, loc: r.ap_location, inferred };
      m.set(key, e); }
    const v = r.row_type === 'PI Doc' ? num(r.gross)
      : (r.row_type === 'AP Unapplied' || r.row_type === 'PE Unallocated') ? num(r.advance)
      : num(r.outstanding) || num(r.gross) || num(r.paid);
    if (Math.abs(v) > Math.abs(e.amount)) e.amount = v;
  });
  return Array.from(m.values()).sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
}

function clsPendingCount() {
  S.clsPending = S.loaded && CAN_CLASSIFY()
    ? clsRows('queue').filter((x) => !x.override && !x.inferred).length : 0;
  return S.clsPending;
}

function clsOpen() {
  if (!CAN_CLASSIFY() || !S.loaded) return;
  const queue = clsRows('queue'), flags = clsRows('flags');
  S.cls = { all: { queue, flags }, rows: queue.concat(flags), view: 'queue', proposed: {}, inferred: false };
  S.clsSel = new Set(); S.page.cls = 1; S.hf.cls = {};
  const qn = el('cls-q'); if (qn) qn.value = '';
  S.search.cls = '';
  paintCls();
}

function paintCls() {
  const c = S.cls; if (!c) return;
  const pend = c.all.queue.filter((r) => !r.override && !r.inferred).length;
  const infN = c.all.queue.filter((r) => r.inferred && !r.override).length;
  el('cls-tabs').innerHTML =
    `<button type="button" data-cls="view:queue" aria-pressed="${c.view === 'queue'}">To classify · ${cnt(pend)}</button>`
    + `<button type="button" data-cls="view:flags" aria-pressed="${c.view === 'flags'}" `
    + `title="Classified by the GST registration on the document, but the cost centre points the other way. Correct the cost centre or the GSTIN in ERP.">`
    + `GSTIN vs cost centre · ${cnt(c.all.flags.length)}</button>`
    + `<button type="button" data-cls="view:history" aria-pressed="${c.view === 'history'}">Manual classifications</button>`
    + (c.view === 'queue'
        ? `<button type="button" data-cls="inferred" aria-pressed="${!!c.inferred}" `
          + `title="Documents the report placed by a pattern (linked document, naming series, supplier history, creator). They can be confirmed or corrected here.">`
          + `Include pattern-inferred · ${cnt(infN)}</button>` : '');

  el('cls-queue').hidden = c.view === 'history';
  el('cls-history').hidden = c.view !== 'history';
  if (c.view === 'history') return paintClsHistory();

  const q0 = (S.search.cls || '').trim().toLowerCase();
  let rows = c.view === 'flags' ? c.all.flags
    : c.all.queue.filter((r) => !r.override && (c.inferred || !r.inferred));
  if (q0) rows = rows.filter((r) => matches(q0, r.dn, r.supplier_name, r.flag, r.reason));
  rows = hfSort(hfPass(rows, 'cls', COLS.cls), 'cls', COLS.cls);
  rows.forEach((r) => r.__key = 'cls:' + r.key);
  c.visible = rows;

  const selRows = c.rows.filter((r) => S.clsSel.has(r.key));
  const selV = selRows.reduce((a, r) => a + Math.abs(r.amount), 0);
  const n = hfCount('cls');
  const clear = n ? ` · ${cnt(n)} column filter${n > 1 ? 's' : ''} `
    + `<button type="button" class="fk-btn fk-btn-ghost fk-xs" data-hfclear="cls">Clear</button>` : '';
  el('cls-preview').innerHTML = c.view === 'flags'
    ? `<b>${cnt(rows.length)}</b> documents · value <b>${esc(rs(rows.reduce((a, r) => a + Math.abs(r.amount), 0)))}</b>`
      + ' · already classified by their GST registration; the cost centre disagrees and needs correcting in ERP' + clear
    : `<b>${cnt(selRows.length)}</b> selected · value <b>${esc(rs(selV))}</b>`
      + (selRows.length ? ' · ' + ['BLR', 'JDP'].map((l) =>
          l + ' ' + cnt(selRows.filter((r) => c.proposed[r.key] === l).length)).join(' · ')
        + ' · no location ' + cnt(selRows.filter((r) => !c.proposed[r.key]).length) : '')
      + ' · showing ' + cnt(rows.length) + clear;

  const cols = c.view === 'flags' ? COLS.cls.filter((x) => x.k !== 'sel' && x.k !== 'prop') : COLS.cls;
  const foot = { amount: rs(rows.reduce((a, r) => a + num(r.amount), 0)) };
  foot[cols[0].k] = 'Total — ' + countOf(rows.length, 'document');
  table('cls-table', cols, slice(rows, S.page.cls), { hf: 'cls', foot,
    empty: c.view === 'flags' ? 'No document disagrees with its cost centre in this scope.'
      : 'Nothing left to classify in the loaded scope.' });
  pager('cls-pager', rows.length, S.page.cls, (n2) => { S.page.cls = n2; paintCls(); });
}

async function paintClsHistory() {
  el('cls-htable').innerHTML = '<tbody><tr><td class="empty">Loading…</td></tr></tbody>';
  const rows = await frappe.call({ method: 'frappe.client.get_list', args: {
    doctype: OVERRIDE_DT,
    fields: ['name', 'source_doctype', 'source_name', 'automatic_result', 'ap_location', 'gstin',
      'reason', 'status', 'reversal_reason', 'owner', 'creation', 'modified_by', 'modified'],
    order_by: 'modified desc', limit_page_length: 500 } })
    .then((r) => r.message || []).catch(() => []);
  const head = ['Override', 'Source', 'Automatic', 'AP Location', 'GSTIN', 'Reason',
    'Status', 'Created by', 'Created', 'Modified by', ''];
  el('cls-htable').innerHTML = '<thead><tr>' + head.map((h) => `<th>${esc(h)}</th>`).join('') + '</tr></thead><tbody>'
    + (rows.length ? rows.map((o) =>
        `<tr><td class="txt">${docLink(OVERRIDE_DT, o.name)}</td>`
        + `<td class="txt">${esc(o.source_doctype)} ${docLink(o.source_doctype, o.source_name)}</td>`
        + `<td class="txt">${locChip(o.automatic_result)}</td><td class="txt">${locChip(o.ap_location)}</td>`
        + `<td class="txt">${esc(o.gstin || '')}</td>`
        + `<td class="txt"><small>${esc(o.reason || '')}`
        + `${o.reversal_reason ? '<br>Reversed: ' + esc(o.reversal_reason) : ''}</small></td>`
        + `<td class="txt">${o.status === 'Active' ? chip('✔ Active', 'c-ok') : chip('↩ Reversed', 'c-neu')}</td>`
        + `<td class="txt">${esc(o.owner)}</td><td class="txt">${esc(String(o.creation).slice(0, 16))}</td>`
        + `<td class="txt">${esc(o.modified_by)}</td>`
        + `<td class="txt">${o.status === 'Active'
            ? `<button type="button" class="fk-btn fk-btn-ghost fk-xs" data-cls="reverse:${esc(o.name)}">Reverse</button>` : ''}</td></tr>`).join('')
      : `<tr><td class="empty" colspan="${head.length}">No manual classifications yet.</td></tr>`)
    + '</tbody>';
}

async function clsAction(a, node) {
  const c = S.cls; if (!c) return;
  if (a.indexOf('view:') === 0) { c.view = a.slice(5); S.page.cls = 1; return paintCls(); }
  if (a === 'inferred') { c.inferred = !c.inferred; S.page.cls = 1; return paintCls(); }
  if (a === 'allf') { (c.visible || []).forEach((r) => S.clsSel.add(r.key)); return paintCls(); }
  if (a === 'clearsel') { S.clsSel = new Set(); return paintCls(); }
  if (a.indexOf('row:') === 0) {
    const k = a.slice(4);
    if (node.checked) S.clsSel.add(k); else S.clsSel.delete(k);
    return paintCls();
  }
  if (a === 'blr' || a === 'jdp') {
    S.clsSel.forEach((k) => c.proposed[k] = a.toUpperCase());
    return paintCls();
  }
  if (a === 'apply') {
    const reason = el('cls-reason').value.trim();
    const sel = c.rows.filter((r) => S.clsSel.has(r.key) && c.proposed[r.key]);
    if (!sel.length) return status('Select rows and assign BLR or JDP first.', 'err');
    if (reason.length < 10) return status('A classification reason of at least 10 characters is mandatory.', 'err');
    const value = sel.reduce((x, r) => x + Math.abs(r.amount), 0);
    const yes = await new Promise((res) => frappe.confirm(
      `Create <b>${sel.length}</b> manual AP Location classification(s) covering <b>${esc(rs(value))}</b>?`
      + '<br><br>Source documents are not changed. Each record can be reversed.',
      () => res(true), () => res(false)));
    if (!yes) return;
    let ok = 0; const fail = [];
    for (const r of sel) {
      const l = locOf(c.proposed[r.key]);
      try {
        await frappe.call({ method: 'frappe.client.insert', args: { doc: {
          doctype: OVERRIDE_DT, source_doctype: r.dt, source_name: r.dn,
          automatic_result: r.auto, ap_location: l.k, registration_state: l.state,
          gstin: l.gstin, status: 'Active', reason,
          supplier: r.supplier_name || r.supplier || '', amount: r.amount } } });
        ok++;
      } catch (e) { console.error('[AP] classify ' + r.dn, e); fail.push(r.dn); }
    }
    status(`${ok} classified${fail.length ? ' · failed: ' + fail.join(', ') : ''}. Reloading…`,
      fail.length ? 'err' : 'busy');
    S.cls = null; CACHE.clear();
    await reload(true);
    return;
  }
  if (a.indexOf('reverse:') === 0) {
    const name = a.slice(8);
    const why = await new Promise((res) => frappe.prompt(
      { fieldname: 'why', fieldtype: 'Small Text', label: 'Reason for reversing ' + name, reqd: 1 },
      (v) => res(v.why || ''), 'Reverse manual classification', 'Reverse'));
    if (!why || why.trim().length < 5) return status('Reversal needs a reason of at least 5 characters.', 'err');
    try {
      await frappe.call({ method: 'frappe.client.set_value', args: {
        doctype: OVERRIDE_DT, name, fieldname: { status: 'Reversed', reversal_reason: why.trim() } } });
      status('Reversed ' + name + '. Reloading…', 'busy');
      CACHE.clear();
      await reload(true);
      paintClsHistory();
    } catch (e) { console.error('[AP] reverse ' + name, e); status('Could not reverse ' + name, 'err'); }
  }
}

/* ══ 22 · exports ════════════════════════════════════════════════════════ */

const stripCols = (cols) => cols.filter((c) => c.k !== 'sel' && c.k !== 'prop' && c.k !== 'door')
  .map((c) => ({ label: c.label, k: c.k, raw: c.raw, plain: c.plain }));

/* the one workbook. Sheet 1 is the merged supplier table under the statement
   bridge, sheet 2 every open bill, sheet 3 the per-supplier tie exploded.
   It carries whatever view is on screen: the basis pills change the ageing,
   and the column filters narrow the population. */
function agWorkbook() {
  const A = S.A;
  const T = (v) => [v, 't'], B = (v) => [v, 'b'], N = (v) => [v, 'n'],
        NB = (v) => [v, 'N'], H = (v) => [v, 'h'], TT = (v) => [v, 'T'], I = (v) => [v, 'i'];
  const supRows = agVisible0();
  const billRows = hfSort(hfPass(A.bills, 'agb', COLS.agBill), 'agb', COLS.agBill);

  const R = [];
  R.push([TT('FURNISHKA TECH PRIVATE LIMITED')]);
  R.push([T('AP Aging by ' + A.basis + ' as at ' + S.asOf), null, null, null, null, T(AP_ACCOUNT)]);
  R.push([T('Scope: ' + ([S.loc && locLabel(S.loc), S.supplier, S.sg, S.cc, S.dept]
    .filter(Boolean).join(' · ') || 'company-wide'))]);
  R.push([T('Sign: credit minus debit. Cr positive = owed. Dr negative = net receivable, '
    + 'reclassified per Schedule III and never netted.')]);
  R.push([T('The native General Ledger prints Dr-positive, so a payable reads negative there '
    + '— same magnitude, opposite sign.')]);
  R.push([]);
  R.push([B('Bills outstanding, aged'), null, null, null, null, NB(A.billTot),
    T(A.billN + ' ' + plural(A.billN, 'bill'))]);
  R.push([B('Debit notes, their own class'), null, null, null, null, NB(A.dn),
    T(A.dnN + ' ' + plural(A.dnN, 'document'))]);
  R.push([B('Less: advance paid, bill booked but not allocated'), null, null, null, null, NB(-A.advHeld)]);
  R.push([B('Less: advance paid, no bill booked'), null, null, null, null, NB(-A.advOnly)]);
  R.push([B('Add: payments carrying a payable balance'), null, null, null, null, NB(A.pp)]);
  R.push([B('Add: open journal items'), null, null, null, null, NB(A.jv)]);
  R.push([B('Net payable per books'), null, null, null, null, NB(A.net)]);
  R.push([B('Party ledger, subledger'), null, null, null, null, NB(A.led),
    T('difference ' + inr(r2(A.net - A.led)))]);
  R.push([B('General ledger control'), null, null, null, null,
    A.gl === null ? T('not comparable at this scope') : NB(A.gl),
    A.gl === null ? null : T('difference ' + inr(r2(A.led - A.gl)))]);
  R.push([B('Suppliers off the identity by more than ' + inr(TOL)), null, null, null, null, I(A.off),
    T('of ' + A.rows.length + ' ' + plural(A.rows.length, 'supplier'))]);
  R.push([B('Schedule III reclass, Dr balances'), null, null, null, null, NB(A.dr),
    T(A.drN + ' ' + plural(A.drN, 'supplier'))]);
  R.push([]);
  R.push([B('Supplier table' + (supRows.length === A.rows.length ? '' : ' (filtered)'))]);
  const h1 = R.length + 1;
  const supCols = stripCols(COLS.agSupplierFull);
  R.push(supCols.map((c) => H(c.label)));
  supRows.forEach((r) => R.push(supCols.map((c) => {
    const v = cellValue(c, r);
    return v === null || v === undefined || v === '' ? null : (typeof v === 'number' ? N(v) : T(v));
  })));
  const e1 = R.length;
  const tot = {};
  supCols.forEach((c) => { if (c.raw) tot[c.k] = r2(supRows.reduce((a, r) => a + num(c.raw(r)), 0)); });
  R.push(supCols.map((c, i) => i === 0
    ? B('TOTAL · ' + supRows.length + ' suppliers')
    : (tot[c.k] !== undefined ? NB(tot[c.k]) : null)));

  const billCols = stripCols(COLS.agBill);
  const S2 = [billCols.map((c) => H(c.label))];
  billRows.forEach((r) => S2.push(billCols.map((c) => {
    const v = cellValue(c, r);
    return v === null || v === undefined || v === '' ? null : (typeof v === 'number' ? N(v) : T(v));
  })));
  S2.push(billCols.map((c, i) => i === 0 ? B('TOTAL · ' + billRows.length + ' bills')
    : (c.k === 'amt' ? NB(r2(billRows.reduce((a, r) => a + num(r.amt), 0))) : null)));

  const S3 = [['vendor_name', 'supplier_id', 'class', 'bills', 'debit_notes', 'advance',
    'payments_on_account', 'journals', 'net_per_report', 'party_ledger', 'difference']
    .map((x) => H(x))];
  let t3 = { bills: 0, dn: 0, adv: 0, pp: 0, jv: 0, net: 0, led: 0, d: 0 };
  supRows.forEach((r) => {
    S3.push([T(r.nm), T(r.id), T(CLASSES[r.cls].label), N(r.billTot), N(r.dn), N(-r.adv),
      N(r.pp), N(r.jv), N(r.net), r.led === null ? null : N(r.led),
      r.delta === null ? null : N(r.delta)]);
    t3.bills += r.billTot; t3.dn += r.dn; t3.adv += r.adv; t3.pp += r.pp;
    t3.jv += r.jv; t3.net += r.net; t3.led += num(r.led); t3.d += num(r.delta);
  });
  S3.push([B('TOTAL'), null, null, NB(t3.bills), NB(t3.dn), NB(-t3.adv), NB(t3.pp),
    NB(t3.jv), NB(t3.net), NB(t3.led), NB(r2(t3.d))]);

  xlsxBook([
    { name: 'ap_aging_summary', rows: R, sel: 1, freeze: h1,
      af: 'A' + h1 + ':' + xlsxCol(supCols.length - 1) + e1,
      widths: supCols.map((c, i) => i === 0 ? 46 : Math.max(14, Math.min(28, c.label.length + 4))) },
    { name: 'bill_wise', rows: S2, freeze: 1,
      af: 'A1:' + xlsxCol(billCols.length - 1) + (S2.length - 1),
      widths: [44, 24, 24, 15, 15, 9, 14, 18, 16] },
    { name: 'supplier_recon', rows: S3, freeze: 1, af: 'A1:K' + (S3.length - 1),
      widths: [46, 20, 18, 18, 16, 16, 20, 16, 18, 18, 16] }
  ], 'AP_Aging_' + (S.agLevel === 'supplier' ? 'Supplier' : 'Bill_Wise')
    + '_By_' + A.basis.replace(/\s+/g, '_') + '_as_at_' + S.asOf);
}

const EXPORTS = {
  'po-export': () => csvExport('Purchase orders', stripCols(COLS.poFull), S.D.po.rows),
  'gr-export': () => S.grView === 'vend'
    ? csvExport('GRN vendor summary', stripCols(COLS.grnVendor), buildGrnVendors())
    : csvExport('Goods receipts', stripCols(COLS.prFull), S.D.pr.rows),
  'in-export': () => csvExport('Invoice register', stripCols(COLS.piFull),
    S.src ? S.D.pi.rows.filter((r) => r.source_class === S.src) : S.D.pi.rows),
  'pa-export': () => csvExport(S.payView === 'register' ? 'Payment register' : 'Payment allocations',
    stripCols(S.payCols), S.payRows || []),
  'tp-export': () => csvExport('Supplier true purchase', stripCols(COLS.truePurchase), S.tpRows || S.D.truePurchase),
  'vb-export': () => csvExport('Vendor balance summary', stripCols(COLS.vbFull), S.vbRows || vbModel()),
  'vc-dn-export': () => csvExport('Vendor credit — purchase returns', stripCols(COLS.vcDn), (S.vcView || {}).dn || []),
  'vc-pe-export': () => csvExport('Vendor credit — unapplied payments', stripCols(COLS.vcPe), (S.vcView || {}).pe || []),
  'vc-je-export': () => csvExport('Vendor credit — journal credits', stripCols(COLS.vcJe), (S.vcView || {}).je || []),
  'ag-xlsx': () => agWorkbook(),
  'ag-csv': () => { const A = S.A;
    csvExport('AP aging by supplier, ' + A.basis.toLowerCase(), stripCols(COLS.agSupplierFull),
      agVisible0(),
      [['Σ buckets', A.bucketSum], ['Debit notes', A.dn],
       ['Bills outstanding', r2(A.bucketSum + A.dn)],
       ['Less advance', -A.adv], ['Add payments on account', A.pp], ['Add journals', A.jv],
       ['Net payable per books', A.net], ['Party ledger', A.led],
       ['General ledger control', A.gl === null ? 'not comparable at this scope' : A.gl],
       ['Suppliers off the identity', A.off], ['Schedule III Dr reclass', A.dr]]); },
  'ag-bill-csv': () => csvExport('AP aging bill-wise, ' + S.A.basis.toLowerCase(),
    stripCols(COLS.agBill), hfSort(hfPass(S.A.bills, 'agb', COLS.agBill), 'agb', COLS.agBill),
    BANDS.map((b, i) => [b.label, S.A.bands[i]])
      .concat([['Debit notes', S.A.dn], ['Bills outstanding', r2(S.A.bucketSum + S.A.dn)]])),
  'dr-export': () => { const d = S.drawer; if (d) csvExport(d.title, stripCols(d.cols), d.rows); }
};

/* ══ 23 · dimensions, from the report itself ═════════════════════════════
   Option lists accumulate across loads, so picking one value never shrinks the
   other lists — a filtered load would otherwise hide the very options needed
   to widen the scope again. */

function loadDimensions() {
  const all = [].concat(S.D.poDoc || [], S.D.prDoc || [], S.D.piDoc || [],
    S.D.peAlloc || [], S.D.peUnalloc || [], S.L.control || []);
  const addTo = (set, fn) => all.forEach((r) => { const v = fn(r); if (v && v !== '—') set.add(v); });
  addTo(S.dims.ccs, (r) => r.cost_center);
  addTo(S.dims.groups, (r) => r.supplier_group);
  addTo(S.dims.depts, (r) => r.department);
  addTo(S.dims.banks, (r) => r.bank);
  all.forEach((r) => { if (r.supplier) S.dims.suppliers.set(r.supplier, r.supplier_name || r.supplier); });

  const fill = (key, set, label, cur) => {
    const s = el(key); if (!s) return;
    const vals = Array.from(set).sort();
    if (cur && vals.indexOf(cur) < 0) vals.unshift(cur);
    s.innerHTML = `<option value="">${esc(label)}</option>`
      + vals.map((v) => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(v)}</option>`).join('');
  };
  const ls = el('f-loc');
  if (ls) ls.innerHTML = '<option value="">Combined</option>'
    + LOCS.map((l) => `<option value="${l.k}"${l.k === S.loc ? ' selected' : ''}>${esc(l.label)}</option>`).join('');
  fill('f-cc',   S.dims.ccs,    'All cost centres', S.cc);
  fill('f-sg',   S.dims.groups, 'All groups',       S.sg);
  fill('f-dept', S.dims.depts,  'All departments',  S.dept);

  /* the roster is the better supplier list: it includes a vendor with no
     document in the period, which is exactly who you go looking for */
  const list = S.roster && S.roster.size
    ? Array.from(S.roster.entries())
    : Array.from(S.dims.suppliers.entries());
  const ss = el('f-supplier');
  if (ss) ss.innerHTML = '<option value="">All suppliers</option>'
    + list.sort((a, b) => String(a[1]).localeCompare(String(b[1])))
      .map(([k, v]) => `<option value="${esc(k)}"${k === S.supplier ? ' selected' : ''}>${esc(v)}</option>`).join('');

  typeSelects();
}

/* FurnishkaType's own select control. The native <select> stays in the DOM and
   remains the single source of truth: every .value read and every 'change'
   listener below is untouched. Called explicitly after the options are written,
   so no MutationObserver is needed to notice them. */
const TYPE_SEL = new Map();
function typeSelects() {
  const FKT = window.FurnishkaType;
  if (!FKT || !FKT.select) return;
  const sig = (s) => s.value + '\u0000' + Array.prototype.map.call(s.options,
    (o) => o.value + '\u0001' + o.textContent).join('\u0002');
  ROOT.querySelectorAll('select[data-fk]').forEach((sel) => {
    const sg = sig(sel), rec = TYPE_SEL.get(sel);
    if (rec && rec.sig === sg) return;
    const opts = Array.prototype.map.call(sel.options,
      (o) => ({ value: o.value, name: o.textContent.trim() }));
    if (rec) { rec.ctl.setOptions(opts); rec.ctl.set(sel.value); rec.sig = sg; return; }
    const allV = (opts.length && /^(all|combined|any)\b/i.test(opts[0].name)) ? opts[0].value : undefined;
    const ctl = FKT.select({
      label: '', value: sel.value, allValue: allV, searchThreshold: 8, options: opts,
      width: Math.max(120, Math.round(sel.getBoundingClientRect().width) || 160),
      onChange: (v) => { if (sel.value === v) return;
        sel.value = v; sel.dispatchEvent(new Event('change', { bubbles: true })); }
    });
    ctl.el.setAttribute('data-fk-type-select', sel.getAttribute('data-fk') || '');
    sel.style.display = 'none';
    sel.insertAdjacentElement('afterend', ctl.el);
    TYPE_SEL.set(sel, { ctl, sig: sg });
  });
}

/* ══ 24 · browser history · every view is revertable ═════════════════════ */

let NAV_SKIP = false;
const NAV_KEYS = ['tab', 'loc', 'cc', 'supplier', 'sg', 'dept', 'src', 'recon',
  'from', 'to', 'asOf', 'preset', 'agBasis', 'agLevel', 'payView', 'grView', 'coView', 'grain'];
const RELOAD_KEYS = ['loc', 'cc', 'supplier', 'sg', 'dept', 'recon', 'from', 'to', 'asOf'];

function navSnap() {
  const o = {};
  NAV_KEYS.forEach((k) => o[k] = S[k]);
  o.hf = {};
  Object.keys(S.hf).forEach((scope) => {
    o.hf[scope] = {};
    Object.keys(S.hf[scope] || {}).forEach((k) => {
      const f = S.hf[scope][k];
      o.hf[scope][k] = f.type === 'set'
        ? { type: 'set', values: Array.from(f.values) }
        : Object.assign({}, f);
    });
  });
  o.sort = JSON.parse(JSON.stringify(S.sort));
  o.page = Object.assign({}, S.page);
  o.search = Object.assign({}, S.search);
  o.scrollY = scroller().scrollTop || 0;
  return o;
}
const pushNav = () => { if (S.loaded) history.pushState({ fk: 'view', v: navSnap() }, ''); };
/* write the exact current view into the CURRENT entry, so Back from an ERP
   report returns to it rather than to whatever was there before */
function saveView() {
  const st = Object.assign({}, history.state || {});
  st.v = navSnap();
  if (!st.fk) st.fk = 'view';
  history.replaceState(st, '');
}

function applySnap(st) {
  NAV_KEYS.forEach((k) => { if (st[k] !== undefined) S[k] = st[k]; });
  S.hf = {};
  Object.keys(st.hf || {}).forEach((scope) => {
    S.hf[scope] = {};
    Object.keys(st.hf[scope] || {}).forEach((k) => {
      const f = st.hf[scope][k];
      S.hf[scope][k] = f.type === 'set' ? { type: 'set', values: new Set(f.values) } : Object.assign({}, f);
    });
  });
  if (st.sort) S.sort = st.sort;
  if (st.page) Object.assign(S.page, st.page);
  if (st.search) {
    Object.assign(S.search, st.search);
    [['po-search', 'po'], ['gr-search', 'gr'], ['in-search', 'inv'], ['tp-search', 'tp'],
     ['ag-search', 'ag'], ['vb-search', 'vb'], ['vc-search', 'vc'], ['cls-q', 'cls']]
      .forEach(([k, key]) => { const n = el(k); if (n) n.value = S.search[key] || ''; });
  }
  el('from').value = S.from; el('to').value = S.to; el('asofin').value = S.asOf;
  paintPresets();
  Object.keys(FILTER_MAP).forEach((k) => {
    const n = el(k); if (n && FILTER_MAP[k] !== 'grain') n.value = S[FILTER_MAP[k]] || '';
  });
  const gn = el('f-grain'); if (gn) gn.value = S.grain;
  typeSelects();
  ['gr-docs', 'gr-vend'].forEach((k) =>
    el(k).setAttribute('aria-pressed', String((k === 'gr-docs') === (S.grView !== 'vend'))));
  ['co-cohort', 'co-activity'].forEach((k) =>
    el(k).setAttribute('aria-pressed', String((k === 'co-cohort') === (S.coView !== 'activity'))));
  ['pa-reg', 'pa-det'].forEach((k) =>
    el(k).setAttribute('aria-pressed', String((k === 'pa-reg') === (S.payView === 'register'))));
  ['tl-abs', 'tl-cum'].forEach((k) =>
    el(k).setAttribute('aria-pressed', String((k === 'tl-abs') === (S.tlMode !== 'cum'))));
}

/* the desk scrolls inside .layout-main-section-wrapper, not the window */
const scroller = () => document.querySelector('.layout-main-section-wrapper') || document.scrollingElement;
const restoreScroll = (y) => { if (y == null) return;
  [80, 450, 1200, 2500].forEach((ms) => setTimeout(() => { scroller().scrollTop = y; }, ms)); };

window.addEventListener('popstate', (e) => {
  if (!ROOT.isConnected) return;
  if (NAV_SKIP) { NAV_SKIP = false; return; }
  const away = S.away; S.away = false;
  if (S.drawer && !away) { drawerBack(true); return; }
  const st = e.state && e.state.v ? e.state.v : null;
  if (!st) return;
  const needReload = RELOAD_KEYS.some((k) => st[k] !== S[k]);
  applySnap(st);
  if (needReload) reload(true).then(() => restoreScroll(st.scrollY));
  else {
    deriveAgeing(); paintAll();
    if (S.drawer) { el('drawer').hidden = false; fitDrawer(); paintDrawer(); }
    restoreScroll(st.scrollY);
  }
});

/* ERP's own reports, opened in this tab. The current view is written into
   history first, so Back returns to exactly what was on screen. */
function erpRoute(fn) { saveView(); S.away = true; fn(); }
function openReconciliation(supplier) {
  erpRoute(() => {
    frappe.route_options = { company: COMPANY, party_type: 'Supplier',
      receivable_payable_account: AP_ACCOUNT,
      party: supplier || undefined,
      to_invoice_date: S.asOf, to_payment_date: S.asOf };
    frappe.new_doc('Payment Reconciliation');
  });
}

/* ══ 25 · events ═════════════════════════════════════════════════════════ */

el('presets').addEventListener('click', (e) => {
  const b = e.target.closest('[data-p]'); if (!b) return;
  applyPreset(b.dataset.p); reload();
});
el('apply').addEventListener('click', () => {
  const f = el('from').value, t = el('to').value, a = el('asofin').value || t;
  if (!f || !t) return status('Pick both dates.', 'err');
  if (f > t) return status('From date is after To date.', 'err');
  if (a < t) return status('The ledger cut-off cannot be before the period end.', 'err');
  S.from = f; S.to = t; S.asOf = a; S.preset = 'custom';
  paintPresets(); reload();
});
el('reset').addEventListener('click', () => { applyPreset('fy'); reload(); });
el('refresh').addEventListener('click', () => { CACHE.clear(); S.lazy = {}; reload(); });

const FILTER_MAP = { 'f-loc': 'loc', 'f-cc': 'cc', 'f-supplier': 'supplier', 'f-sg': 'sg',
  'f-dept': 'dept', 'f-src': 'src', 'f-recon': 'recon', 'f-grain': 'grain' };
Object.keys(FILTER_MAP).forEach((k) => {
  const node = el(k); if (!node) return;
  node.addEventListener('change', (e) => {
    const field = FILTER_MAP[k];
    S[field] = e.target.value;
    Object.keys(S.page).forEach((p) => S.page[p] = 1);
    if (field === 'grain') {
      S.D.series = buildSeries(); S.D.cohort = buildCohort();
      paintTimeline(); paintPO(); paintGRN();
    } else if (field === 'src') {
      pushNav(); paintInvoices(); paintScope(); paintTabs();
    } else { pushNav(); reload(); }
  });
});
el('clearf').addEventListener('click', () => {
  S.loc = S.cc = S.supplier = S.sg = S.dept = S.src = S.recon = '';
  S.hf = {};
  Object.keys(FILTER_MAP).forEach((k) => {
    if (FILTER_MAP[k] === 'grain') return;
    const n = el(k); if (n) n.value = '';
  });
  typeSelects();
  Object.keys(S.page).forEach((p) => S.page[p] = 1);
  pushNav(); reload();
});

el('tabs').addEventListener('click', (e) => {
  const b = e.target.closest('[data-t]'); if (!b) return;
  S.tab = b.dataset.t;
  pushNav(); paintTabs();
  if (S.tab === 'vbal') ensureVBal();
  if (S.tab === 'vcred') ensureVCred();
  if (S.tab === 'cls' && !S.cls) clsOpen();
});

el('tl-legend').addEventListener('click', (e) => {
  const s = e.target.closest('[data-s]'); if (!s) return;
  if (S.tlHidden.has(s.dataset.s)) S.tlHidden.delete(s.dataset.s); else S.tlHidden.add(s.dataset.s);
  paintTimeline();
});

/* one declaration per toggle pair */
const pair = (onKey, offKey, apply) => {
  el(onKey).addEventListener('click', () => {
    el(onKey).setAttribute('aria-pressed', 'true');
    el(offKey).setAttribute('aria-pressed', 'false');
    apply();
  });
};
pair('tl-abs', 'tl-cum', () => { S.tlMode = 'abs'; S.D.series = buildSeries(); paintTimeline(); });
pair('tl-cum', 'tl-abs', () => { S.tlMode = 'cum'; S.D.series = buildSeries(); paintTimeline(); });
pair('co-cohort', 'co-activity', () => { S.coView = 'cohort'; pushNav(); paintCohort(); });
pair('co-activity', 'co-cohort', () => { S.coView = 'activity'; pushNav(); paintCohort(); });
pair('gr-docs', 'gr-vend', () => { S.grView = 'docs'; S.page.gr = 1; pushNav(); paintGRN(); });
pair('gr-vend', 'gr-docs', () => { S.grView = 'vend'; S.page.gr = 1; pushNav(); paintGRN(); });
el('pa-reg').addEventListener('click', () => { S.payView = 'register'; S.page.pay = 1; pushNav(); paintPayments(); });
el('pa-det').addEventListener('click', () => { S.payView = 'detail'; S.page.pay = 1; pushNav(); paintPayments(); });

/* the ageing basis re-ages client-side from dates the grain already carries,
   so switching it never costs a round trip */
[['ag-bill', 'bill'], ['ag-post', 'posting'], ['ag-due', 'due']].forEach(([k, v]) => {
  el(k).addEventListener('click', () => {
    if (S.agBasis === v) return;
    S.agBasis = v; S.page.ag = 1;
    pushNav(); deriveAgeing(); paintAgeing(); paintTabs();
  });
});
pair('ag-lvl-sup', 'ag-lvl-inv', () => { S.agLevel = 'supplier'; S.page.ag = 1; pushNav(); paintAgTable(); });
pair('ag-lvl-inv', 'ag-lvl-sup', () => { S.agLevel = 'invoice'; S.page.ag = 1; pushNav(); paintAgTable(); });

const debounce = (fn, ms) => { let t;
  return (v) => { clearTimeout(t); t = setTimeout(() => fn(v), ms); }; };
[['po-search', 'po', 'po', paintPO], ['gr-search', 'gr', 'gr', paintGRN],
 ['in-search', 'inv', 'inv', paintInvoices], ['tp-search', 'tp', 'tp', paintTruePurchase],
 ['ag-search', 'ag', 'ag', paintAgTable], ['vb-search', 'vb', 'vb', paintVBal],
 ['vc-search', 'vc', 'vcdn', paintVCred], ['cls-q', 'cls', 'cls', paintCls]
].forEach(([k, key, pageKey, fn]) => {
  const node = el(k); if (!node) return;
  const go = debounce((v) => {
    S.search[key] = v;
    S.page[pageKey] = 1;
    if (key === 'vc') { S.page.vcdn = S.page.vcpe = S.page.vcje = 1; }
    fn();
  }, 220);
  node.addEventListener('input', (e) => go(e.target.value));
});

el('dr-close').addEventListener('click', closeDrawer);
el('dr-back').addEventListener('click', () => drawerBack(false));
el('drawer').addEventListener('click', (e) => { if (e.target === el('drawer')) closeDrawer(); });
window.addEventListener('resize', () => { if (!ROOT.isConnected) return; closeHf(); fitDrawer(); });

/* Frappe binds every letter key on window (the type-anywhere awesomebar) and
   decides "am I typing" from document.activeElement — inside this shadow root
   that is the block host, never the input. Keep keystrokes typed into this
   block's own fields inside the block; Escape still bubbles for the drawer. */
['keydown', 'keypress', 'keyup'].forEach((t) => ROOT.addEventListener(t, (e) => {
  const tg = e.composedPath ? e.composedPath()[0] : e.target;
  if (e.key !== 'Escape' && tg && /^(INPUT|SELECT|TEXTAREA)$/.test(tg.tagName)) e.stopPropagation();
}));
document.addEventListener('keydown', (e) => {
  if (!ROOT.isConnected || e.key !== 'Escape') return;
  if (ROOT.querySelector('.fk-hf')) return closeHf();
  if (S.tab === 'cls' && S.cls && S.clsSel.size) { S.clsSel = new Set(); return paintCls(); }
  if (!el('drawer').hidden) closeDrawer();
});
/* composedPath: the block renders inside a shadow root, so document-level
   targets are retargeted to the host */
document.addEventListener('mousedown', (e) => {
  if (!ROOT.isConnected) return;
  const path = e.composedPath ? e.composedPath() : [e.target];
  if (path.some((n) => n && n.classList
      && (n.classList.contains('fk-hf') || n.classList.contains('hf-btn')))) return;
  closeHf();
});
document.addEventListener('scroll', (e) => {
  if (!ROOT.isConnected) return;
  const path = e.composedPath ? e.composedPath() : [];
  if (path.some((n) => n && n.classList && n.classList.contains('fk-hf'))) return;
  if (Date.now() - HF_OPENED < 700) return;
  closeHf();
}, true);

/* ── one delegated click handler for the whole block ─────────────────────── */
const HF_TARGETS = {
  'ag-table': () => S.agLevel === 'supplier'
    ? { scope: 'ag', cols: COLS.agSupplier, rows: S.A.rows, repaint: paintAgTable }
    : { scope: 'agb', cols: COLS.agBill, rows: S.A.bills, repaint: paintAgTable },
  'cls-table': () => ({ scope: 'cls', cols: COLS.cls,
    rows: S.cls ? (S.cls.view === 'flags' ? S.cls.all.flags : S.cls.all.queue) : [], repaint: paintCls }),
  'vb-table': () => ({ scope: 'vb', cols: COLS.vb, rows: S.vbRows || [], repaint: paintVBal })
};

ROOT.addEventListener('click', async (e) => {
  const hit = (sel) => e.target.closest(sel);

  /* header filter popover + header sort */
  const hb = hit('.hf-btn');
  if (hb) {
    e.stopPropagation();
    const tbl = hb.closest('table');
    const key = tbl && tbl.getAttribute('data-fk');
    const spec = key && HF_TARGETS[key] && HF_TARGETS[key]();
    if (spec) hfOpen(spec.scope, hb.dataset.hfk, spec.cols, spec.rows, hb.closest('th'), spec.repaint);
    return;
  }
  const hc = hit('[data-hfclear]');
  if (hc) { S.hf[hc.dataset.hfclear] = {}; pushNav();
    if (hc.dataset.hfclear === 'cls') paintCls(); else if (hc.dataset.hfclear === 'vb') paintVBal();
    else paintAgTable();
    return; }
  const th = hit('th[data-k]');
  if (th) {
    const tbl = th.closest('table'), key = tbl && tbl.getAttribute('data-fk');
    const spec = key && HF_TARGETS[key] && HF_TARGETS[key]();
    if (spec) {
      const k = th.dataset.k, cur = S.sort[spec.scope] || {};
      S.sort[spec.scope] = { key: k, dir: cur.key === k ? -num(cur.dir || -1) : -1 };
      pushNav(); spec.repaint();
    }
    return;
  }

  /* supplier ledger — any party figure */
  const apl = hit('[data-apl]');
  if (apl) { e.preventDefault(); openApLedger(apl.dataset.apl); return; }

  /* KPI cards */
  const k = hit('[data-kpi]');
  if (k) { const i = KPI_REG.get(k.dataset.kpi);
    if (i && i.drill) i.drill(); else if (i && i.go) i.go();
    return; }

  /* row expansion */
  const po = hit('[data-po]');
  if (po) { await expandDoc('po-table', 'po:', po.dataset.po, 'POItem', 'PO Item',
    COLS.itemPO, 'po_id'); return; }
  const pr = hit('[data-pr]');
  if (pr) { await expandDoc('gr-table', 'pr:', pr.dataset.pr, 'PRItem', 'PR Item',
    COLS.itemPR, 'pr_id'); return; }
  const pi = hit('[data-pi]');
  if (pi) {
    const id = pi.dataset.pi;
    const tr = findRow('in-table', 'pi:' + id);
    if (!tr || collapseKids(tr)) return;
    const items = await lazy('PIItem', 'PI Item', {}, 'PI Item lines');
    const kids = items.filter((r) => r.pi_id === id)
      .concat((S.D.piTax || []).filter((r) => r.pi_id === id));
    renderKids(tr, COLS.itemPI, kids, 'item and tax grains · the header gross is never repeated');
    return;
  }
  /* the ageing table expands without a fetch: the bills are already in the model */
  const ags = hit('[data-agsup]');
  if (ags) {
    const id = ags.dataset.agsup;
    const tr = findRow('ag-table', 'ag:' + id);
    if (!tr || collapseKids(tr)) return;
    const v = S.A.rows.find((x) => x.id === id);
    renderKids(tr, COLS.agBill, v ? v.bills : [],
      v && v.bills.length ? 'aged on ' + S.A.basis.toLowerCase() : 'no open bill for this vendor');
    return;
  }
  const vbs = hit('[data-vbsup]');
  if (vbs) {
    const id = vbs.dataset.vbsup;
    const v = (S.vbRows || []).find((x) => x.supplier === id);
    const inv = where(S.L.open, (r) => r.supplier === id);
    if (v) drill('Open invoices — ' + v.name, inv, COLS.openItems, 'outstanding',
      sum(inv, 'outstanding'), null, { supplier: id });
    return;
  }

  /* ageing actions */
  const act = hit('[data-agact]');
  if (act) {
    const id = act.dataset.agid, a = act.dataset.agact;
    if (a === 'Reconcile') openReconciliation(id);
    else if (a === 'Obtain bill') window.open('/app/supplier/' + q(id), '_blank', 'noopener');
    else openApLedger(id);   /* Investigate and Review both open the vendor's own ledger */
    return;
  }
  const band = hit('[data-band]');
  if (band) {
    const i = parseInt(band.dataset.band, 10), b = BANDS[i];
    const rows = S.A.bills.filter((r) => r.amt > 0 && r.band === b.label);
    drill('Ageing bucket — ' + b.label + ' (' + S.A.basis.toLowerCase() + ')',
      rows, COLS.agBill, 'amt', S.A.bands[i]);
    return;
  }

  /* classification panel */
  const cls = hit('[data-cls]');
  if (cls) { clsAction(cls.dataset.cls, cls); return; }

  /* charts, routes, instruments */
  const bar = hit('[data-bar]');
  if (bar) {
    const v = bar.dataset.bar;
    if (v.indexOf('status:') === 0) {
      const st = v.slice(7), g = S.D.po.byStatus.find((x) => x.key === st);
      if (g) drill('Purchase orders — ' + st, g.rows, COLS.poFull, 'gross', g.v);
    } else if (v.indexOf('bankadv:') === 0) {
      const b = v.slice(8), g = S.D.bankAdvInv.find((x) => x.key === b);
      if (g) drill('Advances — ' + b, g.rows, COLS.payAdv, 'advance', g.v);
    }
    return;
  }
  const rt = hit('[data-route]');
  if (rt) { const r = ROUTE_REG[+rt.dataset.route];
    if (r) drill('Booking route — ' + r.label, r.rows, COLS.piFull, 'gross', r.v);
    return; }
  const bk = hit('[data-bank]');
  if (bk) { const b = BANK_REG[+bk.dataset.bank];
    if (b) drill('Instrument — ' + b.key, b.rows, COLS.payFull, null, b.total, payVal);
    return; }
  const lr = hit('[data-locrec]');
  if (lr) {
    const [i, part] = lr.dataset.locrec.split(':');
    const x = LOCREC[+i]; if (!x) return;
    const map = {
      O: ['Open items', x.open, COLS.openItems, 'outstanding', x.O],
      J: ['Journals not linked to an invoice', x.jr, COLS.journal, 'outstanding', x.J],
      U: ['Unapplied cash', x.un, COLS.unapplied, 'advance', x.U],
      C: ['AP control', x.ct, COLS.control, 'outstanding', x.C]
    };
    const m = map[part];
    if (m) drill(m[0] + ' — ' + x.label, m[1], m[2], m[3], m[4]);
    return;
  }
  const cohort = hit('[data-cohort]');
  if (cohort && S.coView === 'cohort') {
    const c = S.D.cohort.find((x) => x.k === cohort.dataset.cohort);
    if (c) drill('PO cohort — ' + bucketLabel(c.k), c.rows, COLS.poFull, 'ordered', c.ordered);
    return;
  }
  const stage = hit('[data-stage]');
  if (stage) {
    const st = stage.dataset.stage, b = stage.dataset.bucket;
    const inB = (r) => bucketKey(r.date) === b;
    const lbl = ' · ' + bucketLabel(b);
    if (st === 'ordered')  drill('Ordered' + lbl, where(S.D.poDoc, inB), COLS.poFull, 'gross', null);
    if (st === 'received') drill('Received' + lbl, where(S.D.prDoc, inB), COLS.prFull, 'gross', null);
    if (st === 'invoiced') drill('Invoiced' + lbl, where(S.D.piDoc, (r) => inB(r) && IS_P2P(r)), COLS.piFull, 'gross', null);
    if (st === 'paid')     drill('Paid' + lbl, where(S.D.p2pAlloc, inB), COLS.payFull, 'paid', null);
    return;
  }

  /* vendor summary row → filter the document rows to that vendor */
  const grv = hit('[data-row^="grv:"]');
  if (grv) {
    S.supplier = grv.dataset.row.slice(4);
    const n = el('f-supplier'); if (n) n.value = S.supplier;
    typeSelects();
    S.grView = 'docs';
    el('gr-docs').setAttribute('aria-pressed', 'true');
    el('gr-vend').setAttribute('aria-pressed', 'false');
    pushNav(); reload();
    return;
  }
  const tp = hit('[data-row^="tp:"]');
  if (tp) {
    const id = tp.dataset.row.slice(3);
    const v = S.D.truePurchase.find((x) => x.supplier === id);
    if (v) drill('Vendor — ' + v.name, v.invRows, COLS.piFull, 'gross', v.net, null, { supplier: id });
    return;
  }

  /* exports */
  const ex = e.target.closest('[data-fk]');
  if (ex && EXPORTS[ex.dataset.fk]) {
    if (!S.loaded) return status('Nothing loaded yet.', 'err');
    try { EXPORTS[ex.dataset.fk](); }
    catch (err) { console.error('[AP] export', err); status('Export failed — see the browser console.', 'err'); }
  }
});

el('cls').addEventListener('change', (e) => {
  const t = e.target.closest('[data-cls]'); if (!t || !S.cls) return;
  const a = t.dataset.cls;
  if (a.indexOf('prop:') === 0) {
    S.cls.proposed[a.slice(5)] = t.value;
    S.clsSel.add(a.slice(5));
    paintCls();
  }
});

/* ══ 26 · boot ═══════════════════════════════════════════════════════════ */

async function reload(keepPages) {
  if (!keepPages) Object.keys(S.page).forEach((p) => S.page[p] = 1);
  await run();
  loadDimensions();
  clsPendingCount();
  if (S.cls) clsOpen();
  paintTabs();
}

(async function boot() {
  status('Starting…', 'busy');

  /* Typography paints the house look into this block's own root. SchemaGuard
     proves every DocType and field this dashboard depends on still exists
     before a single figure is shown — an ERPNext upgrade cannot rename a field
     underneath us and leave a wrong number on screen. */
  window.FurnishkaType.inject(ROOT);
  const deps = await window.FurnishkaGuard.require({
    'GL Entry':         ['posting_date', 'account', 'party_type', 'party', 'debit', 'credit', 'is_cancelled'],
    'Purchase Invoice': ['posting_date', 'due_date', 'supplier', 'outstanding_amount', 'docstatus'],
    'Supplier':         ['supplier_name', 'disabled']
  });
  if (!deps.ok) {
    window.FurnishkaGuard.banner(ROOT, deps.missing);
    status('Dependency check failed — figures suppressed.', 'err');
    return;
  }
  window.FurnishkaGuard.stamp(ROOT, { canon: 'canon v1', built: '06-Oct-2026',
    type: window.FurnishkaType.REV });

  applyPreset('fy');
  if (!COMPANY) { status('No default company is set for your user.', 'err'); return; }
  try {
    /* Back from an ERP report re-creates this block: rebuild the view that was left */
    const back = history.state && history.state.v ? history.state.v : null;
    if (back) applySnap(back);
    await reload(!!back);
    if (back) restoreScroll(back.scrollY);
    else history.replaceState({ fk: 'view', v: navSnap() }, '');
    console.log('[AP v29] ready · ' + countOf(S.A.rows ? S.A.rows.length : 0, 'supplier') + ' · bills '
      + inr(S.A.billTot) + ' · net ' + inr(S.A.net) + ' · party ledger ' + inr(S.A.led)
      + ' · GL ' + (S.A.gl === null ? 'not comparable at this scope' : inr(S.A.gl))
      + ' · off-identity ' + cnt(S.A.off));
  } catch (e) {
    console.error('[AP] boot', e);
    status('Could not start. ' + (e && e.message ? e.message : ''), 'err');
  }
})();

})();
