# Accounts Payable — Procure to Pay · Custom HTML Block

Three files, one per field of the ERPNext **Custom HTML Block** at `/app/payables`.
Paste each into the matching field, replacing the whole field:

| File | Custom HTML Block field |
|---|---|
| `ap.html` | **HTML** |
| `ap.js`   | **Javascript** |
| `ap.css`  | **Style** |

Paste in that order and hard-refresh. Nothing else needs to change: the server-side
Query Report *AP Procure to Pay Master*, the four `Finance-Lib-*` Client Scripts and
the `AP Location Classification Override` DocType are all unchanged.

## What this version is

A full rebuild (**v29**) of what the page rendered at v28.2, with the eleven stacked
patch layers collapsed into one generation. Same tabs, same figures, same drill-downs
— written once instead of eleven times.

### Removed because it was dead

Each of these produced nothing a user could see, because a later layer hid it:

- the Overview block strip, the subledger-to-GL table and the exception census
  (three cards carrying `style="display:none"`)
- the original AP Aging KPIs, chart, table and pager, and the vendor ledger card
  (v24 hid every child of the panel)
- v21's bridge and credits, v24's four-bucket supplier table, and the
  advance-booked / advance-not-booked cards — now the *Bills + advance* and
  *Advance only* classes of the one table
- `po-advnote`, `pa-note`, `vb-why`, `vc-why`, `dr-sub`, `cls-open`, and every
  card subtitle (all permanently hidden by script or stylesheet)

### Removed because it was duplicated

| Was | Now |
|---|---|
| 5 download helpers | 1 `dl()` |
| 3 XLSX writers | 1 `xlsxBook()` |
| 2 header-filter engines | 1, with the union of both feature sets |
| 6 monkey-patched functions | 0 — each declared once, in order |
| 2 MutationObservers, 3 polling intervals | 0 |

### The one substantive change

v28 built the merged supplier table by reading `textContent` out of the `<td>`s of
the two tables it then hid, parsing the formatted strings back into numbers, and
joining that to a ledger read taken at a different moment. That is why it had to
warn about a *capture gap* and would not call itself a reconciliation.

The merged model is now computed from the report rows in the same snapshot as
everything else, so the warning has no reason to exist. The identity it proves is
also complete, where v28's subtraction was not:

```
bills + debit notes − advance + payments on account + journals = party ledger
party ledger (AP GL Control grain)                             = GL control
```

A non-zero **Delta** therefore means the `AP Aging` and `AP GL Control` grains
disagree for that supplier — a real exception, not an artefact of a partial formula.

Fixing that surfaced a matching gap in the Overview bridge, which read
`open items + journals − unapplied` and omitted a Payment Entry carrying a positive
payable balance. That term is now an explicit line in both places, so the page no
longer holds two different formulas for one tie.

## Conventions

- **Sign** — party ledger is credit minus debit. Cr positive = owed to the supplier.
  Dr negative = net receivable. A Dr balance is never netted into the payable total;
  Schedule III requires it under Other Current Assets, and the statement says so.
  The native General Ledger prints Dr-positive, so a payable reads negative there —
  same magnitude, opposite sign. Stated on the face of the table.
- **Numbers** — Indian 2-2-3 grouping with paise everywhere, never abbreviated.
  Tolerance 0.50, shown, never forced.
- **Fiscal year** — from the Fiscal Year master by containment, never by month
  arithmetic. There is no year literal in the file.
- **Writes** — only `AP Location Classification Override`, only from the Categorise
  panel, always reversible. No ERP transaction, master or ledger is ever written.

## Known ERP defects this page reports rather than hides

1. Item-level GST does not tie to posted tax, so all statutory tax comes from the
   `PI Tax` grain and never from an item field.
2. `is_tax_withholding_account` is unreliable; TDS is resolved by account name.
3. No Payment Entry anywhere references a Purchase Order, so the advance card
   reports unapplied supplier cash instead of a PO advance of 0.00.
4. `Purchase Invoice.unallocated_amount` understates unapplied cash; the derived
   `AP Unapplied` grain is the reconciling figure and both are shown side by side.
5. Almost no supplier payment carries a clearance date, so bank reconciliation is
   an exception to act on, not a ratio to watch.
