# JAX Customer Agreements — Extraction & Formatting Conventions

This is the detailed reference for the Jacksonville data-center tenant agreement
tracking project. Read this when extracting from source PDFs and when preparing
the JSON handed to `scripts/build_workbook.py`. It documents the conventions
Scarlett settled on after reviewing several drafts — treat these as the target
format, not just one possible approach.

## 1. Source materials and what "a customer" means

Each customer/tenant has a folder of PDFs: Master Service Agreements (MSAs),
Sales Orders (SOs), amendments, renewal letters, rate-increase notices, and
sometimes full commercial leases (for real-estate tenants like Cologix, not
standard colo customers). A folder may contain multiple SOs stacked over time
(original + amendments + renewals) — these represent ONE evolving commitment
and should become one `so_recurring` or `so_fixed` row (with period-row
formulas or endpoints reflecting the amendments), not separate line items,
UNLESS the amendments describe genuinely distinct suites/spaces/services, in
which case keep them as separate row entries (see Cologix handling below).

**Customer IDs**: use the numeric ID from the folder name (e.g. `1131`,
`2209`). Never invent one and never reuse a stale ID/number sitting in an old
master sheet — always re-derive from the folder's own documents.

## 2. The standing extraction rule

Every field in the workbook must trace back to something found in that
customer's own folder. If a folder has no usable documents, is empty, or
contains only internal/non-billing paperwork, still create the customer's row
with `"rows": []` and a `header_notes` explanation — do not silently omit the
customer, and do not carry forward numbers from a prior/master version of the
sheet on the assumption they're already correct. If a document is ambiguous,
contradictory, illegible, or you're not fully confident in a value, write your
best-effort value AND add a note flagging the uncertainty in that row's
`notes` field — never guess silently.

Common things to flag in notes:
- Illegible/OCR'd text you're not fully confident in
- Conflicting dates or amounts across amendments
- A renewal letter referencing an SO number not found in the folder
- Internal/non-tenant accounts (test accounts, employee accounts) — flag and
  still include with $0/notes rather than dropping
- Filename vs. content mismatches (seen with Cologix — verify before trusting
  a filename)

## 3. Reading a lease vs. a standard colo Sales Order

Standard colo customers: MSA + one or more SOs for cabinets/power/cross-connects/etc.
Real-estate tenants (e.g. Cologix leasing whole suites): read as a commercial
lease. Do NOT blend multiple suites into a single combined rent figure — each
suite gets its own row (or row-group) with its own rent history, even if the
tenant is the same entity across suites. Skip floor plan and CAM reconciliation
documents; they don't contain rent commitment terms. Follow every amendment
chain to determine the CURRENT rent for each suite as of signing, then let the
period-row/escalation mechanism (below) carry it forward.

## 4. Row kinds and when to use each

| kind | use for | begin/end behavior |
|---|---|---|
| `msa` | Governing agreement text, renewal/holdover/payment terms, escalation-rate anchor | No dates in Beginning/Ending/Total MRC columns — text only |
| `so_fixed` | A commitment with a confirmed, bounded end date and no confirmed renewal | Static begin+end; naturally drops out of Current Total once As-of Date passes `end` |
| `so_recurring` | Auto-renewing / open-ended commitment (the common case for active colo customers) | Static first period, then generated EDATE period-rows compounding at the escalation rate, extending past the As-of Date |
| `open` | No term at all — a standing base fee or always-on item | No dates; always counted |

Don't default everything to `so_recurring` just because it's the fanciest
pattern — use `so_fixed` whenever the source document actually states a fixed
end with no automatic renewal language. The Current Total should reflect what
the documents actually commit to, not an optimistic guess.

## 5. Escalation rate

Written as a decimal fraction (e.g. `0.03` for 3%) in column K ("MRC
Escalations", displayed as a percentage) on the FIRST `msa` row of a
customer's block. Every `so_recurring` row's generated period formulas
reference this one cell absolutely (`$K$<anchor_row>`), so a single edit to
the escalation rate updates every projected period for that customer. If no
escalation is documented, leave it blank/0% — never assume 3% as a default.
If different SOs within the same customer genuinely escalate at different
confirmed rates, that's an edge case worth a note and manual adjustment after
the automated build (rare — only do this if the source documents are explicit
about diverging rates).

## 6. The MRC As-of Date and Current Total mechanism

Cell C3 ("MRC As-of Date") is the single control for the whole sheet. Every
customer's "Current Total" row uses a SUMPRODUCT array formula that:
1. Includes a period row if its Beginning Date is blank OR ≤ the As-of Date
2. AND its Ending Date is blank OR ≥ the As-of Date
3. AND its Total MRC is numeric
4. Sums the Total MRC of all matching rows; returns `"N/A"` if nothing matches

This means changing C3 alone re-derives "what's currently active" without any
manual row-picking — this is the mechanism `scripts/build_workbook.py`
generates automatically per customer block. Don't hand-write static
`=N<row>` references to "the current row" — they go stale the moment the
As-of Date changes or a new period is appended.

The customer header row's Total MRC column (D) is a simple pointer
(`=N<current_total_row>`) so subtotals/exports referencing column D stay in
sync automatically.

## 7. Controlled vocabulary (condense to these where the source document
fits one of these patterns; use a close plain-language phrase for anything
that doesn't):

**Renewal Options** — common values seen: "None", "One-time", "Auto-renews
annually", "Auto-renews monthly", "Option to renew — 1x 3-year term", "Option
to renew — 1x 5-year term". Keep phrases short; put the fine print in Notes.

**Payment Terms** — common values seen: "Net 15", "Net 30", "Due upon
receipt", "Prepaid annually", "Prepaid quarterly".

**Holdover** — common values seen: "Month-to-month", "N/A", "150% of MRC
during holdover" (or whatever multiplier the lease states).

Don't force a value into one of these buckets if the source document says
something meaningfully different — write the actual term and let it stand as
its own value rather than distorting it to match the list.

## 8. Category codes (columns O–AG)

```
PWRFIX  Fixed Power       CG     Cage              FC    Fiber/Cabinet
CC      Cross Connect     PWRP   Power (Provisioned) PWRR Power (Reserved)
PRTF    Port/Interface    SMF    Single-mode Fiber  IPA   IP Addresses
INET    Internet          BAS    Base/Space          XCN   Cross-connect (alt)
PAN     Panel             VPN    VPN                 RRS   Remote Reboot/Smart-hands
MST     Managed Services   FBR    Fiber               BCS   Broadband/Circuit
INTRA   Intra-building
```
If a source document uses a category label that doesn't map cleanly to one of
these codes, don't silently drop the dollar amount — keep it in the row's
`total_mrc` and note the unmapped label in `notes` so a human can decide where
it belongs. `scripts/build_workbook.py` does this automatically (appends
`[Unmapped category codes kept in notes rather than dropped: ...]`).

## 9. Extraction mechanics (PDF → JSON)

- Use `pdftotext -layout` first; it handles the large majority of documents.
- If output is empty/garbled (common with DocuSign merge-field encoded PDFs or
  scanned documents), fall back to OCR: `pdftoppm -r 200 -png` then
  `tesseract` on each page image.
- If a specific page is corrupted/malformed and neither tool renders it,
  Ghostscript can often re-render/repair it as a last resort.
- For large folders (dozens to hundreds of customers), dispatch parallel
  subagents rather than working serially. Balance batches by file count
  (greedy bin-packing), not customer count — some customers have 1 PDF, others
  have 15. Give each subagent a fully self-contained prompt: the schema above,
  the standing extraction rule, known patterns/gotchas, and an explicit
  instruction to write its output JSON to an exact file path (so results
  survive even if the subagent's final chat summary is imperfect or gets
  interrupted).
- After a batch returns, spot-check anything the subagent flagged as
  surprising (e.g. filename/content mismatches) independently before treating
  it as fact — this caught a real issue with the Cologix folder where ~13 of
  16 files were misnamed.

## 10. Formula-injection safety

Any string written to a cell that starts with `=`, `+`, `-`, or `@` will be
interpreted by Excel/openpyxl as a formula, silently corrupting the cell.
`scripts/build_workbook.py`'s `safe()` function guards every text write by
prepending a space if needed — always route customer names, descriptions, and
notes through it (or an equivalent) rather than writing raw extracted text.

## 11. Verification before delivering

Always run `recalc.py` (from the `xlsx` skill, LibreOffice-based) against the
finished workbook and confirm `total_errors: 0` before calling the work done.
A clean save from openpyxl is not sufficient proof the formulas are valid —
LibreOffice recalculation is the actual check.
