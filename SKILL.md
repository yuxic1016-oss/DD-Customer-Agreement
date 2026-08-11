---
name: jax-tenant-agreement-extraction
description: Use when the user wants to extract tenant/customer agreement terms (MSAs, Sales Orders, leases, amendments, renewal notices) from a folder of PDFs into a "Customer Agreements Summary" style workbook, or update/rebuild that workbook for a data-center or similar facility. Trigger on requests to pull customer contract data into a spreadsheet, build or refresh a tenant agreement tracker, process a folder of customer/tenant PDFs into structured rows, or extend an existing Customer Agreements Summary with new customers or a new as-of date. Applies even if the user doesn't name this skill directly — e.g. "extract these customer folders into the summary sheet" or "add these new tenants to the agreements workbook."
---

# JAX Tenant Agreement Extraction

## Overview

Turns a folder of customer/tenant agreement PDFs (MSAs, Sales Orders, leases,
amendments, renewal letters) into rows in a Customer Agreements Summary
workbook, with a single "As-of Date" cell that automatically recalculates each
customer's current recurring charge — including compounding escalations —
without anyone re-deriving "what's active right now" by hand every time the
sheet is reopened.

This captures a workflow validated end-to-end across 120+ real customer
folders for a Jacksonville data-center facility, refined through the account
owner's own review pass. Read `references/conventions.md` before extracting —
it has the full data model, controlled vocabulary, category codes, and the
hard-won edge cases (leases vs. colo orders, escalation anchoring, formula
safety). This SKILL.md is the workflow; that file is the detail you'll need
while doing the actual extraction.

## When to Use

Use this whenever the deliverable is a spreadsheet of customer/tenant
agreement terms built from a folder of source contract PDFs, whether that's:
- A brand-new pilot on a handful of customer folders
- Scaling up to dozens/hundreds of folders
- Adding new customers to an existing workbook
- Rebuilding a customer whose folder was updated (new amendment, renewal, etc.)
- Rolling the As-of Date forward on an existing workbook (no re-extraction
  needed — just update cell C3 on each sheet; the formulas do the rest)

Don't use this for one-off single-document questions ("what does this lease
say about parking") — that's just reading a PDF. This skill is for building or
maintaining the structured, formula-driven summary workbook.

## Workflow

### 1. Confirm scope with the user

Before extracting anything, confirm: which folder(s) of customer subfolders,
whether this is a new sheet or an addition to an existing workbook, and what
As-of Date to use (default: today). If the user has an existing workbook in
this format already, use it as the template/style donor and as the source of
truth for any conventions that differ from `references/conventions.md` — her
actual reviewed file always wins over this skill's defaults.

### 2. Extract each customer folder into structured JSON

Read `references/conventions.md` §§1–3 for what counts as a customer, the
standing "always trace to source, always flag uncertainty" rule, and how to
handle real-estate leases differently from standard colo Sales Orders.

For small batches (a handful of customers), extract inline. For large folders
(dozens+), dispatch parallel subagents — see `references/conventions.md` §9
for batching and prompt-construction guidance. Each customer's extraction
should produce an object matching the schema documented at the top of
`scripts/build_workbook.py` (kinds: `msa`, `so_fixed`, `so_recurring`, `open`;
category codes; notes flagging). Save every batch's JSON output to disk before
moving on — don't rely on keeping it only in a chat response, since large
batch runs can be interrupted.

### 3. Assemble the workbook

Combine the extracted customer objects into one JSON file matching the
top-of-file schema in `scripts/build_workbook.py`, then run it:

```
python3 scripts/build_workbook.py \
  --template "<path to the original master template xlsx>" \
  --data customers.json \
  --sheet-name "<new sheet name>" \
  --out output.xlsx
```

Use `--append-to <existing_workbook.xlsx>` instead of relying only on
`--template` when adding a new sheet to a workbook that already has other
sheets you want to preserve (it still uses `--template` for style-donor rows,
but adds the new sheet onto the existing workbook rather than a fresh copy of
the template).

The script handles: banner/subtitle rows, the As-of Date cell, header styling
donated from the template, per-customer blocks, EDATE-based escalating period
rows, the SUMPRODUCT Current Total formula per customer, the header row's
pointer formula back to Current Total, and the `safe()` formula-injection
guard on every text field. You should not need to hand-write these formulas —
if a customer's situation doesn't fit the four row kinds, that's a sign to
re-read `references/conventions.md` §4 rather than hand-patching cells after
the fact.

### 4. Verify before delivering

Run LibreOffice recalculation and confirm zero formula errors:

```
python3 <xlsx-skill-path>/scripts/recalc.py output.xlsx
```

Check `total_errors: 0` in the result. Then spot-check a few customers by
reading back `data_only=True` values: confirm at least one `so_fixed` past its
end date reads `"N/A"`, and at least one `so_recurring` customer's Current
Total reflects the escalated period that actually contains the As-of Date (not
just the first period). This is the same check performed while building this
skill — see the worked example in `references/conventions.md` if you want to
sanity-check your own output against known-good numbers.

### 5. Deliver

Copy the finished workbook to the user's selected folder and present it. If
this is an update to an existing workbook the user already has open/saved
elsewhere, make sure the sheet name doesn't collide with one already present
unless the intent is to replace it.

## Common Mistakes

- **Trusting a prior/master sheet's numbers instead of the source folder.**
  Always re-derive from the documents in that customer's folder, even if a
  number already exists somewhere else — stale master-sheet data was the
  original bug this workflow was built to eliminate.
- **Blending multiple suites/spaces of a real-estate lease tenant into one
  row.** Keep each suite's rent as its own line so the totals stay legible
  and audit-able.
