# Validation rules

Run every check; write each result to the `Validation` sheet as
`File | Check | Invoice value | Packing value | Generated value | Status` with Status `OK`,
`MISMATCH` or `REVIEW`. Any `MISMATCH` blocks the upload. `REVIEW` ships with the file but is
called out in the summary to the reviewer.

## 1. Document integrity (before mapping)
- `Inv-` and `Pac-` invoice numbers match; dates match; ports match. MISMATCH otherwise.
- Shipper name resolves to one vendor. If unknown, every GRPO carries `LOOKUP:CardCode[...]` → REVIEW.
- Table header row found (`CTN. (S)`), ≥ 1 data row, totals row found.
- Every PO# parses to prefix + 6 digits. A row that does not parse → MISMATCH (do not guess).
- Every STYLE# yields a non-empty ItemCode.

## 2. Reconciliation (per invoice)
| Check | Invoice | Packing | Generated |
|---|---|---|---|
| Pieces | Total pcs (`Inv-` r83) | totals row Total PCS | Σ `Quantity` |
| Cartons | Package CTNS (`Inv-` r8) | totals row CTN | Σ `U_NumfCart` |
| Gross weight kg | Total G.W | totals row Total G.W | Σ line Total G.W (2 dp) |
| CBM | - | totals row CBM | Σ CBM (4 dp) |
Tolerance 0.01 (weights, CBM), exact for counts.

- Carton ranges: `CTN.(E) - CTN.(S) + 1 = CTN` on every row; ranges across the whole packing list
  are contiguous and non-overlapping from 1 to total cartons. Gaps or overlaps → REVIEW
  (invoice 4764 has a gap 101-240 in the factory's numbering; report it, do not "fix" it).
- `QTY/CTN × CTN = Total PCS.` on every row; otherwise REVIEW (partial carton).

## 3. Structure
- One GRPO per invoice (one container, many POs). Report distinct PO count; no flag.
- A PO whose rows carry more than one prefix → REVIEW (likely a factory typo).
- Same PO + ItemCode + color on multiple rows → merged; REVIEW with the packing split.
- Same PO on more than one invoice in the batch → REVIEW "split shipment, partial receipt".
- Same ItemCode on the same PO in two colors is normal (two PO lines) - no flag.

## 4. SAP consistency (only when a lookup source is available)
- PO exists, is open, vendor on PO = CardCode.
- PO line found for ItemCode + color (case-insensitive, ignore `/`, `-`, spaces). No match → MISMATCH.
- Received Quantity ≤ PO line open quantity. Over-receipt → MISMATCH with both numbers.
- Price populated from the PO line; `Price = 0` → REVIEW.
- WhsCode populated.

## 5. Output hygiene
- Column names and order identical to the template, both sheets.
- Dates written as true dates, not `DD.MM.YYYY` strings.
- No cell left as an empty string where the template says `NULL`.
- File named `OPDN_upload_<invoice no with non-alphanumerics → _>.xlsx`, e.g. `OPDN_upload_EX_UNI_4762_2026.xlsx`.
- Summary to the reviewer: invoice no., date, vendor, GRPO count, line count, pieces, cartons,
  GW, CBM, number of REVIEW and MISMATCH items, number of unresolved `LOOKUP:` cells.
