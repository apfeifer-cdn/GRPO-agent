---
name: factory-invoice-to-opdn
description: Turn a factory Invoice + Packing List workbook into an SAP B1 GRPO upload file (OPDN_Header / PDN1_Line) for DRI - one GRPO per invoice (container), one line per packing-list row, with reconciliation and a lookup list.
---

# Factory Invoice → OPDN (GRPO) upload

## Purpose
Receive a factory shipping document (`Inv-` + `Pac-` sheets, Unimas-style `.xls`) and produce
an Excel workbook shaped exactly like `OPDN_Excel.xlsx` so a Goods Receipt PO can be created in
SAP Business One. The factory docs carry quantities, packing and shipping facts. They do **not**
carry prices, vendor code, or PO internal keys - those come from the SAP PO. Never invent them.

## Inputs
- Factory workbook(s). Line source is the `Pac-` table (header row starts `CTN. (S)`).
  `Inv-` supplies invoice no., date, ports, EXP no., HS code, totals. Ignore the number-to-words
  grid below row ~90 of `Inv-`. Details: `references/source-document-anatomy.md`.
- Template `OPDN_Excel.xlsx` (column names and order for both sheets; row 2 is an old example).
- PO lookup table (OPOR / POR1 / OITM export, `references/po-lookup-table.md`). If absent, emit `LOOKUP:` placeholders.

## Procedure
1. **Read header facts** from `Inv-`: shipper name/address, Invoice No. (trim), DATE, Contract,
   HS Code, Port of Shipment, Port of Discharge, Package CTNS, EXP No., ICD DATE, Shipped,
   consignee block, Total pcs / G.W / N.W. Confirm `Pac-` header repeats the same invoice no.,
   date and ports. **All dates are DD.MM.YYYY.**
2. **Read packing rows** until the first blank `CTN. (S)`; that row is the totals row.
   Per row parse:
   - `PO#` → collapse spaces → `^([A-Z0-9]{3})(\d{6})\s*(.*)$` → `U_POPrefix`, `U_OrderNo`, `U_CustRefNo`.
   - `STYLE#` → text before `(` is `U_Proto` (OITM.U_Proto, the match key); text in parens is
     the customer `U_StyleId`, except for SPO/VAC where it is the matrix parent style
     `U_MxPrStyl`. Neither is a PDN1 column: both go to FreeTxt and the `GRPO_Lines` sheet.
   - COLOR → `U_Color`; DESCRIPTION → `Dscription`; Total PCS → `Quantity`;
     CTN.(S)-CTN.(E) → `U_CartonNo`; CTN → `U_NumfCart`; QTY/CTN → `U_PcsPerCarton`;
     G.W/CTN → `U_CartonWeight`; L/W/H → `U_CartonLength/Width/Height`; PI# + SEQ → FreeTxt.
   A row whose PO# does not parse stops the run with MISMATCH.
3. **One GRPO per factory invoice** (confirmed - one container, many POs). One `OPDN_Header` row, `DocEntry` = 1; every
   line carries `DocEntry` 1 and `LineNum` 0..k in packing-list order. Each line keeps its own
   PO via `BaseRef`/`BaseEntry`/`BaseLine` and `U_LinePOPrefix`. With more than one PO on the
   invoice the single-PO header UDFs `U_OrderNo`, `U_POPrefix`, `U_CustRefNo` are `NULL` and
   the PO list goes in `Comments` ("Based On Purchase Orders 120601, 120603, …"). Normal - a
   GRPO is one container, which consolidates many POs; not logged.
   **One GRPO line per packing-list row. Never merge rows**, even when two rows hit the same PO
   line (same style and color packed two ways): carton-level data on each line is required for
   reporting. Both lines reference the same `BaseEntry`/`BaseLine`; over-receipt is checked on
   the summed quantity per PO line.
4. **Fill the header** per `references/opdn-field-mapping.md`: dates = invoice date;
   `NumAtCard` = `U_FactoryInv` = invoice no.; `CardName`/`Address` from shipper;
   `Address2` consignee; `Comments` = `Based On Purchase Orders <list>. Factory invoice <no>
   dated <MM/DD/YYYY>. EXP No. <exp>. HS <hs>. Contract <contract>.`; `U_FOB` port of shipment;
   `U_OceanDest` `Kansas City, KS`; `U_COO` from shipper country; `U_TermSale` FOB; `U_MOT` Sea;
   `U_HandoverDate` ICD date; `U_NumCartons` / `U_GrossWeight` / `U_CubeMeters` = invoice totals
   (also `Weight`, `Volume`). New-document status: `DocStatus` O, `InvntSttus` O, `Printed` N,
   `Posted` N. `U_BLConsig` = `U_GenConsig` = `Design Resources, Inc.`
5. **Fill the lines**: `BaseType` 22, `BaseRef` = `BaseDocNum` = PO number, `Currency` `$`,
   `CountryOrg` BD, `LineStatus`/`InvntSttus` O, `ShipDate`/`ActDelDate` = invoice date,
   `InvQty`/`PackQty`/`BaseQty` = Quantity, `U_LinePOPrefix` = prefix.
6. **Resolve lookups** from the PO lookup table (`references/po-lookup-table.md`): match on
   `PO + OITM.U_Proto + color` (color upper-cased, `/ -` spaces removed, GREY→GRAY), e.g.
   `121019 + C17-06417 + NAVY` → `SP500-INDI-OS` line 7. Take `BaseEntry`, `BaseLine` (also
   `U_BaseEntry`/`U_BaseLine`), `Price`, `U_FOBPrice`, `WhsCode`, `CardCode` (also `BaseCard`),
   SAP `ItemCode`. No `U_Proto` hit → ItemCode exact, C-number prefix, then style prefix, each
   with color equality; never match across colors. Summed qty per PO line > OpenQty → MISMATCH.
   Packing style ≠ OITM.U_MxPrStyl → REVIEW.
   Vendor = PO vendor, not the shipper; shipper goes to Comments. Report the table's
   `ExportedAt`. Run exactly: `python scripts/build_opdn.py "<factory.xls>" --lookup "<po-lookup-table.xlsx>"`
   (template and output dir default; output lands in `./out/`). xlrd is bundled. Without a lookup
   source write `LOOKUP:<key>` in the cell and list each on sheet `Lookups_Needed`.
7. **Every other column**: apply `references/opdn-column-defaults.csv`. CONST → template value;
   SYSTEM → `NULL` (SAP assigns: DocNum, TransId, totals, create/update stamps, user signs,
   landed-cost UDFs, account codes, Series, payment terms); NULL → `NULL`. Never leave a cell blank.
8. **Validate** per `references/validation-rules.md`: pieces, cartons, gross weight and CBM
   reconcile across invoice, packing list and generated lines; carton ranges contiguous and
   consistent with CTN; QTY/CTN × CTN = pcs; one prefix per PO; same PO across invoices flagged
   as split shipment. Write results to sheet `Validation`. Any MISMATCH = do not hand off as
   uploadable.
9. **Write** `OPDN_upload_<invoice no, non-alphanumerics → _>.xlsx` with six sheets:
   `OPDN_Header` + `PDN1_Line` (full template columns, identical order, true date cells),
   `GRPO_Header` + `GRPO_Lines` (lean staging: only the ~45 columns that carry data, same
   values - this is the reviewable form and maps 1:1 to the Service Layer payload),
   `Validation`, `Lookups_Needed`.
10. **Report**: relay the script's `=== GRPO SUMMARY ===` stdout block verbatim (counts,
    totals, ExportedAt, ready-to-post, every flagged item). Attach the workbook as a downloadable
    file; a sandbox path is not a deliverable. If it cannot be attached, rerun with `--emit-csv`
    and return the `GRPO_Header` and `GRPO_Lines` CSV blocks.

## Rules
- Prices and PO keys come from SAP only. No estimates, no copying from the template example row.
- Do not renumber or "repair" factory carton numbering; flag gaps.
- One invoice in, one workbook out; several invoices → separate workbooks.
- Unknown factory layout: locate fields by label text; if a required field is missing, stop and name it.

## Decisions and open items
- Decided: one GRPO per invoice (container); one line per packing row; SAP side via Service
  Layer, the workbook is the review artifact (`GRPO_Header`/`GRPO_Lines`).
- Open: `Series`; `U_HandoverDate` for ICD date; `U_MakerTerms` meaning (template `CDN`).

## Package
- `references/source-document-anatomy.md` - where every field lives, parsing rules.
- `references/opdn-field-mapping.md` - header and line mapping, importer caveat.
- `references/opdn-column-defaults.csv` - all 1,174 columns classified with value rules.
- `references/validation-rules.md` - checks and statuses.
- `references/po-lookup-table.md` - OPOR/POR1 export query, table schema, matching procedure (current lookup source).
- `references/service-layer-po-lookup.md` - Service Layer lookup and direct GRPO POST (later replacement).
- `scripts/build_opdn.py` - deterministic reference implementation (xlrd + openpyxl).
- `examples/` - template, generated output for 4762/4763/4764, one joined against a fake table.
