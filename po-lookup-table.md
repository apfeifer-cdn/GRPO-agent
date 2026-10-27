# PO lookup table (OPOR / POR1 snapshot) - resolving BaseEntry, BaseLine, Price, WhsCode, CardCode

Preferred lookup source for now. A SQL export of open PO lines is placed where the agent can
query it exactly (Dataverse table, SharePoint list, or an Excel table read through a flow
action). Service Layer (`service-layer-po-lookup.md`) replaces this later for both lookup and
posting.

## Export query (SQL Server)

```sql
-- Open PO lines for GRPO lookup; refresh before each receiving batch
SELECT
    CAST(H.DocNum AS varchar(10)) + '|' + L.ItemCode + '|'
      + UPPER(REPLACE(REPLACE(REPLACE(ISNULL(L.U_Color,''),'/',''),'-',''),' ',''))  AS LookupKey,
    H.DocNum          AS PONum,
    H.DocEntry        AS BaseEntry,
    L.LineNum         AS BaseLine,
    H.CardCode,
    H.CardName,
    H.U_POPrefix,
    H.U_CustRefNo,
    H.U_SONo,
    L.ItemCode,
    L.Dscription,
    L.U_Color,
    L.Quantity        AS OrderedQty,
    L.OpenQty,
    L.Price,
    L.U_FOBPrice,
    L.WhsCode,
    H.DocDate         AS PODate,
    GETDATE()         AS ExportedAt
FROM OPOR H
JOIN POR1 L ON L.DocEntry = H.DocEntry
WHERE H.DocStatus = 'O'
  AND L.LineStatus = 'O'
  AND H.CANCELED = 'N'
ORDER BY H.DocNum, L.LineNum;
```

Optional scope: `AND H.CardCode IN (...)` for offshore factory vendors. If `U_Color` or
`U_FOBPrice` do not exist on POR1, drop them and match on ItemCode only (see fallback below).

## Table schema

| Column | Type | Use |
|---|---|---|
| LookupKey | text, unique | `<PONum>\|<ItemCode>\|<COLORKEY>`; the one string the agent searches |
| PONum | int | matches packing-list PO number |
| BaseEntry, BaseLine | int | → PDN1 `BaseEntry`, `BaseLine`, `U_BaseEntry`, `U_BaseLine` |
| CardCode, CardName | text | → OPDN `CardCode`, PDN1 `BaseCard`; verify every PO on an invoice shares one CardCode |
| U_POPrefix, U_CustRefNo, U_SONo | text | header UDFs when the invoice has a single PO; cross-check prefix vs packing PO# |
| ItemCode, Dscription, U_Color | text | Dscription may replace the factory DESCRIPTION |
| OrderedQty, OpenQty | number | over-receipt check: packing qty ≤ OpenQty |
| Price, U_FOBPrice | number | → PDN1 `Price`, `U_FOBPrice` |
| WhsCode | text | → PDN1 `WhsCode` |
| ExportedAt | datetime | snapshot age; agent reports it |

## Where to put it

| Option | Exact lookup? | Notes |
|---|---|---|
| Dataverse table as knowledge | Yes | Structured retrieval; best fit if Dataverse is in use |
| SharePoint list as knowledge | Yes | Same; list columns = schema above |
| Excel table on SharePoint + flow action "List rows present in a table" filtered on LookupKey | Yes | Agent calls the action with the key list; exact |
| Excel/CSV uploaded as a knowledge file | No | Semantic, chunked retrieval; can miss or mis-return rows on long tables. Avoid for key lookups |

## Agent matching procedure

Three shapes of SAP ItemCode appear on PO lines:
- **Custom item, plain C-number**: `C26-03388`. PO line usually has no `U_Color`. The packing
  list's STYLE# before the parenthesis is this code.
- **Custom item with color/size suffix**: `C24-02235-KHA-OS`, `U_Color` `11Khaki`. Packing
  shows `C24-02235` and color `KHA`.
- **Stock / program item**: `<customer style>-<color code>-<size>`, `SP500-INDI-OS`,
  `U_Color` `NAVY`. Packing shows a C-number plus the style in parentheses
  (`C17-06417(SP500)`); the C-number is **not** the SAP ItemCode.

Match order, per packing line (PO = 6-digit number; COLORKEY = color upper-cased, `/ -` and
spaces removed, GREY→GRAY):
1. `LookupKey` = `PO|ItemCode|COLORKEY` → exact hit → OK.
2. `PO + ItemCode` exact, one open line, **no color on the PO line** → OK (custom item, normal).
   Open lines exist in other colors → MISMATCH listing them.
3. `PO + C-number prefix` (SAP ItemCode starts with the packing C-number + `-`): take the line
   whose color segment equals COLORKEY or whose `U_Color` contains it → OK (normal; not logged).
   One SKU with a different color → MISMATCH. Several → MISMATCH listing them.
4. `PO + style prefix` (ItemCode before the first `-` equals the parenthesised style) and
   COLORKEY equals the line's `U_Color` → OK (normal; not logged); write SAP's ItemCode on the
   line, keep the packing C-number in FreeTxt. Style with one colorless line → REVIEW.
   Otherwise MISMATCH listing open lines.
5. Packing qty > `OpenQty` → MISMATCH "over-receipt <qty> vs open <OpenQty>".
6. `U_FOBPrice` > 20 × `Price` or > 50 → REVIEW "FOB price looks wrong" (PO data error).
7. All POs on the invoice must resolve to the same `CardCode`; otherwise MISMATCH.
8. **Vendor comes from the PO, not from the shipper name.** The factory on the documents can be
   a sub-factory of the PO vendor (Unimas ships against POs placed on Mainland Sewing V00313).
   `CardCode`/`CardName` = PO vendor; shipper name goes to Comments. Not logged - this is the
   normal case.
9. Packing PO prefix ≠ `U_POPrefix` on the PO → REVIEW.
10. Report `ExportedAt` in the summary. If older than 24 h, say "lookup snapshot is <n> hours old".

Only exceptions are logged: a line that did not resolve, a color conflict, over-receipt, a
price anomaly, or a vendor conflict between POs. Key translation (C-number → suffixed SKU,
style → stock SKU) and a shipper that differs from the PO vendor are correct behavior and stay
silent.

Never match a packing color to an open line of a different color. A wrong line posts stock
against the wrong SKU and closes a PO line that was never delivered.

In DRI's database `BaseEntry` (DocEntry) equals `PONum` (DocNum) on every open PO line, so a
PO number from the packing list is also its DocEntry. Still take `BaseEntry` from the table.

## Testing against an already-received invoice

The open-PO export excludes POs that are closed, so an old invoice (e.g. 4762, shipped
28 Sept) resolves nothing. For a matching test only, export the lines for those POs regardless
of status, then run with `--lookup` on that file:

```sql
-- Test export: all lines for the POs on one invoice, open or closed
... same SELECT ...
FROM OPOR H JOIN POR1 L ON L.DocEntry = H.DocEntry
WHERE H.DocNum IN (121109, 121749, 121750 /* ... POs from the packing list */)
ORDER BY H.DocNum, L.LineNum;
```

`OpenQty` will be 0 on received lines, so every line reports over-receipt; ignore that status
in this test and look only at whether each packing row found its line and which ItemCode it
matched. For a production test use an invoice that has not been received yet.

## Staleness rule

`OpenQty` falls with every posted receipt. Two invoices against the same PO processed between
refreshes will both validate against the pre-receipt open quantity. Refresh the export before
each receiving batch, or accept that over-receipt detection is approximate until Service Layer
is in place.

## Running the join

```
python scripts/build_opdn.py examples/OPDN_Excel_template.xlsx <out_dir> <factory.xls> --lookup <po_lookup.csv|.xlsx>
```

The export's first sheet (or the CSV) must carry the schema above; `LookupKey` is used when
present, otherwise built from `PONum`, `ItemCode`, `U_Color`. Resolved lines get `BaseEntry`,
`BaseLine`, `U_BaseEntry`, `U_BaseLine`, `Price`, `U_FOBPrice`, `WhsCode`, `BaseCard`,
`LineTotal`, and the PO's `Dscription`; the header gets `CardCode`, `JrnlMemo`, `DocTotal`, and
`U_SONo` on single-PO receipts. Unresolved lines keep `LOOKUP:` cells and appear on
`Lookups_Needed` with the reason. `examples/po_lookup_SAMPLE_fake_values.csv` is a fabricated
table for invoice 4762 (fake DocEntry, price 2.50, vendor V99999) that exercises every path:
one color-spelling fallback (REVIEW), one missing PO (MISMATCH), one over-receipt (MISMATCH).
`examples/with-lookup/` holds the result.
