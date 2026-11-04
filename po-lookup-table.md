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
    I.U_Proto,
    I.U_StyleId,
    I.U_MxPrStyl,
    L.Dscription,
    L.U_Color,
    L.Quantity        AS OrderedQty,
    L.OpenQty,
    L.Price,
    L.U_FOBPrice,
    L.WhsCode,
    L.ShipDate,
    H.DocDate         AS PODate,
    GETDATE()         AS ExportedAt
FROM OPOR H
JOIN POR1 L ON L.DocEntry = H.DocEntry
LEFT JOIN OITM I ON I.ItemCode = L.ItemCode
WHERE H.DocStatus = 'O'
  AND L.LineStatus = 'O'
  AND H.CANCELED = 'N'
ORDER BY H.DocNum, L.LineNum;
```

`LEFT JOIN OITM` because service lines have no ItemCode. `LookupKey` is informational; the
join always rebuilds its keys from `PONum`, `ItemCode`, `U_Proto` and `U_Color`.

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
| U_Proto | text | OITM.U_Proto - the DRI C-number the factory prints as STYLE#; **primary match key** |
| U_StyleId | text | OITM customer style id (`83904TF`); match key when the packing STYLE# is not a C-number (DON prints `#83904TF`) |
| U_MxPrStyl | text | OITM matrix parent style (`SP500`); cross-checked against the packing paren text for SPO/VAC |
| ShipDate | date | POR1 line ship date; tie-breaker when several lines match - the line due on the invoice date wins |
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

Packing STYLE# = `<U_Proto>(<style>)`. The text before the parenthesis is OITM.U_Proto on the
SAP item - for custom items it is also the ItemCode (`C26-03388`, or `C24-02235-KHA-OS` with a
color/size suffix); for stock/program items the ItemCode is `<matrix parent style>-<color
code>-<size>` (`SP500-INDI-OS`) and only U_Proto carries the C-number. The paren text is the
customer style id, except for SPO/VAC where it is the matrix parent style (OITM.U_MxPrStyl).

Match order, per packing row (PO = 6-digit number; COLORKEY = color upper-cased, `/ -` and
spaces removed, GREY→GRAY). A line "matches the color" when COLORKEY equals the line's
`U_Color` **or the color segment of its ItemCode** (`C24-00925-CHR-OS` ↔ packing `CHR`, SAP
`U_Color` `26Charcoal`); failing both, containment of 3+ characters (`KHA` in `11Khaki`).
1. `PO + U_Proto` lines that match the color → OK. One line for the proto with no color →
   OK. Several colorless lines → rule 4a. Lines exist but none match → MISMATCH listing them.
1a. No proto hit: `PO + U_StyleId` (packing STYLE# with any leading `#` stripped, or the
   paren text) lines that match the color → OK. Covers DON and other customers whose POs
   carry the customer style rather than a C-number.
2. No `U_Proto` on the PO lines (old items, or export without OITM): `PO + ItemCode + COLORKEY`
   exact → OK; `PO + ItemCode` one colorless line → OK; other colors → MISMATCH.
3. `PO + C-number prefix` (ItemCode starts with the packing C-number + `-`) with color segment
   or `U_Color` equal to COLORKEY → OK.
4. `PO + style prefix` (ItemCode before the first `-` = paren text) with `U_Color` =
   COLORKEY → OK. One colorless line → REVIEW. Otherwise MISMATCH listing open lines.
4a. **Several open lines in the same color** (same SKU ordered twice on one PO, or a split
   line): **ship date, then quantity.** Exactly one line whose POR1 `ShipDate` equals the
   invoice date → OK, silent. Otherwise the line whose `OpenQty` equals the packing row
   quantity → OK, silent. None exact: a line whose `OrderedQty` equals the row (partially received) →
   REVIEW. None: lowest `LineNum` whose `OpenQty` covers the row → REVIEW naming the
   alternatives. None covers → MISMATCH. Applies on every rule above.
5. **One GRPO line per packing row - never merge.** Two rows can reference the same PO line
   (same style and color packed in two carton configurations). Over-receipt is checked on the
   **sum** of packing qty per `BaseEntry`+`BaseLine` against `OpenQty` → MISMATCH.
6. `U_FOBPrice` > 20 × `Price` or > 50 → REVIEW "FOB price looks wrong" (PO data error).
7. SPO/VAC paren text ≠ OITM.U_MxPrStyl (after allowing a prefix, `6440` = `VC6440`) → REVIEW.
7a. Nothing above matched and the PO has exactly one open line whose `OpenQty` covers the row
   (customer styles printed without a C-number, e.g. `#83904TF`) → take it, REVIEW.
8. All POs on the invoice must resolve to the same `CardCode`; otherwise MISMATCH.
9. Vendor comes from the PO, not from the shipper name (Unimas ships against POs placed on
   Mainland Sewing V00313). `CardCode`/`CardName` = PO vendor; shipper to Comments. Not logged.
10. Packing PO prefix ≠ `U_POPrefix` on the PO → REVIEW.
11. Report `ExportedAt` in the summary. If older than 24 h, say "lookup snapshot is <n> hours old".

Only exceptions are logged: a row that did not resolve, a color conflict, over-receipt, a price
anomaly, a style cross-check failure, or a vendor conflict between POs. Key translation and a
shipper that differs from the PO vendor are correct behavior and stay silent.

Never match a packing color to an open line of a different color. A wrong line posts stock
against the wrong SKU and closes a PO line that was never delivered.

Multiple GRPO lines drawing from one PO line is standard DI API / Service Layer behavior as
long as the total does not exceed the open quantity. Confirm once in the test company.

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
python scripts/build_opdn.py "<factory.xls>" --lookup "<po-lookup-table.xlsx>"     # output in ./out/
```

The export's first sheet (or the CSV) must carry the schema above; `LookupKey` is used when
present, otherwise built from `PONum`, `ItemCode`, `U_Color`. Resolved lines get `BaseEntry`,
`BaseLine`, `U_BaseEntry`, `U_BaseLine`, `Price`, `U_FOBPrice`, `WhsCode`, `BaseCard`,
`LineTotal`, and the PO's `Dscription`; the header gets `CardCode`, `JrnlMemo`, `DocTotal`, and
`U_SONo` on single-PO receipts. Unresolved lines keep `LOOKUP:` cells and appear on
`Lookups_Needed` with the reason. `examples/po_lookup_SAMPLE_fake_values.csv` is a fabricated
table for invoice 4762 (fake DocEntry, price 2.50, vendor V99999) that exercises every path:
a GREY/GRAY spelling (silent match), one missing PO (MISMATCH), one over-receipt (MISMATCH).
`examples/OPDN_upload_EX_UNI_4762_2026_with-lookup.xlsx` is the result.
