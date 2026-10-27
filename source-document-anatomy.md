# Factory Invoice / Packing List - document anatomy

Observed format: Unimas Sportswear Ltd. (Bangladesh) `.xls`, filename pattern
`Invoice__Packing_List_-_<4-digit seq>__Caps_Direct_-_DD.MM.YYYY_-_Sea_USA.xls`.
Two worksheets. Other factories will differ in layout; the field meanings below still apply.

## Sheet `Inv-` (commercial/shipping invoice) - header fields only

| Field | Where (row, col - 0-based) | Example | Notes |
|---|---|---|---|
| Shipper name | r1 c1 | `UNIMAS SPORTSWEAR LTD.` | → `CardName`; key for `CardCode` lookup |
| Shipper address | r2 c1 + r3 c1 | `Baghbari,Gazipur Sadar, Gazipur` / `Bangladesh.` | → `Address`; country → `U_COO` |
| Invoice No. | r5: label c1, value c3 | ` EX/UNI/4762/2026` | trim; → `NumAtCard`, `U_FactoryInv` |
| Invoice date | r5 c6 `DATE : 28.09.2026` | 28.09.2026 | **DD.MM.YYYY** - never parse as US |
| Contract No. | r5: label c11, value c12 | `US260008ST` | → Comments |
| HS Code | r6 c7 `HS Code: 650500` | 650500 | → Comments (no UDF identified) |
| Port of Shipment | r7 c3 | `CHITTAGONG, BANGLADESH` | → `U_FOB` |
| Port of Discharge | r7 c8 | `KANSAS CITY, USA` | → `U_OceanDest` = `Kansas City, KS` |
| Package (cartons) | r8 c3 | 405 | reconcile vs packing list |
| EXP. No. | r9 c3 | `3707-001682-2026` | Bangladesh export permit → Comments |
| ICD DATE | r9 c7 | 29.09.2026 | inland depot handover; → `U_HandoverDate` (confirm) |
| Shipped | r9 c15 | `By Sea` | → `U_MOT` |
| Consignee / Ship-to | r11-r14 c1 and c6 | Design Resources Inc, 7007 College Blvd... | → `Address2` |
| Line block | r17 onward: qty c3, PO# c5, STYLE c7 | | **Do not use as line source** - no color, no cartons. Use `Pac-`. |
| Total G.W / N.W | r82-83 c6 | 6825.98 / 5121.6 KGS | reconcile |
| Total pcs | r83 c3 | 58260 | reconcile |

Everything from ~r90 down (`Amount`, `Words`, `Data Table`, number-to-words grid, "Janet, the maximum amount…")
is spreadsheet plumbing for the amount-in-words formula. Ignore it. **There are no unit prices or
amounts on this invoice** (Amount = 0). Prices come from the SAP PO.

## Sheet `Pac-` (packing list) - the line source

Header fields in rows 1-6 repeat the invoice fields (INVOICE NO, DATE, PORT OF LOADING, PORT OF
DISCHARGE, SHIP MODE, TERMS: FOB, PAYMENT TERMS: COLLECT, HS CODE, ICD DATE). Use them as a
cross-check against `Inv-`; they must agree.

Table header is the row whose first cell is `CTN. (S)`. Columns, in order:

| Col | Header | Meaning | Target |
|---|---|---|---|
| 0 | CTN. (S) | first carton no. | `U_CartonNo` start |
| 1 | CTN. (E) | last carton no. | `U_CartonNo` end |
| 2 | CTN | carton count | `U_NumfCart`, sum → `U_NumCartons` |
| 3 | QTY/ CTN | pieces per carton | `U_PcsPerCarton` |
| 4 | PI# | factory proforma invoice | `FreeTxt` |
| 5 | SEQ PI# | PI sequence/line | `FreeTxt` |
| 6 | PO# | `<PREFIX><6-digit PO>  <customer ref>` | `U_POPrefix`, `U_OrderNo`, `U_CustRefNo`, `BaseRef` |
| 7 | STYLE# | `<ItemCode>(<customer style>)` | `ItemCode`; parens → `FreeTxt` |
| 8 | COLOR | color name as factory writes it | `U_Color`; match to PO line |
| 9 | DESCRIPTION | fabric content + type | `Dscription` |
| 10 | Total PCS. | line quantity | `Quantity` |
| 11 | G.W/CTN | kg per carton | `U_CartonWeight` |
| 12 | N.W/ CTN | kg | - |
| 13 | Total G.W | kg | sum → `U_GrossWeight`, `Weight` |
| 14 | Total N.W | kg | - |
| 15-17 | LENGTH / WIDTH / HEIGHT | cm | `U_CartonLength/Width/Height` |
| 18 | CBM | m³ | sum → `U_CubeMeters`, `Volume` |

The table ends at the first row with a blank `CTN. (S)`; that row carries the totals
(CTN, Total PCS, Total G.W, Total N.W, CBM). The next row repeats them with labels.

## Parsing rules

- **PO#**: collapse whitespace, then `^([A-Z0-9]{3})(\d{6})\s*(.*)$`.
  `SPO120599  91690` → prefix `SPO`, PO `120599`, cust ref `91690`.
  `3BM121109  371331A` → `3BM`, `121109`, `371331A`. The ref may be absent or contain letters.
  Prefixes seen: 3BM ANL BDA DON GLD MCK MPC NRS SPP TIG TSI WPM SPO VAC. Example GRPO used `ABM`.
- **STYLE#**: text before `(` is a DRI C-number (`C17-06417`, `C4-2134`) or a full code
  (`DON-1400CT-000`); text inside parens is the customer's style (`SP500`, `2004975560001`,
  `AUD4000.BK`). For custom items the C-number is the SAP ItemCode. For stock/program items SAP's
  ItemCode is `<style>-<color code>-<size>` (`SP500-PACBL-OS`) and the C-number will not match -
  see the match order in `po-lookup-table.md`.
- **Dates**: all `DD.MM.YYYY`. `04.05.2026` is 4 May, not 5 April.
- **Colors**: free text, inconsistent case (`White`, `GREY-WHITE`, `Charcoal/Black`). Compare case-insensitively
  and tolerate `/`, `-`, spaces when matching PO lines.
- **One PO may span several rows**: different colors (distinct PO lines) or the same color split across
  carton configurations (`3x144 + 1x72` - same PO line, merge).
- **One PO may span several invoices** (`SPO120599` on 4762 and 4763; `VAC120601` on 4763 and 4764).
  Each invoice produces its own partial GRPO against that PO.
