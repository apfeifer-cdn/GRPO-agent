# OPDN / PDN1 field mapping

Output workbook mirrors `OPDN_Excel.xlsx`: sheet `OPDN_Header` (715 columns, OPDN table) and
sheet `PDN1_Line` (459 columns, PDN1 table), same column names, same order, one header row
per GRPO and one line row per PO line received. Two extra sheets are appended for the reviewer:
`Validation` and `Lookups_Needed`. Full per-column rules are in `opdn-column-defaults.csv`
(classes: SOURCE, LOOKUP, CONST, SYSTEM, NULL).

## Document structure rule

**One GRPO per factory invoice** (decided). DRI ships at container level: one invoice is one
container consolidating many POs, so one `OPDN_Header` row (`DocEntry` 1) with every packing
row as a line linked to its own PO line through `BaseEntry`/`BaseLine`. SAP B1 allows a GRPO to
draw from several POs of the same vendor. The header UDFs `U_OrderNo`, `U_POPrefix`,
`U_CustRefNo`, `U_SONo` are single-PO fields and stay `NULL`; `Comments` lists the POs. A
single-PO invoice is the rare case and fills them. The per-line fields `BaseRef`,
`U_LinePOPrefix` and `FreeTxt` (cust ref) keep the PO identity at line level.
`scripts/build_opdn.py --group po` produces the per-PO alternative if that is ever needed.

The lean sheets `GRPO_Header` / `GRPO_Lines` carry the same values restricted to the columns
that hold data (SOURCE + LOOKUP + the few CONSTs a reviewer cares about). They are the form a
person reviews and the form a Service Layer POST is built from.

## Header (OPDN_Header) - fields that carry data

| Column | Value | Class |
|---|---|---|
| DocEntry | sequence 1..n within the file (link key to lines) | SOURCE |
| DocType / ObjType | `I` / `20` | CONST |
| DocStatus / InvntSttus / Printed / Posted | `O` / `O` / `N` / `N` (new open doc; template row was a closed historical doc) | CONST |
| DocDate, DocDueDate, TaxDate | factory invoice date | SOURCE |
| CardCode | vendor code for shipper name - **lookup** (OCRD, CardType S) | LOOKUP |
| CardName | shipper name | SOURCE |
| Address | shipper address | SOURCE |
| Address2 | consignee block (DRI, 7007 College Blvd…) | SOURCE |
| NumAtCard, U_FactoryInv | factory invoice no. (`EX/UNI/4762/2026`) | SOURCE |
| DocCur / DocRate / SysRate / CurSource | `$` / 1 / 1 / `L` | CONST |
| Comments | `Based On Purchase Order <PO>. Factory invoice <no> dated <MM/DD/YYYY>. EXP No. <exp>. HS <hs>. Contract <contract>.` | SOURCE |
| JrnlMemo | `Goods Receipt PO - <CardCode>` | SOURCE |
| U_OrderNo / U_POPrefix / U_CustRefNo | from PO# token | SOURCE |
| U_FOB | port of shipment, title case (`Chittagong, Bangladesh`) | SOURCE |
| U_OceanDest | `Kansas City, KS` | SOURCE |
| U_BLConsig, U_GenConsig | `Design Resources, Inc.` | CONST |
| U_COO | ISO-2 of shipper country (`BD`) | SOURCE |
| U_TermSale / U_MOT | `FOB` / `Sea` | SOURCE |
| U_HandoverDate | ICD date (confirm this is the intended UDF) | SOURCE |
| U_NumCartons, U_GrossWeight, U_CubeMeters | sums over this PO's lines | SOURCE |
| Weight / WeightUnit, Volume / VolUnit | same sums / 3 / 4 (template units) | SOURCE |
| U_SONo, U_ActManuf, U_BrokerRef, U_ContREfNo, U_Notify, U_ActShipDt, U_OrigShipDt, CntctCode, Series, GroupNum, PeyMethod, PayToCode, CtlAccount | come from the PO / BP master or the forwarder - NULL | SYSTEM |
| DocNum, TransId, DocTotal*, Paid*, CreateDate, UpdateDate, UserSign*, DocTime, StationID, FinncPriod, Ref1, VersionNum | SAP assigns | SYSTEM |
| All other Y/N flags and zeros | carried from the template example | CONST |
| Everything `NULL` in the template | `NULL` | NULL |

Open items on the header: `U_MakerTerms` (template `CDN`, meaning unknown - left NULL);
`U_PackListRemarks` / `U_CustComm` are candidates for EXP no. and HS code if Comments is not
the right home.

## Lines (PDN1_Line)

| Column | Value | Class |
|---|---|---|
| DocEntry | parent header DocEntry | SOURCE |
| LineNum, VisOrder | 0..n within the GRPO | SOURCE |
| BaseType, U_BaseType | 22 | CONST |
| BaseRef, BaseDocNum | PO number (DocNum) | SOURCE |
| BaseEntry, BaseLine, U_BaseEntry, U_BaseLine | PO DocEntry and the PO line matched on ItemCode + U_Color - **lookup** (OPOR/POR1) | LOOKUP |
| BaseCard | header CardCode | LOOKUP |
| ItemCode | STYLE# before `(` | SOURCE |
| Dscription | packing DESCRIPTION; prefer the PO line description once looked up | SOURCE |
| Quantity, InvQty, PackQty, BaseQty | Total PCS. | SOURCE |
| ShipDate, ActDelDate | header DocDate | SOURCE |
| Price, U_FOBPrice | PO line price - **lookup**; not on factory docs | LOOKUP |
| Currency | `$` | CONST |
| WhsCode | default receiving warehouse (template `1`) - **confirm** | LOOKUP |
| LineStatus / InvntSttus | `O` | CONST |
| CountryOrg | `BD` | SOURCE |
| U_Color | COLOR | SOURCE |
| U_LinePOPrefix | header U_POPrefix | SOURCE |
| U_CartonNo | `<CTN.(S)>-<CTN.(E)>` | SOURCE |
| U_PcsPerCarton, U_NumfCart, U_CartonWeight | QTY/CTN, CTN, G.W/CTN | SOURCE |
| U_CartonLength/Width/Height | cm | SOURCE |
| FreeTxt | `PI <PI#> seq <n>; cust style <x>; cust ref <y>[; mixed packing 3x144 + 1x72]` | SOURCE |
| UomCode/UomCode2 `Manual`, Factor1-4 = 1, NumPerMsr = 1, TreeType `N`, ItemType 4, etc. | template | CONST |
| Landed-cost UDFs (U_Duty, U_MPF, U_BrokerFee, U_Freight, U_LocalWH, U_QCFee, U_Insurance, U_Finance, U_FixedCharge, U_Incremental, U_Marketing…) | copied from PO line by SAP - NULL | SYSTEM |
| LineTotal, OpenQty, OpenSum*, StockSum*, GTotal*, AcctCode, CogsAcct, ShipToCode/Desc, TargetType/TrgetEntry | SAP computes / derives | SYSTEM |

### Merge rule
Packing rows with the same PO + ItemCode + color are one PO line. Merge: sum Quantity, cartons,
weight, CBM; `U_CartonNo` = min start - max end; `U_PcsPerCarton` = the larger pack; note the
split in FreeTxt and flag REVIEW in Validation.

### Lookup placeholders
Where the agent has no SAP access, write `LOOKUP:<key>` in the cell, e.g.
`LOOKUP:CardCode[UNIMAS SPORTSWEAR LTD.]`, `LOOKUP:PO121109[C26-06700|BLACK].LineNum`,
`LOOKUP:default_whs`, and list every one on `Lookups_Needed`. A file containing any `LOOKUP:`
is not uploadable until resolved.

## Importer caveat
`OPDN_Excel.xlsx` is a straight dump of the OPDN and PDN1 tables, not a DTW
(`oPurchaseDeliveryNotes` Documents / Document_Lines) template. If the uploader is DTW or Service
Layer, map the SOURCE/LOOKUP columns above onto that template's names (`RecordKey` = DocEntry,
`CardCode`, `DocDate`, `NumAtCard`, `Comments`, lines: `RecordKey`, `LineNum`, `ItemCode`,
`Quantity`, `Price`, `WarehouseCode`, `BaseType`, `BaseEntry`, `BaseLine`, `U_…`) and drop the
SYSTEM/CONST columns. Confirm with the SAP owner which importer consumes the file before go-live.
