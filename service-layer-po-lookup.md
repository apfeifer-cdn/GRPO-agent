# Resolving OPOR / POR1 lookups through SAP B1 Service Layer from Power Automate

Scope: the agent has PO number, ItemCode and color from the packing list. It needs the PO's
`DocEntry`, the matching line's `LineNum`, `Price`, `WarehouseCode`, open quantity, and the
vendor `CardCode`. Service Layer exposes OPOR as `PurchaseOrders` and POR1 as its
`DocumentLines` collection; UDFs are exposed by their `U_` names.

## 1. Connectivity - decide this first

| B1 hosting | Path from Power Automate | Notes |
|---|---|---|
| On-prem, not internet-facing | **Custom connector + on-premises data gateway** | Gateway installed on a server that can reach `https://<b1host>:50000`. The HTTP action cannot use the gateway; a custom connector can. |
| On-prem, published through a reverse proxy / Azure App Proxy / APIM | HTTP action or custom connector, direct | Needs a public-CA certificate. The HTTP action rejects self-signed certs with no override. |
| Hosted / B1 Cloud | HTTP action, direct | Same cert rule. |

Service Layer base URL: `https://<host>:50000/b1s/v1/` (OData v3 style) or `/b1s/v2/` (OData v4).
Examples below use v1. Create a dedicated B1 user for the flow with authorization to Purchase
Orders (read) and Goods Receipt PO (add) and nothing else. Store the password in a Power Automate
environment variable marked secret, or Azure Key Vault.

## 2. Session handling

Login returns a session cookie; every call carries it. Default timeout 30 min.

```
POST /b1s/v1/Login
{"CompanyDB":"<SBO_DRI>","UserName":"<svc_user>","Password":"<secret>"}
→ 200 {"SessionId":"<guid>","Version":"...","SessionTimeout":30}
   Set-Cookie: B1SESSION=<guid>; ROUTEID=.node1
```

In Power Automate: `outputs('Login')['headers']['Set-Cookie']` contains both cookies. Build the
request header once:

```
Cookie: B1SESSION=@{body('Login')?['SessionId']}; ROUTEID=.node1
```

If the Service Layer is load-balanced, copy `ROUTEID` from the Set-Cookie header instead of
hard-coding `.node1`. Call `POST /b1s/v1/Logout` at the end of the flow (also in the failure
branch) so sessions do not pile up against the license.

## 3. Read the PO

Filter on DocNum (what the factory prints), not DocEntry (what SAP needs).

```
GET /b1s/v1/PurchaseOrders?$filter=DocNum eq 120599
    &$select=DocEntry,DocNum,CardCode,CardName,DocumentStatus,DocDate,U_POPrefix,U_CustRefNo,U_SONo,DocumentLines
Prefer: odata.maxpagesize=1
```

Response (trimmed):

```json
{"value":[{"DocEntry":58213,"DocNum":120599,"CardCode":"V00410","CardName":"UNIMAS SPORTSWEAR LTD.",
  "DocumentStatus":"bost_Open","U_POPrefix":"SPO","U_CustRefNo":"91690",
  "DocumentLines":[
    {"LineNum":0,"ItemCode":"C17-06417","ItemDescription":"...","Quantity":20160,"RemainingOpenQuantity":11520,
     "Price":2.85,"WarehouseCode":"01","LineStatus":"bost_Open","U_Color":"NAVY","U_FOBPrice":2.85},
    {"LineNum":1,"ItemCode":"C17-06417","Quantity":3024,"RemainingOpenQuantity":3024,"Price":2.85,"U_Color":"AQUA", ...}
  ]}]}
```

Nested `DocumentLines` are returned in full; paging applies only to the top-level collection.
Checks on the response:
- `value` empty → PO not found → MISMATCH.
- `DocumentStatus` ≠ `bost_Open` → closed/cancelled PO → MISMATCH.
- `CardCode` must equal the CardCode resolved for the shipper (see §4).

## 4. Resolve the vendor once per invoice

```
GET /b1s/v1/BusinessPartners?$filter=CardType eq 'cSupplier' and CardName eq 'UNIMAS SPORTSWEAR LTD.'
    &$select=CardCode,CardName,U_ActManuf
```

Factory spelling drifts. Safer: take `CardCode` from the first PO on the invoice and verify every
other PO on that invoice has the same `CardCode`. Keep a small alias table
(shipper name → CardCode) in a SharePoint list or Dataverse table for names that do not match.

## 5. Match the packing row to a PO line

Match key: `ItemCode` exact, `U_Color` normalized (upper-case, strip `/`, `-`, spaces).
Power Automate: `Filter array` on `body('Get_PO')?['value'][0]?['DocumentLines']` with

```
and(
  equals(item()?['ItemCode'], variables('ItemCode')),
  equals(toUpper(replace(replace(replace(coalesce(item()?['U_Color'],''),'/',''),'-',''),' ','')), variables('ColorKey'))
)
```

Then:
- 0 matches → try ItemCode only; if exactly one open line, take it and flag REVIEW (color spelled
  differently). Otherwise MISMATCH.
- >1 match → MISMATCH (ambiguous; needs a human).
- 1 match → return `LineNum`, `Price`, `U_FOBPrice`, `WarehouseCode`, `RemainingOpenQuantity`.
  If packing `Quantity > RemainingOpenQuantity` → MISMATCH (over-receipt), with both numbers.

## 6. Flow shape

One flow, called by the Copilot Studio agent as an action, takes the whole invoice in one call
to avoid one login per line.

**Trigger:** "Run a flow from Copilot" - input `lines` (string, JSON array):
`[{"po":120599,"item":"C17-06417","color":"NAVY","qty":11520}, ...]`

**Steps**
1. `Parse JSON` on `lines`.
2. `HTTP` Login → `Compose` cookie header.
3. `Select` distinct PO numbers (`union()` over the parsed array).
4. `Apply to each` PO (concurrency 1 to keep the session simple):
   `HTTP` GET PurchaseOrders by DocNum → `Parse JSON` → append `{po, DocEntry, CardCode, status, lines[]}` to an array variable.
5. `Apply to each` input line: pick the PO record, `Filter array` lines per §5, append result
   `{po, item, color, docEntry, lineNum, price, fobPrice, whs, openQty, cardCode, status, note}`.
6. `HTTP` Logout (configure "run after" to include failures).
7. "Respond to Copilot" - output `result` (string, JSON array of step-5 records).

The agent writes `docEntry`/`lineNum` into `BaseEntry`/`BaseLine`, `price` into `Price`,
`whs` into `WhsCode`, `cardCode` into `CardCode`/`BaseCard`, and replaces every `LOOKUP:` cell.
Lines with `status` ≠ `ok` go to the Validation sheet as MISMATCH/REVIEW.

Retry: set the HTTP actions' retry policy to exponential, 3 tries. Service Layer returns
`401` with `"Invalid session"` when the cookie expires - on 401 re-run Login and retry once.

## 7. Phase 2 - post the GRPO directly, skip the spreadsheet

Once the lookups come from Service Layer, the same session can create the document:

```
POST /b1s/v1/PurchaseDeliveryNotes
{
 "CardCode":"V00410","DocDate":"2026-09-28","DocDueDate":"2026-09-28","TaxDate":"2026-09-28",
 "NumAtCard":"EX/UNI/4762/2026","Comments":"Factory invoice EX/UNI/4762/2026 ... EXP No. 3707-001682-2026. HS 650500.",
 "U_FactoryInv":"EX/UNI/4762/2026","U_FOB":"Chittagong, Bangladesh","U_OceanDest":"Kansas City, KS",
 "U_COO":"BD","U_NumCartons":80,"U_GrossWeight":1296.0,"U_CubeMeters":10.3958,
 "DocumentLines":[
   {"BaseType":22,"BaseEntry":58213,"BaseLine":0,"Quantity":11520,"WarehouseCode":"01",
    "U_Color":"NAVY","U_CartonNo":"326-405","U_PcsPerCarton":144,"U_NumfCart":80,"U_CartonWeight":16.2,
    "U_CartonLength":59.5,"U_CartonWidth":42,"U_CartonHeight":52,
    "FreeText":"PI BA1089093 seq 7; cust style SP500; cust ref 91690"}
 ]
}
```

Rules that still apply: one document per PO (header UDFs are single-PO); omit `Price` so SAP
copies it from the PO line; omit `U_OrderNo`/`U_POPrefix`/`U_CustRefNo` if SAP's existing
copy-from-PO logic fills them (test one); `DocEntry`, totals, landed-cost UDFs are not sent.
Use `Prefer: return-no-content` to skip the echo. Keep the human review step before the POST:
the agent produces the review workbook and Validation sheet, a person approves, the flow posts.
Test against the test company DB first; a GRPO posts inventory and cannot be deleted, only
cancelled or returned.

## 8. Service Layer field map (OPDN table name → Service Layer property)

| Table column | Service Layer | | Table column | Service Layer |
|---|---|---|---|---|
| CardCode | CardCode | | ItemCode | ItemCode |
| DocDate / DocDueDate / TaxDate | DocDate / DocDueDate / TaxDate | | Dscription | ItemDescription |
| NumAtCard | NumAtCard | | Quantity | Quantity |
| Comments | Comments | | WhsCode | WarehouseCode |
| JrnlMemo | JournalMemo | | BaseType / BaseEntry / BaseLine | BaseType / BaseEntry / BaseLine |
| DocCur | DocCurrency | | FreeTxt | FreeText |
| Series | Series | | CountryOrg | CountryOrg |
| U_* (header) | U_* | | U_* (line) | U_* |
