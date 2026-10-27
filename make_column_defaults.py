"""Emit references/opdn-column-defaults.csv: every template column with its class and value rule."""
import csv, sys, datetime as dt
sys.path.insert(0, __file__.rsplit("/", 1)[0])
from build_opdn import load_template, SYSTEM_HDR, SYSTEM_LINE, OPEN_STATUS

SOURCE_HDR = {
    "DocEntry": "sequence per GRPO in file; links to PDN1_Line.DocEntry",
    "DocDate": "Inv: DATE (DD.MM.YYYY)", "DocDueDate": "= DocDate", "TaxDate": "= DocDate",
    "CardCode": "LOOKUP vendor by Inv: shipper name", "CardName": "Inv: shipper name (row 2)",
    "Address": "Inv: shipper address lines", "Address2": "Inv: Consignee block",
    "NumAtCard": "Inv: Invoice No.", "U_FactoryInv": "Inv: Invoice No.",
    "Comments": "Based On Purchase Order <PO>. Factory invoice <no> dated <date>. EXP No. <exp>. HS <hs>. Contract <contract>.",
    "JrnlMemo": "Goods Receipt PO - <CardCode>",
    "U_OrderNo": "Pac: PO# digits (6)", "U_POPrefix": "Pac: PO# first 3 chars", "U_CustRefNo": "Pac: PO# token after the number",
    "U_FOB": "Inv: Port of Shipment", "U_OceanDest": "Inv: Port of Discharge -> 'Kansas City, KS'",
    "U_BLConsig": "'Design Resources, Inc.'", "U_GenConsig": "'Design Resources, Inc.'",
    "U_COO": "country of shipper ('BD' for Bangladesh)", "U_TermSale": "Pac: TERMS (FOB)", "U_MOT": "Inv: Shipped (Sea)",
    "U_HandoverDate": "Inv/Pac: ICD DATE (confirm UDF)", "U_NumCartons": "sum Pac: CTN for this PO",
    "U_GrossWeight": "sum Pac: Total G.W for this PO", "U_CubeMeters": "sum Pac: CBM for this PO",
    "Weight": "= U_GrossWeight", "Volume": "= U_CubeMeters",
}
SOURCE_LINE = {
    "DocEntry": "parent OPDN_Header.DocEntry", "LineNum": "0..n within GRPO", "VisOrder": "= LineNum",
    "BaseType": "22 (Purchase Order)", "U_BaseType": "22", "BaseRef": "PO number", "BaseDocNum": "PO number",
    "BaseEntry": "LOOKUP PO DocEntry", "BaseLine": "LOOKUP PO line by ItemCode + U_Color",
    "U_BaseEntry": "= BaseEntry", "U_BaseLine": "= BaseLine", "BaseCard": "= header CardCode",
    "ItemCode": "Pac: STYLE# before '('", "Dscription": "Pac: DESCRIPTION (or PO line description)",
    "Quantity": "Pac: Total PCS.", "InvQty": "= Quantity", "PackQty": "= Quantity", "BaseQty": "= Quantity",
    "ShipDate": "= header DocDate", "ActDelDate": "= header DocDate",
    "Price": "LOOKUP PO line Price (not on factory docs)", "U_FOBPrice": "LOOKUP PO line U_FOBPrice",
    "Currency": "'$'", "WhsCode": "LOOKUP default receiving warehouse", "CountryOrg": "'BD'",
    "U_Color": "Pac: COLOR", "U_LinePOPrefix": "= U_POPrefix",
    "U_CartonNo": "Pac: CTN.(S)-CTN.(E)", "U_PcsPerCarton": "Pac: QTY/CTN", "U_NumfCart": "Pac: CTN",
    "U_CartonWeight": "Pac: G.W/CTN", "U_CartonLength": "Pac: LENGTH", "U_CartonWidth": "Pac: WIDTH", "U_CartonHeight": "Pac: HEIGHT",
    "FreeTxt": "PI <PI#> seq <SEQ PI#>; cust style <STYLE# in parens>; cust ref <PO# ref>; mixed packing note",
}

def classify(sheet, col, ex):
    src = (SOURCE_HDR if sheet == "OPDN_Header" else SOURCE_LINE)
    sysset = SYSTEM_HDR if sheet == "OPDN_Header" else SYSTEM_LINE
    if col in src:
        rule = src[col]
        return ("LOOKUP" if "LOOKUP" in rule else "SOURCE"), rule
    if sheet == "OPDN_Header" and col in OPEN_STATUS:
        return "CONST", OPEN_STATUS[col]
    if sheet == "PDN1_Line" and col in ("LineStatus", "InvntSttus"):
        return "CONST", "O"
    if col in sysset:
        return "SYSTEM", "leave NULL - SAP assigns or derives from the PO"
    if ex is None or ex == "NULL":
        return "NULL", "NULL"
    if isinstance(ex, dt.datetime):
        return "SYSTEM", "leave NULL"
    return "CONST", ex

if __name__ == "__main__":
    tpl = load_template(sys.argv[1])
    with open(sys.argv[2], "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Sheet", "Column", "Class", "Value or source rule", "Template example"])
        for sheet, (hdr, ex) in tpl.items():
            for col in hdr:
                cls, rule = classify(sheet, col, ex.get(col))
                w.writerow([sheet, col, cls, rule, ex.get(col)])
    print("ok")
