"""Reference implementation: factory Invoice/Packing List (.xls) -> OPDN upload workbook.

Usage: python build_opdn.py <template OPDN_Excel.xlsx> <out_dir> <factory1.xls> [...] [--group invoice|po] [--lookup po_lookup.csv|.xlsx]
--lookup: OPOR/POR1 export per references/po-lookup-table.md; resolves BaseEntry/BaseLine/Price/U_FOBPrice/WhsCode/CardCode.
Default: one GRPO per invoice (header row + all lines). --group po: one GRPO per DRI PO.
Sheets: OPDN_Header / PDN1_Line (full template shape), GRPO_Header / GRPO_Lines (lean staging), Validation, Lookups_Needed.
Fields that need SAP lookups are written as LOOKUP:<key>; system-assigned fields are left NULL.
"""
import re, sys, datetime as dt
from collections import OrderedDict, defaultdict
import xlrd, openpyxl

NULL = "NULL"

# ---------- parsing ----------
def cells(sheet, r):
    return {c: sheet.cell_value(r, c) for c in range(sheet.ncols) if sheet.cell_value(r, c) not in ("", None)}

def parse_date(s):
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(s))  # factory dates are DD.MM.YYYY
    return dt.datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None

def after_colon(s):
    return str(s).split(":", 1)[1].strip() if ":" in str(s) else str(s).strip()

def parse_po(s):
    # "SPO120599  91690" -> prefix SPO, po 120599, custref 91690
    s = " ".join(str(s).split())
    m = re.match(r"^([A-Z0-9]{3})(\d{6})\s*(.*)$", s)
    if not m:
        return None, None, s
    return m.group(1), int(m.group(2)), m.group(3).strip() or None

def parse_style(s):
    # "C17-06417(SP500)" -> item C17-06417, cust style SP500
    s = str(s).strip()
    m = re.match(r"^([^()]+?)\s*(?:\((.+)\))?$", s)
    return m.group(1).strip(), (m.group(2) or "").strip() or None

def parse_factory(path):
    wb = xlrd.open_workbook(path)
    inv = next(s for s in wb.sheets() if s.name.lower().startswith("inv"))
    pac = next(s for s in wb.sheets() if s.name.lower().startswith("pac"))
    h = {"file": path.split("/")[-1]}
    h["vendor_name"] = str(inv.cell_value(1, 1)).strip()
    h["vendor_addr"] = f"{inv.cell_value(2,1)} {inv.cell_value(3,1)}".strip()
    for r in range(0, 20):
        row = cells(inv, r)
        for c, v in row.items():
            t = str(v)
            if t.startswith("Invoice No"):
                h["invoice_no"] = str(row.get(c + 2, "")).strip()
            if t.startswith("DATE"):
                h["invoice_date"] = parse_date(t)
            if t.startswith("Contract No"):
                h["contract_no"] = str(row.get(c + 1, "")).strip()
            if t.startswith("HS Code"):
                h["hs_code"] = after_colon(t)
            if t.startswith("Port of Shipment"):
                h["port_load"] = str(row.get(c + 2, "")).strip()
            if t.startswith("Port of Discharge"):
                h["port_discharge"] = str(row.get(c + 1, "")).strip()
            if t.startswith("Package"):
                h["cartons"] = int(row.get(c + 2, 0))
            if t.startswith("EXP. No"):
                h["exp_no"] = str(row.get(c + 2, "")).strip()
            if t.startswith("ICD"):
                h["icd_date"] = parse_date(t)
            if t.startswith("Shipped"):
                h["ship_mode"] = str(row.get(c + 1, "")).strip()
            if t.startswith("Consignee"):
                h["consignee"] = " ".join(str(inv.cell_value(r + k, c)).strip() for k in (1, 2, 3))
    # invoice totals block
    for r in range(inv.nrows):
        row = cells(inv, r)
        for c, v in row.items():
            if str(v).startswith("Total G.W"):
                h["inv_gw"] = float(row.get(c + 1, 0))
            if str(v).startswith("Total N.W"):
                h["inv_nw"] = float(row.get(c + 1, 0))
                h["inv_pcs"] = float(row.get(c - 2, 0))
    # packing table
    hdr_r = next(r for r in range(pac.nrows) if str(pac.cell_value(r, 0)).startswith("CTN. (S)"))
    cols = {str(pac.cell_value(hdr_r, c)).strip(): c for c in range(pac.ncols)}
    lines = []
    for r in range(hdr_r + 1, pac.nrows):
        row = cells(pac, r)
        if not row:
            continue  # blank spacer rows
        if 0 not in row or not isinstance(row[0], (int, float)):  # totals row has blank CTN.(S)
            if 10 in row and isinstance(row[10], (int, float)):
                h["pac_pcs"], h["pac_ctn"] = float(row.get(10, 0)), float(row.get(2, 0))
                h["pac_gw"], h["pac_nw"], h["pac_cbm"] = float(row.get(13, 0)), float(row.get(14, 0)), float(row.get(18, 0))
                break
            if str(row.get(0, "")).startswith("Total PCS"):  # labelled totals row (fallback)
                h["pac_pcs"], h["pac_ctn"] = float(row.get(2, 0)), float(row.get(6, 0))
                h["pac_gw"], h["pac_cbm"], h["pac_nw"] = float(row.get(8, 0)), float(row.get(12, 0)), h.get("pac_nw", 0)
                break
            continue
        g = lambda k: row.get(cols[k])
        prefix, po, custref = parse_po(g("PO#"))
        item, cust_style = parse_style(g("STYLE#"))
        lines.append(dict(
            ctn_s=int(g("CTN. (S)")), ctn_e=int(g("CTN. (E)")), ctn=int(g("CTN")), qty_ctn=int(g("QTY/ CTN")),
            pi=str(g("PI#")).strip(), seq=int(g("SEQ PI#") or 0), po_raw=" ".join(str(g("PO#")).split()),
            prefix=prefix, po=po, custref=custref, item=item, cust_style=cust_style,
            color=str(g("COLOR")).strip(), desc=str(g("DESCRIPTION")).strip(), qty=int(g("Total PCS.")),
            gw_ctn=float(g("G.W/CTN")), nw_ctn=float(g("N.W/ CTN")), gw=float(g("Total G.W")), nw=float(g("Total N.W")),
            L=float(g("LENGTH")), W=float(g("WIDTH")), H=float(g("HEIGHT")), cbm=float(g("CBM"))))
    h["lines"] = lines
    for k in ("pac_pcs", "pac_ctn", "pac_gw", "pac_nw", "pac_cbm"):
        h.setdefault(k, None)  # missing totals row -> reconciliation reports MISMATCH, not a crash
    return h

# ---------- template ----------
SYSTEM_HDR = {"DocNum","TransId","DocTime","CreateDate","UpdateDate","CreateTS","UpdateTS","UserSign","UserSign2",
    "StationID","FinncPriod","LogInstanc","VersionNum","Instance","Ref1","ReceiptNum","OwnerCode","Max1099",
    "DocTotal","DocTotalSy","DocTotalFC","PaidToDate","PaidSys","PaidFC","ImportEnt","SumAbsId","LndCstNum",
    "VatJENum","ReqType","CtlAccount","CntctCode","U_SONo","U_ActManuf","U_BrokerRef","U_ContREfNo","U_Notify",
    "U_OrigShipDt","U_ActShipDt","PayToCode","Series","GroupNum","PeyMethod","U_AllowFreightForwarderUpdate"}
SYSTEM_LINE = {"TargetType","TrgetEntry","OpenQty","OpenSum","OpenSumFC","OpenSumSys","AcctCode","CogsAcct",
    "FinncPriod","LogInstanc","OwnerCode","DocDate","StockSum","StockSumSc","StckSumApp","StckAppSc","GTotal",
    "GTotalSC","TotalSumSy","LineTotal","INMPrice","PriceBefDi","PriceAfVAT","GPBefDisc","ShipToCode","ShipToDesc",
    "BaseOpnQty","OpenInvQty","SlpCode","U_FOBPrice","U_Duty","U_MPF","U_BrokerFee","U_Freight","U_LocalWH",
    "U_QCFee","U_Insurance","U_Finance","U_FixedCharge","U_Incremental","U_Marketing","U_MiscDuty","U_LDP",
    "U_DRIGM","U_AgentCommLDP","U_MiscExp","U_DomHT","U_DRIGMLDP","U_KidLT","U_CamoRoyalty","U_CamoDuty",
    "U_Cunsump","U_Rebate","U_AgentComm","U_SpecialTariff","U_CustRetailPrice","Price","BaseEntry","BaseLine",
    "U_BaseEntry","U_BaseLine","BaseCard","WhsCode"}
OPEN_STATUS = {"DocStatus": "O", "InvntSttus": "O", "Printed": "N", "Posted": "N", "Confirmed": "Y"}

def load_template(path):
    wb = openpyxl.load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        hdr = [c.value for c in ws[1] if c.value is not None]
        ex = {h: ws.cell(2, i + 1).value for i, h in enumerate(hdr)}
        out[ws.title] = (hdr, ex)
    return out

def header_row(h, po_lines, docentry, hdr, ex):
    pos = sorted({l["po"] for l in po_lines})
    single = len(pos) == 1
    po = po_lines[0]
    row = OrderedDict()
    for col in hdr:
        v = ex.get(col)
        if col in SYSTEM_HDR: v = NULL
        if col in OPEN_STATUS: v = OPEN_STATUS[col]
        row[col] = v
    d = h["invoice_date"]
    card = f"LOOKUP:CardCode[{h['vendor_name']}]"
    row.update({
        "DocEntry": docentry, "DocDate": d, "DocDueDate": d, "TaxDate": d,
        "CardCode": card, "CardName": h["vendor_name"], "Address": h["vendor_addr"], "Address2": h["consignee"],
        "NumAtCard": h["invoice_no"], "U_FactoryInv": h["invoice_no"],
        "Comments": (f"Based On Purchase Order{'s' if not single else ''} {', '.join(map(str, pos))}. "
                     f"Factory invoice {h['invoice_no']} dated {d:%m/%d/%Y}. "
                     f"EXP No. {h.get('exp_no')}. HS {h.get('hs_code')}. Contract {h.get('contract_no')}."),
        "JrnlMemo": f"Goods Receipt PO - {card}",
        "U_OrderNo": po["po"] if single else NULL, "U_POPrefix": po["prefix"] if single else NULL,
        "U_CustRefNo": (po["custref"] or NULL) if single else NULL,
        "U_FOB": h["port_load"].title(), "U_OceanDest": "Kansas City, KS",
        "U_BLConsig": "Design Resources, Inc.", "U_GenConsig": "Design Resources, Inc.",
        "U_COO": "BD", "U_TermSale": "FOB", "U_MOT": re.sub(r"(?i)^by\s+", "", h.get("ship_mode", NULL)).title(),
        "U_HandoverDate": h.get("icd_date") or NULL,
        "U_NumCartons": sum(l["ctn"] for l in po_lines),
        "U_GrossWeight": round(sum(l["gw"] for l in po_lines), 2),
        "U_CubeMeters": round(sum(l["cbm"] for l in po_lines), 4),
        "Weight": round(sum(l["gw"] for l in po_lines), 2), "WeightUnit": 3,
        "Volume": round(sum(l["cbm"] for l in po_lines), 4), "VolUnit": 4,
        "U_MakerTerms": NULL, "U_PackListRemarks": NULL,
    })
    return row

def line_row(h, l, docentry, linenum, hdr, ex):
    row = OrderedDict()
    for col in hdr:
        v = ex.get(col)
        if col in SYSTEM_LINE: v = NULL
        if col in ("LineStatus", "InvntSttus"): v = "O"
        row[col] = v
    key = f"LOOKUP:PO{l['po']}[{l['item']}|{l['color']}]"
    row.update({
        "DocEntry": docentry, "LineNum": linenum, "VisOrder": linenum,
        "BaseType": 22, "U_BaseType": 22, "BaseRef": l["po"], "BaseDocNum": l["po"],
        "BaseEntry": key + ".DocEntry", "BaseLine": key + ".LineNum",
        "U_BaseEntry": key + ".DocEntry", "U_BaseLine": key + ".LineNum",
        "BaseCard": f"LOOKUP:CardCode[{h['vendor_name']}]",
        "ItemCode": l["item"], "Dscription": l["desc"], "Quantity": l["qty"], "InvQty": l["qty"], "PackQty": l["qty"],
        "BaseQty": l["qty"], "ShipDate": h["invoice_date"], "ActDelDate": h["invoice_date"],
        "Price": key + ".Price", "U_FOBPrice": key + ".U_FOBPrice", "Currency": "$", "WhsCode": "LOOKUP:default_whs",
        "CountryOrg": "BD", "U_Color": l["color"], "U_LinePOPrefix": l["prefix"],
        "U_CartonNo": f"{l['ctn_s']}-{l['ctn_e']}", "U_PcsPerCarton": l["qty_ctn"], "U_NumfCart": l["ctn"],
        "U_CartonWeight": l["gw_ctn"], "U_CartonLength": l["L"], "U_CartonWidth": l["W"], "U_CartonHeight": l["H"],
        "FreeTxt": f"PI {l['pi']} seq {l['seq']}; cust style {l['cust_style'] or '-'}; cust ref {l['custref'] or '-'}"
                   + (f"; mixed packing {l['mixed']}" if l.get("mixed") else ""),
    })
    return row

def merge_same_po_line(po_lines):
    # packing rows with same PO+item+color hit one PO line -> one GRPO line; keep packing detail in FreeTxt note
    merged = OrderedDict()
    for l in po_lines:
        k = (l["item"], l["color"].upper())
        if k not in merged:
            merged[k] = dict(l, packs=[(l["ctn"], l["qty_ctn"])])
            continue
        m = merged[k]
        m["packs"].append((l["ctn"], l["qty_ctn"]))
        m["ctn_s"], m["ctn_e"] = min(m["ctn_s"], l["ctn_s"]), max(m["ctn_e"], l["ctn_e"])
        m["ctn"] += l["ctn"]; m["qty"] += l["qty"]; m["gw"] += l["gw"]; m["nw"] += l["nw"]; m["cbm"] += l["cbm"]
        m["qty_ctn"] = max(m["qty_ctn"], l["qty_ctn"])
        m["gw_ctn"] = round(m["gw"] / m["ctn"], 2)
    for m in merged.values():
        if len(m["packs"]) > 1:
            m["mixed"] = " + ".join(f"{c}x{q}" for c, q in m["packs"])
    return list(merged.values())

LEAN_HDR = ["DocEntry","CardCode","CardName","DocDate","NumAtCard","Comments","U_FactoryInv","U_OrderNo","U_POPrefix",
    "U_CustRefNo","U_FOB","U_OceanDest","U_COO","U_TermSale","U_MOT","U_HandoverDate","U_NumCartons","U_GrossWeight",
    "U_CubeMeters","U_BLConsig","U_GenConsig","Address","Address2"]
LEAN_LINE = ["DocEntry","LineNum","BaseType","BaseRef","BaseEntry","BaseLine","ItemCode","Dscription","Quantity","Price",
    "U_FOBPrice","WhsCode","U_Color","U_LinePOPrefix","U_CartonNo","U_PcsPerCarton","U_NumfCart","U_CartonWeight",
    "U_CartonLength","U_CartonWidth","U_CartonHeight","CountryOrg","FreeTxt"]

def color_key(c):
    k = re.sub(r"[\s/-]", "", str(c or "")).upper()
    return "" if k == "NULL" else k.replace("GREY", "GRAY")

def load_lookup(path):
    # PO lookup export (po-lookup-table.md schema), csv or xlsx; keyed by LookupKey and by (PONum, ItemCode)
    import csv
    if path.lower().endswith(".csv"):
        rows = list(csv.DictReader(open(path, newline="", encoding="utf-8-sig")))
    else:
        ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        hdr = [c.value for c in ws[1]]
        rows = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True) if any(v is not None for v in r)]
    by_key, by_po_item, by_po_style, by_po_cnum = {}, defaultdict(list), defaultdict(list), defaultdict(list)
    for r in rows:
        if r.get("ItemCode") in (None, "NULL"): continue
        po = int(float(r["PONum"])); item = str(r["ItemCode"]).strip()
        key = r.get("LookupKey") if r.get("LookupKey") not in (None, "NULL") else f"{po}|{item}|{color_key(r.get('U_Color'))}"
        r["_po"], r["_item"] = po, item
        by_key[str(key).strip()] = r
        by_po_item[(po, item)].append(r)
        by_po_style[(po, item.split("-")[0].upper())].append(r)   # stock SKUs are <style>-<color>-<size>
        parts = item.split("-")
        if len(parts) >= 3:
            by_po_cnum[(po, "-".join(parts[:2]).upper())].append(r)  # C-number SKUs: C24-02235-KHA-OS
    exported = next((r.get("ExportedAt") for r in rows if r.get("ExportedAt")), None)
    return by_key, by_po_item, exported, by_po_style, by_po_cnum

def num(v):
    try: return float(v)
    except (TypeError, ValueError): return None

def resolve(l, lk):
    # match order: exact key -> C-number any color -> customer style prefix + color -> style prefix single line
    by_key, by_po_item, _, by_po_style, by_po_cnum = lk
    ck = color_key(l["color"])
    r = by_key.get(f"{l['po']}|{l['item']}|{ck}")
    status, note = "OK", ""
    if not r:
        cands = by_po_item.get((l["po"], l["item"]), [])
        if len(cands) == 1 and not color_key(cands[0].get("U_Color")):
            r = cands[0]  # custom item, PO line carries no color: normal
        elif len(cands) >= 1:
            return None, "MISMATCH", f"PO {l['po']} {l['item']}: no open line in color '{l['color']}'; open: {sorted(str(c.get('U_Color')) for c in cands)}"
    if not r:
        sc = by_po_cnum.get((l["po"], l["item"].upper()), [])
        hit = [x for x in sc if ck and (color_key(x["_item"].split("-")[2]) == ck or ck in color_key(x.get("U_Color")) or color_key(x.get("U_Color")) in ck)]
        if len(hit) == 1:
            r = hit[0]  # C-number SKU with color/size suffix: normal
        elif len(sc) == 1 and not hit:
            return None, "MISMATCH", f"PO {l['po']} {l['item']}: only open SKU is {sc[0]['_item']} color '{sc[0].get('U_Color')}', packing color '{l['color']}'"
        elif len(sc) > 1:
            return None, "MISMATCH", f"ambiguous: {len(sc)} open SKUs for {l['item']} on PO {l['po']}: {sorted(x['_item'] for x in sc)}; packing color '{l['color']}'"
    if not r and l.get("cust_style"):
        sc = by_po_style.get((l["po"], l["cust_style"].split("-")[0].split(".")[0].upper()), [])
        hit = [x for x in sc if color_key(x.get("U_Color")) == ck]
        if len(hit) == 1:
            r = hit[0]  # stock SKU keyed by customer style + color: normal
        elif len(sc) == 1 and not color_key(sc[0].get("U_Color")):
            r, status, note = sc[0], "REVIEW", f"matched by customer style {l['cust_style']}: SAP item {sc[0]['_item']} (no color on PO line)"
        elif len(sc) >= 1:
            return None, "MISMATCH", f"PO {l['po']} style {l['cust_style']}: no open line in color '{l['color']}'; open: {sorted(x['_item'] + ' ' + str(x.get('U_Color')) for x in sc)}"
    if not r:
        return None, "MISMATCH", f"PO line not found or not open: PO {l['po']} {l['item']} {l['color']}"
    open_qty = num(r.get("OpenQty"))
    if open_qty is not None and l["qty"] > open_qty + 1e-6:
        status, note = "MISMATCH", f"over-receipt: packing {l['qty']} vs open {open_qty:g}" + (f"; {note}" if note else "")
    if num(r.get("Price")) in (None, 0):
        status = "MISMATCH" if status == "MISMATCH" else "REVIEW"; note = (note + "; " if note else "") + "PO price is zero/blank"
    fob, pr = num(r.get("U_FOBPrice")), num(r.get("Price"))
    if fob and pr and (fob > 20 * pr or fob > 50):
        status = "MISMATCH" if status == "MISMATCH" else "REVIEW"; note = (note + "; " if note else "") + f"U_FOBPrice {fob:g} looks wrong vs Price {pr:g} - check the PO"
    return r, status, note

def apply_lookup(lrow, r):
    be, bl = int(float(r["BaseEntry"])), int(float(r["BaseLine"]))
    if r["_item"] != lrow["ItemCode"]:
        lrow["FreeTxt"] = f"packing style {lrow['ItemCode']}; " + str(lrow.get("FreeTxt") or "")
        lrow["ItemCode"] = r["_item"]
    lrow.update({"BaseEntry": be, "BaseLine": bl, "U_BaseEntry": be, "U_BaseLine": bl,
                 "BaseCard": r["CardCode"], "Price": num(r.get("Price")), "U_FOBPrice": num(r.get("U_FOBPrice")) if r.get("U_FOBPrice") not in (None, "") else NULL,
                 "WhsCode": r.get("WhsCode") or NULL})
    if r.get("Dscription"): lrow["Dscription"] = r["Dscription"]
    lrow["LineTotal"] = round(lrow["Quantity"] * (lrow["Price"] or 0), 2)

def build(template_path, out_path, factory_paths, group="invoice", lookup_path=None):
    tpl = load_template(template_path)
    lk = load_lookup(lookup_path) if lookup_path else None
    hdr_cols, hdr_ex = tpl["OPDN_Header"]
    line_cols, line_ex = tpl["PDN1_Line"]
    wb = openpyxl.Workbook()
    ws_h = wb.active; ws_h.title = "OPDN_Header"; ws_h.append(hdr_cols)
    ws_l = wb.create_sheet("PDN1_Line"); ws_l.append(line_cols)
    ws_gh = wb.create_sheet("GRPO_Header"); ws_gh.append(LEAN_HDR)
    ws_gl = wb.create_sheet("GRPO_Lines"); ws_gl.append(LEAN_LINE)
    ws_v = wb.create_sheet("Validation"); ws_v.append(["File", "Check", "Invoice", "Packing", "Generated", "Status"])
    ws_k = wb.create_sheet("Lookups_Needed"); ws_k.append(["DocEntry", "PO", "ItemCode", "Color", "Qty", "Needs"])
    docentry = 0
    po_seen = {}
    for fp in factory_paths:
        h = parse_factory(fp)
        by_po = defaultdict(list)
        for l in h["lines"]:
            by_po[l["po"]].append(l)
        groups = [sum((merge_same_po_line(v) for v in by_po.values()), [])] if group == "invoice" \
            else [merge_same_po_line(v) for v in by_po.values()]
        gen_qty = gen_ctn = gen_gw = gen_cbm = 0
        for po_lines in groups:
            docentry += 1
            hrow = header_row(h, po_lines, docentry, hdr_cols, hdr_ex)
            lrows, cardcodes, doc_total = [], set(), 0.0
            for i, l in enumerate(po_lines):
                po = l["po"]
                lrow = line_row(h, l, docentry, i, line_cols, line_ex)
                if lk:
                    r, status, note = resolve(l, lk)
                    if r:
                        apply_lookup(lrow, r); cardcodes.add(r["CardCode"]); doc_total += lrow["LineTotal"] or 0
                        if r.get("U_POPrefix") and l["prefix"] and r["U_POPrefix"] != l["prefix"]:
                            ws_v.append([h["file"], f"PO {po} prefix: packing {l['prefix']} vs SAP {r['U_POPrefix']}", None, None, None, "REVIEW"])
                    if status != "OK":
                        ws_v.append([h["file"], f"Lookup PO {po} {l['item']} {l['color']}: {note}", None, l["qty"], None, status])
                    if not r:
                        ws_k.append([docentry, po, l["item"], l["color"], l["qty"], f"UNRESOLVED - {note}"])
                else:
                    ws_k.append([docentry, po, l["item"], l["color"], l["qty"], "BaseEntry, BaseLine, Price, U_FOBPrice, WhsCode"])
                lrows.append(lrow)
            if lk and cardcodes:
                if len(cardcodes) > 1:
                    ws_v.append([h["file"], f"POs on this invoice belong to different vendors: {sorted(cardcodes)}", None, None, None, "MISMATCH"])
                cc = sorted(cardcodes)[0]
                cn = next((resolve(l, lk)[0].get("CardName") for l in po_lines if resolve(l, lk)[0] and resolve(l, lk)[0].get("CardCode") == cc), None)
                hrow.update({"CardCode": cc, "JrnlMemo": f"Goods Receipt PO - {cc}", "DocTotal": round(doc_total, 2), "DocTotalSy": round(doc_total, 2)})
                if cn and cn != h["vendor_name"]:  # sub-factory shipping against the PO vendor: normal
                    hrow["CardName"] = cn
                    hrow["Comments"] = f"Shipper {h['vendor_name'].rstrip('.')}. " + hrow["Comments"]
                for lrow in lrows: lrow["BaseCard"] = cc
                single = len({l["po"] for l in po_lines}) == 1
                if single:
                    r0 = next((resolve(l, lk)[0] for l in po_lines if resolve(l, lk)[0]), None)
                    if r0 and r0.get("U_SONo") not in (None, ""): hrow["U_SONo"] = r0["U_SONo"]
            ws_h.append(list(hrow.values())); ws_gh.append([hrow[c] for c in LEAN_HDR])
            for lrow in lrows:
                ws_l.append(list(lrow.values())); ws_gl.append([lrow[c] for c in LEAN_LINE])
            for i, l in enumerate(po_lines):
                po = l["po"]
                gen_qty += l["qty"]; gen_ctn += l["ctn"]; gen_gw += l["gw"]; gen_cbm += l["cbm"]
                if l.get("mixed"):
                    ws_v.append([h["file"], f"PO {po} {l['item']} {l['color']}: mixed packing {l['mixed']} merged into one line", None, None, l["qty"], "REVIEW"])
        for po in by_po:
            if po in po_seen:
                ws_v.append([h["file"], f"PO {po} also on {po_seen[po]} (split shipment, partial receipt)", None, None, None, "REVIEW"])
            po_seen.setdefault(po, h["invoice_no"])
        if not lk:
            ws_k.append([None, None, None, None, None, f"CardCode / U_ActManuf for {h['vendor_name']}"])
        else:
            ws_v.append([h["file"], f"PO lookup table ExportedAt {lk[2]}", None, None, None, "OK"])
        checks = [("Pieces", h["inv_pcs"], h["pac_pcs"], gen_qty), ("Cartons", h["cartons"], h["pac_ctn"], gen_ctn),
                  ("Gross weight kg", h["inv_gw"], h["pac_gw"], round(gen_gw, 2)), ("CBM", None, h["pac_cbm"], round(gen_cbm, 4))]
        for name, a, b, c in checks:
            ok = all(abs((x if x is not None else c) - c) < 0.01 for x in (a, b)) and (name == "CBM" or b is not None)
            ws_v.append([h["file"], name, a, b, c, "OK" if ok else "MISMATCH"])
        ws_v.append([h["file"], f"GRPOs generated (one per {group}); distinct POs", None, len(by_po), len(groups), "OK"])
        for l in h["lines"]:
            if l["ctn_e"] - l["ctn_s"] + 1 != l["ctn"]:
                ws_v.append([h["file"], f"Carton range {l['ctn_s']}-{l['ctn_e']} != CTN {l['ctn']} ({l['po_raw']})", None, None, None, "REVIEW"])
            if l["qty_ctn"] * l["ctn"] != l["qty"]:
                ws_v.append([h["file"], f"QTY/CTN x CTN != pcs on {l['po_raw']} {l['item']} {l['color']}", None, l["qty"], l["qty_ctn"] * l["ctn"], "REVIEW"])
        spans = sorted((l["ctn_s"], l["ctn_e"]) for l in h["lines"])
        for (s1, e1), (s2, e2) in zip(spans, spans[1:]):
            if s2 != e1 + 1:
                ws_v.append([h["file"], f"Carton numbering {'gap' if s2 > e1 + 1 else 'overlap'} between {e1} and {s2}", None, None, None, "REVIEW"])
        for po, po_lines in by_po.items():
            prefixes = {l["prefix"] for l in po_lines}
            if len(prefixes) > 1:
                ws_v.append([h["file"], f"PO {po} has mixed prefixes {prefixes}", None, None, None, "REVIEW"])
    for ws in (ws_h, ws_l, ws_gh, ws_gl):
        for col in ws.iter_cols(min_row=2):
            for c in col:
                if isinstance(c.value, dt.datetime): c.number_format = "yyyy-mm-dd"
    wb.save(out_path)
    return out_path

if __name__ == "__main__":
    args = sys.argv[1:]
    def opt(name, default):
        if name in args:
            i = args.index(name); v = args[i + 1]; del args[i:i + 2]; return v
        return default
    group = opt("--group", "invoice"); lookup = opt("--lookup", None)
    template, out_dir, files = args[0], args[1], args[2:]
    for f in files:
        inv = re.sub(r"[^A-Za-z0-9]+", "_", parse_factory(f)["invoice_no"]).strip("_")
        print(build(template, f"{out_dir}/OPDN_upload_{inv}.xlsx", [f], group, lookup))
