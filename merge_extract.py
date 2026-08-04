# -*- coding: utf-8 -*-
"""Fold the new deck extractions into the page data (dsc_data.json),
and convert exported slide PNGs into the repo's slides/<key>/NNN.jpg reader assets."""
import json, os, re, sys

SP = os.path.dirname(os.path.abspath(__file__))
EX = os.path.join(SP, "..", "decks2", "extract")
RAW = os.path.join(SP, "..", "slides_raw2")
SLIDES = os.path.join(SP, "slides")
DATA = os.path.join(SP, "dsc_data.json")
IDX = os.path.join(SP, "index.html")

# deckKey -> (catalog id, source pptx filename, variant folder label)
MAP = {
 "ariba_na_buying_invoicing": ("isbn", "Ariba_BI_June_2026.pptx"),
 "ariba_na_contracts": ("isbn", "Ariba_Contracts_T-Shirts_June_2026.pptx"),
 "ariba_na_slp": ("isbn", "Ariba_SLP_T-Shirts_June_2026.pptx"),
 "ariba_na_sourcing": ("isbn", "Ariba_Sourcing_T-Shirts_June_2026.pptx"),
 "ariba_na_supplier_risk": ("isbn", "Ariba_Risk_June_2026.pptx"),
 "ariba_lac_contracts_group": ("isbn", "Ariba Contracts Stand Alone XXS–L (EN/ES)"),
 "ariba_lac_sourcing_group": ("isbn", "Ariba Sourcing Stand Alone / Integrado (EN/ES)"),
 "cerp_na_s4pc_greenfield": ("cerp", "S4 HANA Public Cloud June 2026.pptx"),
 "cerp_na_s4pc_startup": ("cerp", "S4 HANA Public Cloud Startup_072026_v3.pptx"),
 "bdc_sac_sac_business_intelligence": ("bdc", "SAC Business Intelligence June 2026.pptx"),
 "bdc_sac_sac_planning": ("bdc", "SAC Analytics Cloud Planning June 2026.pptx"),
 "bdc_sac_sac_la": ("bdc", "ROM SAC June 2026.pptx"),
 "bdc_bdc_bdc_lift_shift": ("bdc", "ASC BDC Analytics D@S - BDC BW4HANA_ConversionTShirts_V1.0.pptx"),
 "bdc_bdc_bdc_analaytics_bdc_greenfield": ("bdc", "NA - ASC BDC Analytics D@S - BDC Greenfield T-Shirt Sizing V1.0.pptx"),
 "scm_ewm": ("scm", "EWM_T-Shirt_for_NA_June_2026.pptx"),
 "scm_eam": ("scm", "EAM_T-Shirt_Presentation.pptx"),
 "scm_eppm": ("scm", "EPPM T-Shirt June 2026.pptx"),
 "scm_dmc": ("scm", "SAP Digital Manufacturing Cloud Presentation.pptx"),
 "scm_apm": ("scm", "IAM APM Tshirt and Pricing.pptx"),
 "scm_ibp_ibp_demand": ("scm", "IBP_Demand.pptx"),
 "scm_ibp_ibp_inventory_optimization": ("scm", "IBP Inventory Optimization.pptx"),
 "scm_ibp_ibp_response_supply": ("scm", "IBP Response & Supply.pptx"),
 "scm_ibp_ibp_sales_operations_planning": ("scm", "IBP S&OPT.pptx"),
 "scm_ibp_ibp_supply_chain_control_tower": ("scm", "IBP Supply Chain Control Tower.pptx"),
 "scm_ehs_chemical_data_management": ("scm", "SAP EHS Chemical Data Management.pptx"),
 "scm_ehs_environment_management": ("scm", "SAP EHS Environment Management.pptx"),
 "scm_ehs_incident_management": ("scm", "SAP EHS Incident Management.pptx"),
 "scm_ehs_industrial_hygiene_and_safety": ("scm", "SAP EHS Industrial Hygiene and Safety.pptx"),
 "scm_ehs_management_of_change": ("scm", "SAP EHS Management of Change.pptx"),
 "scm_ehs_occupational_health": ("scm", "SAP EHS Occupational Health.pptx"),
 "scm_ehs_risk_assessment": ("scm", "SAP EHS Risk Assessment.pptx"),
 "scm_product_compliance_classic": ("scm", "SAP Product Compliance Classic_TShirt size.pptx"),
 "scm_product_compliance_discrete": ("scm", "SAP Product Compliance for Discrete T-shirt Presentation 1.pptx"),
 "scm_retail": ("scm", "Tshirt - S4 for Retail.pptx"),
 "scm_tm": ("scm", "Transportation_Management_SAP_Corp_Branded.pptx"),
 "scm_sap_drc_brazil": ("scm", "ROM NFe Cloud - package with main dev enhancements June 2026.pptx"),
 "s4ext_afc": ("s4ext", "SAP_Advanced_Financial_Closing_AFC ES June 2026.pptx"),
 "s4ext_papm": ("s4ext", "SAP PaPM (2) June 2026.pptx"),
 "s4ext_mbc": ("s4ext", "SAP Multi-Bank Connectivity EN June 2026.pptx"),
 "security_iag": ("sec", "SAP IAG Tshirt and Pricing.pptx"),
 "sap_btp_btp_process_automation_and_orchestration": ("btp", "BTP Process Automation & Orchestration.pptx"),
 "btm_signavio": ("btm", "BTM Baseline Tshirt.pptx"),
 "btm_calm": ("btm", "Cloud ALM - paquetes.pptx"),
 "btm_leanix_group": ("btm", "LeanIX service package decks"),
 "pce_bf_north_america": ("s4pce", "CSD_West_SAE_Brownfield_Bootcamp_June_2026.pptx"),
 "pce_bf_lac_group": ("s4pce", "ROM L&S ECC / S4HANA / +Upgrade (LAC)"),
 "drc_drc_implementation_for_poland": ("drc", "Huntsman_DRC_Impl_Read-out (Poland).pptx"),
 "ux_service_and_asset_management": ("ux", "50114833_Mobile_Apps_Asset_Mgt_V6_1_Service_Overview.pptx"),
 "ux_ux_workzone_sizer_and_mobile_start": ("ux", "SAP_Build_Work_Zone_Sizer.pptx"),
 "business_ai_joule_activation": ("bai", "Joule Activation LAC June 2026.pptx"),
}

# ---------- load current page data ----------
if os.path.exists(DATA):
    with open(DATA, encoding="utf-8") as f:
        data = json.load(f)
else:
    with open(IDX, encoding="utf-8") as f:
        idx = f.read()
    def extract(name):
        m = "const " + name + " = "
        a = idx.index(m) + len(m)
        b = idx.index(";\nconst ", a)
        return json.loads(idx[a:b])
    data = {"SIZES": extract("SIZES"), "THUMBS": extract("THUMBS"),
            "DETAILS": extract("DETAILS"), "SLIDECOUNTS": extract("SLIDECOUNTS")}
    print(f"seeded dsc_data.json from index.html "
          f"({len(data['SIZES'])} cats, {len(data['DETAILS'])} details, {len(data['SLIDECOUNTS'])} decks)")

SIZES, DETAILS, SLIDES_N, THUMBS = data["SIZES"], data["DETAILS"], data["SLIDECOUNTS"], data["THUMBS"]

# ---------- fold in extractions ----------
SIZE_ORDER = {"XXS": 0, "XS": 1, "S": 2, "M": 3, "L": 4, "XL": 5, "one-size": 6}
added_off = added_det = 0
skipped = []
for fn in sorted(os.listdir(EX)) if os.path.isdir(EX) else []:
    if not fn.endswith(".json"): continue
    key = fn[:-5]
    with open(os.path.join(EX, fn), encoding="utf-8-sig") as f:
        r = json.load(f)
    if key not in MAP:
        skipped.append(key + " (no catalog mapping)")
        continue
    cat, deckname = MAP[key]
    offs = [o for o in r.get("offerings", []) if o.get("sizes")]
    if not offs:
        skipped.append(key + " (no priced offerings)")
    SIZES.setdefault(cat, [])
    for o in offs:
        sizes = sorted(o["sizes"], key=lambda s: SIZE_ORDER.get(s.get("size", ""), 9))
        clean = []
        for s in sizes:
            e = {"size": s.get("size") or "one-size"}
            for fld in ("tier", "usd", "cad", "weeks", "hours", "phases"):
                if s.get(fld) is not None: e[fld] = s[fld]
            e["scope"] = s.get("scope") or []
            clean.append(e)
        entry = {"offering": o["offering"], "variant": o.get("variant") or "", "deck": deckname,
                 "thumb": key, "sizes": clean}
        # replace any prior entry for the same offering+deck (idempotent re-runs)
        SIZES[cat] = [x for x in SIZES[cat] if not (x.get("thumb") == key and x.get("offering") == o["offering"])]
        SIZES[cat].append(entry)
        added_off += 1
    d = r.get("details") or {}
    d = {k: v for k, v in d.items() if v}
    if d:
        DETAILS[key] = d
        added_det += 1

# ---------- slide images ----------
try:
    from PIL import Image
except ImportError:
    Image = None
    print("WARNING: Pillow missing — skipping slide image conversion")

converted = 0
if Image and os.path.isdir(RAW):
    for key in sorted(os.listdir(RAW)):
        src = os.path.join(RAW, key)
        if not os.path.isdir(src) or not os.path.exists(os.path.join(src, ".complete")): continue
        pngs = sorted(p for p in os.listdir(src) if p.endswith(".png"))
        if not pngs: continue
        dst = os.path.join(SLIDES, key)
        os.makedirs(dst, exist_ok=True)
        for p in pngs:
            out = os.path.join(dst, p[:-4] + ".jpg")
            if os.path.exists(out): continue
            im = Image.open(os.path.join(src, p)).convert("RGB")
            if im.width != 720:  # match the v0.3 reader assets
                im = im.resize((720, round(im.height * 720 / im.width)), Image.LANCZOS)
            im.save(out, "JPEG", quality=60, optimize=True, progressive=True)
            converted += 1
        SLIDES_N[key] = len(pngs)

# ---------- deck filename -> reader key, so EVERY deck with slides is readable in-page ----------
def norm(name):
    return re.sub(r"[^a-z0-9]+", "", re.sub(r"\.pptx$", "", name, flags=re.I).lower())

DECKMAP = {}
for cat, offs in SIZES.items():          # priced decks (incl. the original v0.3 set)
    for o in offs:
        if o.get("deck") and o.get("thumb") and o["thumb"] in SLIDES_N:
            DECKMAP[norm(o["deck"])] = o["thumb"]

crawls = sorted(f for f in os.listdir(SP) if re.match(r"crawl_\d{4}-\d{2}-\d{2}\.json$", f))
if crawls:                                # unpriced decks, matched by the fetch script's slug
    with open(os.path.join(SP, crawls[-1]), encoding="utf-8") as f:
        tree = json.load(f)["tree"]
    def slug(path):
        b = re.sub(r"[^a-z0-9]+", "_", path.strip("/").lower()).strip("_")
        for a, z in (("supply_chain_manufacturing", "scm"), ("isbn_currentgen_ariba", "ariba"),
                     ("sap_cloud_erp_public", "cerp"), ("analytics_bdc", "bdc"),
                     ("sap_s4_hana_extensions", "s4ext"), ("sap_s4_hana_cloud_private", "pce"),
                     ("s4_hana_private_cloud_brownfield", "bf"), ("s4_hana_public_cloud", "s4pc")):
            b = b.replace(a, z)
        return re.sub(r"_+", "_", b)[:52].strip("_")
    for path, files in tree.items():
        if not isinstance(files, list) or "/Archive" in path: continue
        for fl in files:
            if not fl["name"].lower().endswith(".pptx"): continue
            k = slug(path)
            if k in SLIDES_N:
                DECKMAP.setdefault(norm(fl["name"]), k)

data["DECKMAP"] = DECKMAP
orphan = sorted(set(SLIDES_N) - set(DECKMAP.values()))

# ---------- write ----------
with open(DATA, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False)
print(f"reader map: {len(DECKMAP)} deck filenames -> {len(set(DECKMAP.values()))}/{len(SLIDES_N)} slide sets"
      + (f"; {len(orphan)} slide sets unreachable: {', '.join(orphan[:6])}" if orphan else ""))

kb = lambda p: round(os.path.getsize(p) / 1024)
print(f"merged: +{added_off} offerings, +{added_det} deck deep-dives, {converted} slide images converted")
print(f"totals: {sum(len(v) for v in SIZES.values())} offerings across {len(SIZES)} categories, "
      f"{len(DETAILS)} deep-dives, {sum(SLIDES_N.values())} readable slides")
if skipped:
    print("skipped:")
    for s in skipped: print("   " + s)
print(f"dsc_data.json = {kb(DATA)}KB")
