# -*- coding: utf-8 -*-
"""Regenerate CATALOG + HEALTH in index.template.html from crawl_2026-08-04.json,
then rebuild index.html by re-injecting the existing SIZES/THUMBS/DETAILS/SLIDES tokens."""
import json, os, re, sys

SP = os.path.dirname(os.path.abspath(__file__))
CRAWLS = sorted(f for f in os.listdir(SP) if re.match(r"crawl_\d{4}-\d{2}-\d{2}\.json$", f))
if not CRAWLS:
    raise SystemExit("FATAL: no crawl_YYYY-MM-DD.json found — run crawl_sp.py first")
CRAWL = CRAWLS[-1]
CRAWL_DATE = CRAWL[6:16]
with open(os.path.join(SP, CRAWL), encoding="utf-8") as f:
    TREE = json.load(f)["tree"]
print(f"using {CRAWL}")

SIZE_ORDER = {"XXS": 0, "XS": 1, "S": 2, "M": 3, "L": 4, "XL": 5}
WORD_SIZE = {"XXSMALL": "XXS", "XSMALL": "XS", "SMALL": "S", "MEDIUM": "M", "LARGE": "L", "XLARGE": "XL"}
TYPE_RANK = {"deck": 0, "link": 1, "sheet": 2, "doc": 3, "mail": 4}

def ftype(name):
    n = name.lower()
    if n.endswith(".pptx"): return "deck"
    if n.endswith(".url"): return "link"
    if n.endswith((".xlsx", ".xlsm")): return "sheet"
    if n.endswith(".msg"): return "mail"
    return "doc"  # docx, pdf, zip, misc

def fsize(name):
    stem = re.sub(r"\.[A-Za-z0-9]+$", "", name)
    toks = [t for t in re.split(r"[^A-Za-z0-9]+", stem) if t]
    found = None
    for t in toks:
        up = t.upper()
        if up in SIZE_ORDER: found = up
        elif up in WORD_SIZE: found = WORD_SIZE[up]
    return found

def fentry(f):
    name = re.sub(r"\.url$", "", f["name"], flags=re.I)
    e = {"n": name, "t": ftype(f["name"])}
    s = fsize(f["name"])
    if s: e["s"] = s
    e["d"] = f["mod"][:7]
    return e

def files_at(key):
    v = TREE.get(key)
    if v is None or not isinstance(v, list): return None
    return v

def sortkey(e):
    return (TYPE_RANK.get(e["t"], 9), SIZE_ORDER.get(e.get("s"), 9), e["n"].lower())

def build_files(key, gather=False):
    out = []
    direct = files_at(key)
    if direct is None:
        raise SystemExit(f"FATAL: crawl has no folder {key!r}")
    out += [fentry(f) for f in direct]
    if gather:
        pfx = key + "/"
        for k, v in TREE.items():
            if k.startswith(pfx) and "/" not in k[len(pfx):] and isinstance(v, list):
                base = k[len(pfx):]
                if base in ("Archives", "Archive"): continue
                out += [fentry(f) for f in v]
    return sorted(out, key=sortkey)

def has_unclaimed_children(key, claimed):
    pfx = key + "/"
    for k, v in TREE.items():
        if not k.startswith(pfx) or not isinstance(v, list) or not v: continue
        if "/Archives" in k or "/Archive" in k or k.endswith("/Archives") or k.endswith("/Archive"): continue
        if any(k == c or k.startswith(c + "/") for c in claimed if c != key): continue
        return True
    return False

def js(o):
    return json.dumps(o, ensure_ascii=False)

def build_cat(cfg):
    folder = cfg.get("folder", cfg["name"])
    root_key = "/" + folder
    parts = [f'{{id:{js(cfg["id"])}, name:{js(cfg["name"])}, icon:{js(cfg["icon"])}']
    if cfg.get("folder"): parts[0] += f', folder:{js(cfg["folder"])}'
    if cfg.get("soon"): parts[0] += ", soon:true"
    body = f'\n  desc:{js(cfg["desc"])}'
    if cfg.get("rootFiles"):
        rf = build_files(root_key)
        if not rf: raise SystemExit(f"FATAL: expected rootFiles in {root_key}")
        body += f',\n  rootFiles:[{",".join(jsfile(e) for e in rf)}]'
    var_cfgs = cfg.get("vars", [])
    claimed = ["/" + folder + "/" + (v[1] if v[1] else "") if not v[1].startswith("../") else "/" + v[1][3:] for v in var_cfgs] if var_cfgs else []
    claimed = [c.rstrip("/") for c in claimed]
    vjs = []
    for v in var_cfgs:
        label, sub = v[0], v[1]
        opts = v[2] if len(v) > 2 else {}
        key = ("/" + sub[3:]) if sub.startswith("../") else "/" + folder + "/" + sub
        files = build_files(key, gather=opts.get("gather", False))
        entry = f'{{n:{js(label)}'
        if sub != label: entry += f', sub:{js(sub)}'
        entry += f', files:[{",".join(jsfile(e) for e in files)}]'
        if not files:
            entry += ", empty:true"
        elif not opts.get("gather") and has_unclaimed_children(key, claimed):
            entry += ", partial:true"
        entry += "}"
        vjs.append(entry)
    body += ",\n  vars:[" + ",\n        ".join(vjs) + "]" if vjs else ", vars:[]"
    return parts[0] + "," + body + "}"

def jsfile(e):
    s = f'{{n:{js(e["n"])},t:{js(e["t"])}'
    if e.get("s"): s += f',s:{js(e["s"])}'
    s += f',d:{js(e["d"])}}}'
    return s

EHS = ["Chemical Data Management", "Environment Management", "Incident Management",
       "Industrial Hygiene and Safety", "Management of Change", "Occupational Health", "Risk Assessment"]
IBP = ["IBP Demand", "IBP Inventory Optimization", "IBP Response & Supply",
       "IBP Sales & Operations Planning", "IBP Supply Chain Control Tower"]
SF_MODS = ["Performance and Goals", "Employee Central", "Employee Central Payroll", "Recruiting",
           "Onboarding", "Learning", "Compensation", "Succession and Development",
           "Workforce Analytics", "Platform Foundation", "Cloud Architecture"]

CONFIG = [
 dict(id="s4pce", name="SAP S4 HANA Cloud Private", icon="cloud",
  desc="RISE with SAP S/4HANA Cloud Private Edition t-shirts — Greenfield S/M/L fully priced ($3.8M–$14.1M, 31–54 wks) with scope-item, RICEFW and staffing detail; Brownfield now populated: NA Conversion & Upgrade sizing links, bootcamp deck + EMT master, and LAC Lift & Shift ROMs.",
  vars=[("S4 HANA Private Cloud - Greenfield", "S4 HANA Private Cloud - Greenfield"),
        ("Brownfield NA — overview & EMT master", "S4 HANA Private Cloud - Brownfield/North America"),
        ("Brownfield NA — S4 Conv & RISE Migration", "S4 HANA Private Cloud - Brownfield/North America/Brownfield S4 Conv & RISE Migration"),
        ("Brownfield NA — S4 Upgrade & RISE Migration", "S4 HANA Private Cloud - Brownfield/North America/Brownfield S4 Upgrade & RISE Migration"),
        ("Brownfield NA — Customer PPT Examples", "S4 HANA Private Cloud - Brownfield/North America/Customer PPT Examples"),
        ("Brownfield LAC — Lift & Shift ROMs", "S4 HANA Private Cloud - Brownfield/Latin America/Lift & Shift - LAC")]),
 dict(id="cerp", name="SAP Cloud ERP (Public)", icon="globe",
  desc="Public Cloud ERP (GROW) t-shirts by market unit. NA Greenfield now priced Startup→XL ($156K–$1.5M, 12–25 wks, incl. two Large LoB variants), plus the standalone Startup t-shirt ($227K, 12 weeks, finance core). Note the Startup price contradiction flagged in Catalog Health. LAC still awaits content.",
  vars=[("NA — S4 Public Cloud Greenfield", "NA/S4 HANA Public Cloud - Greenfield"),
        ("NA — S4 Public Cloud Startup", "NA/S4 HANA Public Cloud Startup"),
        ("LAC", "LAC")]),
 dict(id="s4ext", name="SAP S4 HANA Extensions", icon="puzzle",
  desc="Line-of-business extensions on S/4: Group Reporting (S/M/L $270K–$964K), Receivables Management, AFC (ES deck), PaPM and CLM (S/M/L links), and Multi-Bank Connectivity.",
  vars=[("Receivables Management (RM)", "Receivables Management (RM)"),
        ("Group Reporting", "Group Reporting"),
        ("AFC", "AFC"), ("PaPM", "PaPM"), ("CLM", "CLM"), ("MBC", "MBC")]),
 dict(id="bai", name="Business AI", icon="ai",
  desc="AI Activation and Joule Activation offers — activation bundles (S/M/L), Spark service, and LAC Joule decks (June 2026).",
  vars=[("AI Activation", "AI Activation"), ("Joule Activation", "Joule Activation")]),
 dict(id="bdc", name="Analytics & BDC", icon="chart",
  desc="Business Data Cloud, SAC, Datasphere and Databricks scoping assets — BDC Greenfield plus four Lift & Shift flavors, BW/4HANA conversions (In-Place / Remote / Shell), SAC BI / Planning / LA, and the Datasphere+SAC ROM.",
  vars=[("BDC — BW on PCE overview", "BDC"),
        ("BDC — Greenfield", "BDC/BDC Analaytics - BDC Greenfield"),
        ("BDC — Lift & Shift", "BDC/BDC Lift & Shift"),
        ("BDC — L&S + DB Migrate", "BDC/BDC Lift and Shift with DB Migrate"),
        ("BDC — L&S + Upgrade", "BDC/BDC Lift and Shift with Upgrade"),
        ("BDC — L&S + DB Migrate + Upgrade", "BDC/BDC Lift and Shift _DB Migrate_Upgrade"),
        ("BW/4HANA In-Place Conversion", "BDC/BW4HANA Conversion/BW 4HANA In-Place Conversion"),
        ("BW/4HANA Remote System Conversion", "BDC/BW4HANA Conversion/BW 4 HANA Remote System Conversion"),
        ("BW/4HANA Shell Conversion", "BDC/BW4HANA Conversion/BW4HANA Shell Conversion"),
        ("Datasphere", "Datasphere"),
        ("SAC — Business Intelligence", "SAC/SAC Business Intelligence"),
        ("SAC — Planning", "SAC/SAC Planning"),
        ("SAC — LA (ROM)", "SAC/SAC LA"),
        ("Databricks", "Databricks")]),
 dict(id="sf", name="Success Factors", icon="people", rootFiles=True,
  desc="Full HCM t-shirt pricing (NA, June 2026): EC $248–483K, ECP $1.4–2.7M, Recruiting, Onboarding, Learning, Compensation, P&G and more — every module's solutioning links now indexed (EC and ECP up to XL).",
  vars=[(m, "Effort Estimations - Solutioning Templates/SuccessFactors " + m) for m in SF_MODS]),
 dict(id="isbn", name="ISBN (Ariba & FG)", icon="wallet", folder="ISBN",
  desc="Intelligent Spend & Business Network — Ariba NA now fully priced from the June-2026 decks: Buying & Invoicing (S–XL $510K–$905K), Contracts ($200K/$247K/$529K), SLP, Sourcing (from $67K) and Supplier Risk, plus LAC EN/ES Contracts & Sourcing ladders (XXS–L). Fieldglass folder still in process — the FG matrices below remain seeded from the FG T-Shirt ROM Explorer draft.",
  vars=[("Ariba NA — Buying & Invoicing", "CurrentGen/Ariba/NA/Buying & Invoicing"),
        ("Ariba NA — B&I Solutioning Links", "CurrentGen/Ariba/NA/Buying & Invoicing/Solutioning Tool Links"),
        ("Ariba NA — B&I EMT workbooks", "CurrentGen/Ariba/NA/Buying & Invoicing/EMTs"),
        ("Ariba NA — Contracts", "CurrentGen/Ariba/NA/Contracts"),
        ("Ariba NA — SLP", "CurrentGen/Ariba/NA/SLP"),
        ("Ariba NA — Sourcing", "CurrentGen/Ariba/NA/Sourcing"),
        ("Ariba NA — Supplier Risk", "CurrentGen/Ariba/NA/Supplier Risk"),
        ("Ariba LAC — Contracts Stand Alone (EN/ES)", "CurrentGen/Ariba/LAC/SAP Ariba Contracts Stand Alone", {"gather": True}),
        ("Ariba LAC — Sourcing Stand Alone (EN/ES)", "CurrentGen/Ariba/LAC/SAP Ariba Sourcing Stand Alone", {"gather": True}),
        ("Ariba LAC — Sourcing Integrado (EN/ES)", "CurrentGen/Ariba/LAC/SAP Ariba Sourcing Integrado", {"gather": True}),
        ("Fieldglass (FG)", "CurrentGen/FG")]),
 dict(id="scm", name="Supply Chain", icon="boxes", folder="Supply Chain - Manufacturing",
  desc="Digital supply chain & manufacturing — now the largest priced category: 29 offerings extracted from 21 decks. EWM $950K–$2.8M, IBP across 5 areas, EHS ×7 (each S/M/L + EMT workbooks), EPPM, EAM, DMC (mandatory + 7 optional scopes), APM, Product Compliance Classic & Discrete, S/4 Retail, DRC Brazil, TM. ATTP, BN4L and Track & Trace await content.",
  vars=[("EAM", "EAM"), ("EPPM", "EPPM"), ("EWM", "EWM"), ("DMC", "DMC"), ("APM", "APM"),
        ("IAM — APM (duplicate of APM)", "IAM/IAM APM")] +
       [("EHS — " + a, "EHS/" + a, {"gather": True}) for a in EHS] +
       [(a, "IBP/" + a) for a in IBP] +
       [("Product Compliance — Classic", "Product Compliance/Classic"),
        ("Product Compliance — Discrete", "Product Compliance/Discrete"),
        ("S/4 for Retail", "Retail"),
        ("SAP DRC Brazil (NFe Cloud)", "SAP DRC Brazil"),
        ("Transportation Management", "TM"),
        ("ATTP", "ATTP"), ("BN4L", "BN4L"), ("Track and Trace", "Track and Trace")]),
 dict(id="btp", name="SAP BTP", icon="circuit", rootFiles=True,
  desc="BTP service t-shirts with full pricing: Work Zone Jump Start/Foundation/Expand ($70K/$249K/$486K), Intelligent Forms ($191–389K), Custom Integrations ($118–232K), Process Automation & Orchestration and DXP — each with S/M/L solutioning links.",
  vars=[("SAP Build (Work Zone)", "SAP Build"),
        ("BTP Intelligent Forms and Processes Framework", "BTP Intelligent Forms and Processes Framework"),
        ("BTP Process Automation and Orchestration", "BTP Process Automation and Orchestration"),
        ("BTP Digital Experience Platform", "BTP Digital Experience Platform"),
        ("Custom Integrations Leveraging BTP", "Custom Integrations Leveraging BTP")]),
 dict(id="vas", name="VAS", icon="rocket", rootFiles=True,
  desc="Value Assurance Services — 2026 services catalog & briefing decks, solutioning-tool links (AI, BDC, Joule, Signavio bundles) and real scope-document samples.",
  vars=[("Solutioning Tool Links", "Solutioning Tool Links"),
        ("Scope Document Samples", "Scope Document Samples")]),
 dict(id="btm", name="BTM", icon="tools",
  desc="Business Transformation Management — 14 accelerated-service links (Aug 2026), the CALM paquetes deck, Signavio baseline t-shirts, and four LeanIX service packages.",
  vars=[("Accelerated Services", "Accelerated Services"),
        ("CALM", "CALM"),
        ("Signavio", "Signavio"),
        ("LeanIX — service packages", "LeanIX", {"gather": True})]),
 dict(id="mdg", name="MDG", icon="db", rootFiles=True,
  desc="Master Data Governance t-shirts in four domains — Finance, Material, BP-Customer, BP-Supplier — each priced XS ($191K) through L ($1M), plus the MDG overview deck.",
  vars=[("Finance", "Finance"), ("Material", "Material"),
        ("Business Partner - Customer", "Business Partner - Customer"),
        ("Business Partner - Supplier", "Business Partner - Supplier")]),
 dict(id="drc", name="DRC", icon="doccheck",
  desc="Document & Reporting Compliance implementations — country tracks for Belgium + Singapore (incl. customer read-out) and Poland, each with deck + solutioning link.",
  vars=[("DRC Implementation for Belgium and Singapore", "DRC Implementation for Belgium and Singapore"),
        ("DRC Implementation for Poland", "DRC Implementation for Poland")]),
 dict(id="ux", name="UX", icon="ux",
  desc="User experience t-shirts — Fiori apps (XSmall–XLarge links + EMT workbooks across NI / LS+Upgrade / SC service models), Work Zone sizer & Mobile Start, and Service & Asset Management mobile apps.",
  vars=[("UX Fiori Apps", "UX Fiori Apps"),
        ("UX WorkZone Sizer and Mobile Start", "UX WorkZone Sizer and Mobile Start"),
        ("Service and Asset Management", "Service and Asset Management")]),
 dict(id="mc", name="Model Company", icon="factory", rootFiles=True,
  desc="Model Company for SAP Business Suite — L1 overview plus ten industry template packs (Automotive, Retail, CPG, Utilities, Oil & Gas, Fashion, Airline Back Office, EAM, R&D, EHS), each with its solutioning link.",
  vars=[(v, "Model Company Template " + v) for v in
        ["Automotive", "Retail", "CPG", "Utilities", "Oil and Gas", "Fashion", "Airline Back Office",
         "Enterprise Asset Management", "Research and Development", "EHS"]]),
 dict(id="ind", name="Industry Templates", icon="grid", folder="Industry", soon=True,
  desc="Industry-specific t-shirt packs. No content published yet — request via the Fast Track Team.", vars=[]),
 dict(id="sec", name="Security", icon="shield",
  desc="Security services t-shirts — Identity Access Governance (IAG) t-shirt & pricing deck with solutioning link.",
  vars=[("IAG", "IAG")]),
 dict(id="ds2", name="Design Success 2.0", icon="pen", folder="Design Success", soon=True,
  desc="Design Success 2.0 assets. No content published yet — request via the Fast Track Team.", vars=[]),
 dict(id="base", name="Baseline Templates", icon="layers",
  desc="The working baseline: NA Fast Track template, America's baseline template, FLC rates link, and the EMT (HLRE) estimation workbook with how-to guide.",
  vars=[("NA Fast Track Template", "NA Fast Track Template"),
        ("America's Baseline template", "America's Baseline template"),
        ("FLC Rates", "FLC Rates"),
        ("EMT", "EMT")]),
 dict(id="gov", name="Content & Governance", icon="doccheck", folder="Content", rootFiles=True,
  desc="Library reference content — CS&D practice-leads contact list, Enterprise AI Explorer, AI accelerators, Activate Methodology 2026, general assumptions, the DSC Knowledge Library (SME master list), the template governance model, and the templates-update tracker.",
  vars=[("AI Accelerators", "AI Acclerators"),
        ("Activate Methodology", "Activate Methodology"),
        ("General Assumptions", "General Assumptions"),
        ("SME Master List", "SME Master List"),
        ("Template Governance", "Template Governance"),
        ("Templates Update List", "../Templates Update List")]),
]

HEALTH = [
 dict(c="SAP Cloud ERP (Public)", st="data", note="PRICE CONTRADICTION — the Startup t-shirt is priced twice, differently, in the same folder tree. 'S4 HANA Public Cloud June 2026.pptx' (Greenfield, 31-Jul) sizes Startup at USD 156K / CAD 219K with Core Finance AND Sourcing & Procurement included; 'S4 HANA Public Cloud Startup_072026_v3.pptx' (3-Aug) sizes the same 12-week, identically phased Startup at USD 227K / CAD 318K with finance scope only. Smaller scope, higher price — a $71K (46%) gap an SAE could quote either way. Needs an owner decision on which is current before either deck is presented."),
 dict(c="Library re-index", st="ok", note="Re-indexed 04-Aug-2026 via SharePoint REST (204 folders, 421 files): ~180 files now indexed that the 30-Jul draft showed as empty or unindexed — Ariba NA+LAC, Supply Chain–Manufacturing, SAC, BDC migration flavors, Public Cloud NA (Greenfield + Startup), PCE Brownfield, all SF module links, BTP, UX Fiori, Model Company industries, LeanIX, IAG. Decks added since June are indexed as assets — pricing extraction still reflects the June-2026 deck set."),
 dict(c="ISBN (Ariba & FG)", st="gap", note="Ariba content arrived: five NA June-2026 decks (B&I, Contracts, SLP, Sourcing, Supplier Risk) + LAC EN/ES packages (XXS–L). Fieldglass subfolder still 'In Process' — the FG t-shirt matrices in this draft remain seeded from the FG T-Shirt ROM Explorer and need owner adoption."),
 dict(c="Supply Chain", st="gap", note="Populated Aug 2026 — 24 sub-areas indexed (EHS ×7 + EMTs, IBP ×5, EWM, EPPM, EAM, DMC, APM, Product Compliance ×2, Retail, DRC Brazil, TM). APM content sits in two places (/APM and /IAM/IAM APM) — consolidate. ATTP, BN4L and Track & Trace still empty. None of the ~15 decks are pricing-extracted yet."),
 dict(c="Industry Templates", st="empty", note="Empty folder behind a live tile."),
 dict(c="Design Success 2.0", st="empty", note="Empty folder behind a live tile."),
 dict(c="SAP Cloud ERP (Public)", st="gap", note="NA now populated: Greenfield June-2026 deck + XS–XL sizing links, and the Startup t-shirt (Jul 2026 v3, 12-week finance core). LAC subfolder still empty. Neither NA deck is pricing-extracted in this draft yet."),
 dict(c="SAP S4 HANA Cloud Private", st="data", note="Brownfield now populated (NA Conv/Upgrade sizing links + bootcamp deck + EMT master; LAC Lift & Shift ROMs). Greenfield deck-internal: comparison slides state 16,666/38,769/71,159 total hrs while staffing slides state 19,070/44,359/80,885 (incl. Governance & PMO) — reconcile which total is quoted."),
 dict(c="Analytics & BDC", st="gap", note="SAC now populated (BI / Planning / LA) and BDC carries Greenfield + four L&S flavors + BW/4HANA conversions; Databricks still empty. Hygiene: folder name 'BDC Analaytics' (typo), archive deck 'SAC Busnuss Intelligence.pptx' (typo), scope-options deck still carries '(1) (1) (3) (1)' suffixes. Datasphere+SAC deck partially Spanish for an NA/LAC-shared asset."),
 dict(c="VAS", st="gap", note="Two catalog versions side-by-side (2026 and 2026-05-20-v2) — retire the stale one or mark current."),
 dict(c="Success Factors", st="data", note="Deck data-quality: EC phase splits sum short of stated durations at every size (e.g. L sums 30 vs 34 wks); Work Zone splits exceed stated durations; Recruiting/Onboarding XL columns show scope+phases but no price; Cloud Architecture M has no duration and XL no price; Work Zone FX ratio (~1.11) differs from the rest of the deck (~1.40)."),
 dict(c="MDG", st="data", note="BP-Supplier Medium ($595K) sits ~$175-195K below the other three domains' Mediums ($770-792K) with the same scope structure — verify intentional. Medium CAD values rounded to '1.0M' break the deck's ~1.37 FX ratio. Overview deck filename has a double space."),
 dict(c="SAP BTP", st="data", note="Work Zone deck: slide-6 matrix and slide-15 detail show two different Medium phase plans; Small phase bars sum 4 wks against a stated 3. Overview deck still titled v0.1. A copy of the Custom Integrations deck also sits inside the SAP Build folder — misfiled."),
 dict(c="BTM", st="ok", note="Accelerated services actively updated (Aug 2026 — newest content in the catalog)."),
 dict(c="UX", st="data", note="The 'SAP BTP UX Services Overview v0.1' deck stored under SAP BTP is actually 'SAP S/4HANA UX NA T-shirts — Q4 2024' content (Dec 2024): stale vintage, nonstandard size names (X-Small/Standard/Strategic), and no complete price ladder. Rendered here under UX with a stale flag — needs a 2026 refresh."),
 dict(c="VAS", st="data", note="The same services catalog is embedded twice (briefing deck slides 9-17 ≡ catalog v2 slides 2-9) — one canonical source needed. Agentic AI Lab-to-Production states '10-20 weeks' in the range line but 'Large: 500d → 26 wks' in the tiers — reconcile."),
 dict(c="All categories", st="gap", note="No owner or description metadata exists in the library — this register proposes tracking both; descriptions and pricing extractions in this draft are AI-generated pending owner sign-off."),
]

cat_js = "const CATALOG = [\n " + ",\n ".join(build_cat(c) for c in CONFIG) + "\n];"

# extraction-derived findings, appended by the deck-issue summarizer
add_path = os.path.join(SP, "health_add.json")
if os.path.exists(add_path):
    with open(add_path, encoding="utf-8-sig") as f:
        extra = json.load(f)
    known = {(h["c"], h["note"][:60]) for h in HEALTH}
    n = 0
    for h in extra:
        if (h["c"], h["note"][:60]) not in known:
            HEALTH.append({"c": h["c"], "st": h.get("st", "data"), "note": h["note"]})
            n += 1
    print(f"health: +{n} rows from deck extraction findings")

health_js = "const HEALTH = [\n " + ",\n ".join(
    f'{{c:{js(h["c"])}, st:{js(h["st"])}, note:{js(h["note"])}}}' for h in HEALTH) + "\n];"

tpl_path = os.path.join(SP, "index.template.html")
with open(tpl_path, encoding="utf-8") as f:
    tpl = f.read()

# splice CATALOG
a = tpl.index("const CATALOG = [")
b = tpl.index("];", a) + 2
assert tpl.count("const CATALOG = [") == 1
tpl = tpl[:a] + cat_js + tpl[b:]
# splice HEALTH
a = tpl.index("const HEALTH = [")
b = tpl.index("\n];", a) + 3
assert tpl.count("const HEALTH = [") == 1
tpl = tpl[:a] + health_js + tpl[b:]

with open(tpl_path, "w", encoding="utf-8") as f:
    f.write(tpl)

# rebuild index.html: prefer dsc_data.json (the durable data file), else scrape the last build
idx_path = os.path.join(SP, "index.html")
data_path = os.path.join(SP, "dsc_data.json")
if os.path.exists(data_path):
    with open(data_path, encoding="utf-8") as f:
        d = json.load(f)
    tokens = {"__SIZES__": json.dumps(d["SIZES"], ensure_ascii=False),
              "__THUMBS__": json.dumps(d["THUMBS"], ensure_ascii=False),
              "__DETAILS__": json.dumps(d["DETAILS"], ensure_ascii=False),
              "__SLIDES__": json.dumps(d["SLIDECOUNTS"], ensure_ascii=False),
              "__DECKMAP__": json.dumps(d.get("DECKMAP", {}), ensure_ascii=False)}
    print(f"data: dsc_data.json ({sum(len(v) for v in d['SIZES'].values())} offerings, "
          f"{len(d['DETAILS'])} deep-dives, {sum(d['SLIDECOUNTS'].values())} slides)")
else:
    with open(idx_path, encoding="utf-8") as f:
        idx = f.read()

    def extract(name):
        m = "const " + name + " = "
        a = idx.index(m) + len(m)
        b = idx.index(";\nconst ", a)
        return idx[a:b]

    tokens = {"__SIZES__": extract("SIZES"), "__THUMBS__": extract("THUMBS"),
              "__DETAILS__": extract("DETAILS"), "__SLIDES__": extract("SLIDECOUNTS"),
              "__DECKMAP__": "{}"}
out = tpl
for k, v in tokens.items():
    assert out.count(k) == 1, f"token {k} count != 1"
    out = out.replace(k, v)

with open(idx_path, "w", encoding="utf-8") as f:
    f.write(out)

ncats = len(CONFIG)
nvars = sum(len(c.get("vars", [])) for c in CONFIG)
print(f"OK: {ncats} categories, {nvars} vars, CATALOG {len(cat_js)//1024}KB; index.html {len(out)//1024}KB")

# ---- coverage audit: any non-empty library folder the catalog does not surface? ----
covered = set()
for cfg in CONFIG:
    folder = cfg.get("folder", cfg["name"])
    covered.add("/" + folder)
    for v in cfg.get("vars", []):
        sub = v[1]
        key = ("/" + sub[3:]) if sub.startswith("../") else "/" + folder + "/" + sub
        covered.add(key)
        if len(v) > 2 and v[2].get("gather"):
            covered.update(k for k in TREE if k.startswith(key + "/"))

missing = []
for path, files in sorted(TREE.items()):
    if path == "/" or not isinstance(files, list) or not files: continue
    if "/Archives" in path or "/Archive" in path: continue
    if path in covered: continue
    # a folder is also covered if its files were pulled in by a parent's rootFiles
    parent_root = "/" + path.strip("/").split("/")[0]
    if path == parent_root and any(c.get("rootFiles") and "/" + c.get("folder", c["name"]) == path for c in CONFIG):
        continue
    missing.append((path, len(files)))

if missing:
    print("\n*** COVERAGE GAP — library folders with files that the catalog does NOT show ***")
    for p, n in missing:
        print(f"  {p}  ({n} files)")
    print("  -> add these to CONFIG vars, or they stay invisible on the page.")
else:
    print("coverage audit: every non-empty library folder is surfaced by the catalog")
