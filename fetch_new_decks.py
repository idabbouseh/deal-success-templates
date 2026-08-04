# -*- coding: utf-8 -*-
"""Download every non-archive .pptx from the crawl that isn't already extracted,
then dump its text (zip+XML, no python-pptx needed)."""
import json, os, re, sys, zipfile, urllib.request, urllib.parse, io

SP = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(SP, "..", "decks2")
os.makedirs(OUT, exist_ok=True)

with open(os.path.join(SP, "crawl_2026-08-04.json"), encoding="utf-8") as f:
    TREE = json.load(f)["tree"]
with open(os.path.join(SP, "..", "spauth", "sap_cookies.json"), encoding="utf-8") as f:
    auth = json.load(f)
COOKIE = "; ".join(f"{c['name']}={c['value']}" for c in auth["cookies"])

# decks already extracted in v0.3 (match on distinctive filename fragments)
HAVE = ["S4H_PCE_Greenfield", "HCM NA SAP SuccessFactors", "Master Data Governance",
        "SAP Work Zone", "BTP Intelligent Forms", "Custom Integrations Leveraging BTP",
        "BTP as the Digital Experience", "SAP BTP UX Services Overview", "SAP Group Reporting",
        "ROM Receivables Management June", "ROM Datasphere + SAC June", "AIActivationOnePager June",
        "Joule Activation LAC V2 June", "NA version - ASC BDC Analytics",
        "ASC BDC Analytics D@S - Scope options", "L1_model company", "Huntsman_DRC_Impl_Read-out_20_Jun_2025 (2)",
        "VAS_Services_Catalog-2026-05-20-v2", "VAS_Services_Briefing", "VAS_Services_Catalog-2026.pptx"]

def slug(path, name):
    p = path.strip("/").lower()
    base = re.sub(r"[^a-z0-9]+", "_", p).strip("_")
    base = base.replace("supply_chain_manufacturing", "scm").replace("isbn_currentgen_ariba", "ariba")
    base = base.replace("sap_cloud_erp_public", "cerp").replace("analytics_bdc", "bdc")
    base = base.replace("sap_s4_hana_extensions", "s4ext").replace("sap_s4_hana_cloud_private", "pce")
    base = base.replace("s4_hana_private_cloud_brownfield", "bf").replace("s4_hana_public_cloud", "s4pc")
    base = re.sub(r"_+", "_", base)[:52].strip("_")
    return base

targets = []
for path, files in sorted(TREE.items()):
    if not isinstance(files, list): continue
    if "/Archives" in path or "/Archive" in path or path.endswith(("/Archives", "/Archive")): continue
    for f in files:
        if not f["name"].lower().endswith(".pptx"): continue
        if any(h.lower() in f["name"].lower() for h in HAVE): continue
        targets.append((slug(path, f["name"]), f, path))

# de-dup identical slugs
seen = {}
final = []
for k, f, path in targets:
    n = seen.get(k, 0); seen[k] = n + 1
    final.append((k if n == 0 else f"{k}_{n+1}", f, path))

print(f"{len(final)} new decks to fetch")
manifest = []
for key, f, path in final:
    dest = os.path.join(OUT, key + ".pptx")
    if not os.path.exists(dest):
        url = "https://sap.sharepoint.com" + urllib.parse.quote(f["url"])
        try:
            req = urllib.request.Request(url, headers={"Cookie": COOKIE})
            with urllib.request.urlopen(req, timeout=180) as r:
                data = r.read()
            with open(dest, "wb") as fh:
                fh.write(data)
        except Exception as e:
            print(f"FAIL {key}: {e}")
            continue
    # text dump
    tpath = os.path.join(OUT, key + ".txt")
    try:
        with zipfile.ZipFile(dest) as z:
            names = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                           key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
            chunks = []
            for i, n in enumerate(names, 1):
                xml = z.read(n).decode("utf-8", "ignore")
                txt = " ".join(re.findall(r"<a:t>(.*?)</a:t>", xml, re.S))
                txt = re.sub(r"&amp;", "&", txt)
                txt = re.sub(r"&lt;", "<", txt); txt = re.sub(r"&gt;", ">", txt)
                chunks.append(f"--- SLIDE {i} ---\n{txt}")
            with open(tpath, "w", encoding="utf-8") as fh:
                fh.write(f"DECK: {f['name']}\nFOLDER: {path}\nSLIDES: {len(names)}\n\n" + "\n\n".join(chunks))
        manifest.append({"key": key, "file": f["name"], "folder": path, "slides": len(names),
                         "kb": f["kb"], "mod": f["mod"], "txtKb": round(os.path.getsize(tpath)/1024)})
        print(f"  {key}: {len(names)} slides, text {round(os.path.getsize(tpath)/1024)}KB")
    except Exception as e:
        print(f"DUMPFAIL {key}: {e}")

with open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8") as fh:
    json.dump(manifest, fh, indent=1)
print(f"\nDONE {len(manifest)} decks dumped -> {OUT}")
