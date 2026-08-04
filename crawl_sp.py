import json, urllib.request, urllib.parse, sys, os

SP = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SP, "..", "spauth", "sap_cookies.json"), encoding="utf-8") as f:
    auth = json.load(f)
cookie_hdr = "; ".join(f"{c['name']}={c['value']}" for c in auth["cookies"])

BASE = "https://sap.sharepoint.com/teams/Americas-CBDO"
ROOT = "/teams/Americas-CBDO/Shared Documents/Deal Success Center/Sharepoint Resources/Templates"

def get(path):
    url = (BASE + "/_api/web/GetFolderByServerRelativeUrl('" + urllib.parse.quote(path.replace("'", "''")) + "')"
           "?$expand=Folders,Files&$select=Folders/Name,Folders/ServerRelativeUrl,"
           "Files/Name,Files/TimeLastModified,Files/Length,Files/ServerRelativeUrl")
    req = urllib.request.Request(url, headers={"Cookie": cookie_hdr, "Accept": "application/json;odata=nometadata"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))

tree = {}
queue = [ROOT]
fetched = 0
while queue and fetched < 500:
    path = queue.pop(0)
    fetched += 1
    rel = path[len(ROOT):] or "/"
    try:
        j = get(path)
    except Exception as e:
        tree[rel] = f"ERR {e}"
        continue
    tree[rel] = [
        {"name": f["Name"], "mod": f["TimeLastModified"][:10], "kb": round(int(f["Length"]) / 1024), "url": f["ServerRelativeUrl"]}
        for f in j.get("Files", [])
    ]
    for fo in j.get("Folders", []):
        if fo["Name"] != "Forms":
            queue.append(fo["ServerRelativeUrl"])

import datetime
out = os.path.join(SP, "crawl_" + datetime.date.today().isoformat() + ".json")
with open(out, "w", encoding="utf-8") as f:
    json.dump({"foldersCrawled": fetched, "queueLeft": len(queue), "tree": tree}, f, indent=1)
nonempty = {k: len(v) for k, v in tree.items() if v and isinstance(v, list)}
print(f"folders={fetched} queueLeft={len(queue)} nonempty={len(nonempty)} totalFiles={sum(nonempty.values())}")
print(out)
