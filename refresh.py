# -*- coding: utf-8 -*-
"""One-command refresh of the Templates Explorer.

    python refresh.py            # crawl, diff, rebuild, report
    python refresh.py --push     # ...and push to GitHub Pages if anything changed

Requires a current SSO cookie jar at ../spauth/sap_cookies.json
(refresh it with the sap-auth-mcp sap_authenticate tool against
https://sap.sharepoint.com/teams/Americas-CBDO/ when it expires).
"""
import json, os, re, subprocess, sys, datetime

SP = os.path.dirname(os.path.abspath(__file__))
PUSH = "--push" in sys.argv

def crawls():
    return sorted(f for f in os.listdir(SP) if re.match(r"crawl_\d{4}-\d{2}-\d{2}\.json$", f))

def load(name):
    with open(os.path.join(SP, name), encoding="utf-8") as f:
        return json.load(f)["tree"]

def flat(tree):
    out = {}
    for path, files in tree.items():
        if not isinstance(files, list): continue
        for f in files:
            out[path + "/" + f["name"]] = f["mod"]
    return out

def run(script):
    r = subprocess.run([sys.executable, os.path.join(SP, script)], capture_output=True, text=True)
    sys.stdout.write(r.stdout)
    if r.returncode != 0:
        sys.stderr.write(r.stderr)
        raise SystemExit(f"{script} failed")
    return r.stdout

prev_name = crawls()[-1] if crawls() else None
prev = load(prev_name) if prev_name else {}

print("=" * 70)
print("1/3  crawling SharePoint library")
run("crawl_sp.py")

new_name = crawls()[-1]
if new_name == prev_name and prev_name:
    # same-day re-run: compare against the previous day's file instead
    older = [c for c in crawls() if c != new_name]
    prev = load(older[-1]) if older else prev
    prev_name = older[-1] if older else prev_name
cur = load(new_name)

a, b = flat(prev), flat(cur)
added = sorted(set(b) - set(a))
removed = sorted(set(a) - set(b))
changed = sorted(k for k in set(a) & set(b) if a[k] != b[k])
empty_before = {p for p, f in prev.items() if isinstance(f, list) and not f}
now_filled = sorted(p for p in empty_before if isinstance(cur.get(p), list) and cur.get(p))
new_folders = sorted(set(k for k, v in cur.items() if isinstance(v, list)) - set(prev))

print("=" * 70)
print(f"2/3  diff vs {prev_name}")
print(f"     +{len(added)} files   ~{len(changed)} modified   -{len(removed)} removed")
for label, items in (("NEW FILES", added), ("MODIFIED", changed), ("REMOVED", removed)):
    if items:
        print(f"\n  {label}:")
        for k in items[:40]:
            print("    " + k)
        if len(items) > 40: print(f"    … and {len(items)-40} more")
if now_filled:
    print("\n  FOLDERS THAT WERE EMPTY AND NOW HAVE CONTENT (these are the ones that go stale on the page):")
    for p in now_filled: print("    " + p)
if new_folders:
    print("\n  NEW FOLDERS:")
    for p in new_folders: print("    " + p)

print("=" * 70)
print("3/3  regenerating catalog + index.html")
out = run("gen_catalog.py")
gap = "COVERAGE GAP" in out

new_decks = [k for k in added if k.lower().endswith(".pptx") and "/Archive" not in k]
chg_decks = [k for k in changed if k.lower().endswith(".pptx") and "/Archive" not in k]
if new_decks or chg_decks:
    print("\n  DECKS NEEDING PRICING RE-EXTRACTION (index shows them as assets; matrices will be stale until re-run):")
    for k in new_decks: print("    NEW " + k)
    for k in chg_decks: print("    CHG " + k)
    print("    -> run fetch_new_decks.py, then the extraction workflow, for these keys.")

dirty = subprocess.run(["git", "-C", SP, "status", "--porcelain", "index.html", "index.template.html"],
                       capture_output=True, text=True).stdout.strip()
print("=" * 70)
if not dirty:
    print("RESULT: no change to the published page.")
elif gap:
    print("RESULT: page rebuilt BUT there is a coverage gap above — map those folders in gen_catalog.py CONFIG before pushing.")
elif PUSH:
    msg = f"chore: library re-index {datetime.date.today().isoformat()} (+{len(added)} files, ~{len(changed)} modified)"
    subprocess.run(["git", "-C", SP, "add", "index.html", "index.template.html", new_name], check=True)
    subprocess.run(["git", "-C", SP, "commit", "-m", msg], check=True)
    subprocess.run(["git", "-C", SP, "push", "origin", "main"], check=True)
    print("RESULT: pushed — https://idabbouseh.github.io/deal-success-templates/ updates in ~1 min.")
else:
    print("RESULT: page rebuilt locally. Re-run with --push to publish.")
