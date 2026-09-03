"""
Fail if app.py imports a local module that index.html doesn't mount.

stlite loads only the files listed in index.html's `files` map into the
browser's virtual filesystem; a missing entry surfaces as ModuleNotFoundError
on the live GitHub Pages site. Run by .github/workflows/pages.yml before deploy.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = open(os.path.join(ROOT, "app.py")).read()
html = open(os.path.join(ROOT, "index.html")).read()

mods = set(re.findall(r"^(?:from|import)\s+([A-Za-z_]\w*)", src, re.M))
local = sorted(m for m in mods if os.path.exists(os.path.join(ROOT, m + ".py")))
missing = [m for m in local if '"%s.py"' % m not in html]

for m in local:
    print("  %-16s %s" % (m + ".py", "MISSING from index.html" if m in missing else "mounted"))
if missing:
    print("\nAdd to the `files` map in index.html: " +
          ", ".join('"%s.py": { url: "./%s.py" }' % (m, m) for m in missing))
    sys.exit(1)
print("all local imports are mounted")
