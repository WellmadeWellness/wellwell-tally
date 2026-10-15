"""Build the page.

  python scripts/build.py            -> docs/index.html  (GitHub Pages + Firebase; no prices inside)
  python scripts/build.py --claude   -> build/claude.html (the Claude artifact version)

Prices are never built into docs/: an admin imports them once from Excel and they are
stored in Firebase. A private copy can be built with prices if data/prices.json exists
(it is git-ignored):  python scripts/build.py --claude --with-prices
"""
import json, sys, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
claude = '--claude' in sys.argv
pfile = root / 'data/prices.json'
prices = json.loads(pfile.read_text(encoding='utf-8')) if ('--with-prices' in sys.argv and pfile.exists()) else {}
head = (root / 'src/head.part').read_text(encoding='utf-8')
body = (root / 'src/body.part').read_text(encoding='utf-8')
core = (root / 'src/core.js').read_text(encoding='utf-8')
html = head + body.replace('/*CORE*/', core).replace('/*PRICES*/', json.dumps(prices, ensure_ascii=False))
XLSX_TAG = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
if claude:
    out = root / 'build' / 'claude.html'
    page = html
else:
    fb = ('<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js"></script>\n'
          '<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-auth-compat.js"></script>\n'
          '<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore-compat.js"></script>\n'
          '<script src="firebase-config.js"></script>\n')
    assert XLSX_TAG in html
    html = html.replace(XLSX_TAG, XLSX_TAG + '\n' + fb)
    page = ('<!doctype html>\n<html lang="ko"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<meta name="robots" content="noindex"></head><body>\n' + html + '\n</body></html>\n')
    out = root / 'docs' / 'index.html'
out.parent.mkdir(exist_ok=True)
out.write_text(page, encoding='utf-8')
print(f'{out.relative_to(root)} written, {len(prices)} prices built in, {len(page)//1024} KB')
