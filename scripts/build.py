"""Build the single-file page into docs/index.html (served by GitHub Pages).
   python scripts/build.py            -> prices built in if data/prices.json exists (keep that file private)
   python scripts/build.py --public   -> NO prices (safe for a public website);
                                         an editor imports the price list from Excel in the Rules tab."""
import json, sys, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
pfile = root/'data/prices.json'
public = '--public' in sys.argv or not pfile.exists()
head = (root/'src/head.part').read_text(encoding='utf-8')
body = (root/'src/body.part').read_text(encoding='utf-8')
core = (root/'src/core.js').read_text(encoding='utf-8')
prices = {} if public else json.loads(pfile.read_text(encoding='utf-8'))
html = head + body.replace('/*CORE*/', core).replace('/*PRICES*/', json.dumps(prices, ensure_ascii=False))
page = '<!doctype html>\n<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"></head><body>\n' + html + '\n</body></html>\n'
out = root/'docs'; out.mkdir(exist_ok=True)
(out/'index.html').write_text(page, encoding='utf-8')
print(f"docs/index.html written ({'no prices – public' if public else str(len(prices)) + ' prices built in'}), {len(page)//1024} KB")
