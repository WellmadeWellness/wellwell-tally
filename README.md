# Wellwell Monthly Tally

A single-page tool that turns Wellwell's order files into a monthly total for **Retail (리테일)**, **택배** and **기타**, valued at factory price (제조원가) including VAT.

## What it does
- Reads the Excel files people already have, recognising each one by its contents (not its file name):
  - Retail: B마트, 쿠팡 and 마켓컬리 order files (twice a week)
  - 택배: the daily 발주 files
  - 기타: 슈퍼키친, 삼성웰스토리 and N배송 files
  - Month-end 정산 file: used only in the Compare tab
- Builds up the current month day by day; earlier months stay available from the Month list.
- Flags missing prices, duplicate lines, 택배 lines with no channel and weekdays with no 택배 file.
- Compares the order-file tally with the month-end settlement, and shows a 12-month trend.
- Downloads the result as Excel.

Customer names, phone numbers and addresses in the 발주 files are read in the browser and never stored.

## Where data is kept
| Where the page runs | Where uploads are saved |
|---|---|
| As a Claude artifact | Shared online storage – the whole team sees the same data; editors manage prices |
| On GitHub Pages, from a file or any other website | This browser only (IndexedDB) – not shared |

## Prices
This public repository contains **no prices**. On the website, open **Rules & prices → Import prices (.xlsx)** and load an Excel sheet with the columns `Product` and `Price` (factory price incl. VAT). The prices are kept in that browser. The company price list is kept outside this repository.

## Build
```
python scripts/build.py --public
```
The output is `docs/index.html`, one self-contained file, which GitHub Pages serves. SheetJS 0.18.5 is loaded from cdnjs. To build a private copy with prices built in, put the price list at `data/prices.json` (it is git-ignored) and run `python scripts/build.py`.

## Files
- `src/core.js` – file recognition, parsing, pricing and totals (no UI)
- `src/head.part`, `src/body.part` – page styles, layout and UI script
- `scripts/build.py` – assembles `docs/index.html`

## Rules in short
- One price basis everywhere: units × factory price incl. VAT.
- Free gifts (사은품) are valued at the cost of the same product.
- B2B, 벌크 and 오프라인 lines found inside 택배 files count in 기타.
- Coupang lines from other suppliers are left out unless ticked.
