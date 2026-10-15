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
| **GitHub Pages + Firebase** (main setup) | Firebase (Google's free database) – permanent, shared by the team; people sign in with Google or email; admins manage prices and the team |
| As a Claude artifact | Claude's shared storage |
| Without Firebase settings, or from a saved file | This browser only (IndexedDB) – not shared |

**Setup:** follow [SETUP.md](SETUP.md).

## Prices
This public repository contains **no prices**. An admin opens **Rules & prices → Import prices (.xlsx)** once and loads an Excel sheet with the columns `Product` and `Price` (factory price incl. VAT). The prices are stored in Firebase for the whole team, never in GitHub.

## Build
```
python scripts/build.py            # docs/index.html – served by GitHub Pages
python scripts/build.py --claude   # build/claude.html – the Claude artifact version
```
SheetJS 0.18.5 is loaded from cdnjs and the Firebase SDK 10.12.2 from gstatic.

## Files
- `src/core.js` – file recognition, parsing, pricing and totals (no UI)
- `src/head.part`, `src/body.part` – page styles, layout and UI script
- `scripts/build.py` – assembles `docs/index.html`
- `docs/firebase-config.js` – your Firebase project settings (fill in during setup)
- `firestore.rules` – who may read and write what

## Rules in short
- One price basis everywhere: units × factory price incl. VAT.
- Free gifts (사은품) are valued at the cost of the same product.
- B2B, 벌크 and 오프라인 lines found inside 택배 files count in 기타.
- Coupang lines from other suppliers are left out unless ticked.
