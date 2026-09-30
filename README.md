# Rail Rate Tool

Password-locked rail price tracker (China origins to European destinations, SOC and COC).
The page in `docs/index.html` is encrypted with [StatiCrypt](https://github.com/robinmoisson/staticrypt):
without the password it is unreadable ciphertext. Source price sheets and extracted data are not in this repo.

## Files

| File | What it is |
|---|---|
| `docs/index.html` | The published page, encrypted. GitHub Pages serves this. |
| `template.html` | The page with a `__DATA__` placeholder (no data). |
| `extract.py` | Reads the price sheets under `Rail Freight` into `rail_prices.csv/json` (local only). |
| `etd.py`, `prep2.py` | Publish-step scripts: slim the data and add departure dates (ETD). |
| `IMPORT_CONVENTIONS.md` | How each source file type is read, and the agreed interpretation decisions. Read before importing new files. |

## Refreshing

1. Back up: copy the whole `rail-price-tool` folder to `..\rail-price-tool-backup-YYYYMMDD` (skip `.cache`).
2. Add the new export folder under `Rail Freight`, run `python extract.py`.
3. Ask Claude to rebuild the page from `rail_prices.json` and encrypt it into `docs/index.html`.
4. Push: `git add -A && git commit -m "Update" && git push`.

## Encrypting a page by hand

    npx staticrypt@3.5.4 page.html -d docs --config .staticrypt.json --short --template-title "Rail Rate Tool"

It asks for the password if `-p` is not given. Save the result as `docs/index.html`.
Reusing `.staticrypt.json` (local only) keeps the same salt, so browsers that ticked "Remember me" stay unlocked.
