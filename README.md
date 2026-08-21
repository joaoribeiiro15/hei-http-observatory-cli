# Observatory Scanner

Scans institutional websites via the **MDN HTTP Observatory API v2** and produces
a terminal report plus a CSV with grades, scores, failing tests, and recommendations.

Built for the thesis *"An Assessment of Web-Related Security in Norwegian Higher
Education Institutions"* (Østfold University College, 2026).

---

## Data availability

**This repository contains code only.** The scan results produced for the thesis
are not published, because they contain per-institution security findings that
cannot be disclosed.

`source/` and `results/` ship empty (with a `.gitkeep` placeholder). Place your
own institution list in `source/` and run the scanner to regenerate output. The
Norwegian and Portuguese HEI lists used in the thesis are published separately as
[hei-norway-dataset](../hei-norway-dataset) and
[hei-portugal-dataset](../hei-portugal-dataset).

---

## Requirements

- Python 3.10 or newer
- `openpyxl` (`pip install openpyxl`) — for Excel output

---

## Project structure

```
hei-http-observatory-cli/
├── scanner.py          # Main script
├── source/             # Drop your institution CSVs here (ships empty)
└── results/            # Output files written here automatically (ships empty)
```

---

## Usage

```bash
python scanner.py
```

The script will:

1. Discover every `.csv` file in `source/`
2. Scan each institution URL against the Observatory API
3. Print to the terminal results from the grade, score, failing tests, and recommendations
4. Write a timestamped `.csv` and `.xlsx` to `results/`

---

## Input CSV format

The CSV must contain at least a `url` column (case-insensitive). All other columns
are carried through to the output. The `no-heis-2026.csv` list used in the
thesis (published as [hei-norway-dataset](../hei-norway-dataset)) contains:

| Column | Description |
|---|---|
| ID | Internal institution identifier |
| Name | Full institution name |
| Category | Public / Private |
| url | Hostname to scan (e.g. `www.uia.no`) |
| NUTS2 | Eurostat NUTS2 region code |
| NUTS2_Label | NUTS2 region name |

---

## Output CSV columns

| Column | Description |
|---|---|
| ID, Name, Category, NUTS2, NUTS2_Label, url | Copied from input |
| grade | Observatory letter grade (A+ to F) |
| score | Numeric score (0–100+) |
| tests_passed | Number of tests that passed |
| tests_failed | Number of tests that failed |
| tests_quantity | Total number of tests run |
| failing_tests | Semicolon-separated list of failing test names |
| error | Error code if the scan failed |
| recommendation | Actionable recommendation based on the grade |
| details_url | Direct link to the full Observatory report |
| scanned_at | ISO 8601 timestamp of the scan |

---

## API rate limit

The Observatory API enforces one live scan per host per minute. If the same host
was scanned recently, a cached result is returned immediately. The script waits
2 seconds between requests to be respectful to the public API.

---

## Notes

- The API endpoint is `https://observatory-api.mdn.mozilla.net/api/v2/scan` (POST).
- The script strips `http://` and `https://` from URLs automatically.
- If a host cannot be resolved or returns a network error, the row is still written
  to the output CSV with an `error` column populated.
