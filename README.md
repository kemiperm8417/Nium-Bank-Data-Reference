# Bank Reference Data Explorer

Pulls bank and branch reference data from Nium's public **Reference Data API**
and exports it as an Excel workbook with **one sheet per country**, normalised
around each country's **domestic routing code**.

Source: `https://refdata.prod.nium.com/ref-data-service` ([OpenAPI](https://refdata.prod.nium.com/ref-data-service/v3/api-docs))
— read-only GETs, no credentials required.

> **Network requirement.** The API host resolves to a private address
> (`100.64.x.x`) and is reachable **only from Nium's network / VPN**. Run this
> tool on a machine connected to the Nium VPN or on a Nium-internal host.
> Cloud hosts — GitHub Codespaces, GitHub Actions' hosted runners, Streamlit
> Cloud, etc. — cannot reach it and will show *"Could not reach the Reference
> Data API"*.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then pick countries (or hit the **SEPA (36)** / **Common corridors** presets)
and download the workbook.

### Run on GitHub — only with Nium network access

The two GitHub paths below are wired up but **will not work as-is**, because
GitHub's cloud machines cannot reach the API (see the network note above).
They become usable only if (a) Nium exposes the API publicly, or (b) the
Actions workflow is pointed at a **self-hosted runner** inside the Nium
network (`runs-on: self-hosted` in `.github/workflows/export.yml`). Until then,
run the app locally on the VPN, or on an internal server and share its URL.


**Interactive app — Codespaces.** On the repo page click **Code → Codespaces →
Create codespace on main**. The dev container installs the requirements and
starts the app; a browser tab opens at a `https://…app.github.dev` URL. The
codespace sleeps after 30 idle minutes — just reopen it. By default only you
can open that URL; to share it, right-click port 8512 in the *Ports* panel and
set visibility to *Public* (or *Org*).

**Excel export — Actions.** Go to **Actions → Export bank reference data →
Run workflow**, choose a preset or type country codes, and run it. The
workbook lands in two places:

- as a downloadable artifact on the run page, and
- at a stable URL that always points at the most recent export:
  `https://github.com/kemiperm8417/Nium-Bank-Data-Reference/releases/download/latest/bank_reference_data.xlsx`

The Actions runner has 7 GB of RAM, so India's ~176k rows are fine there.

The same logic is available headless:

```bash
python3 refdata.py --countries US,GB,AU --out banks.xlsx
python3 refdata.py --sepa --out sepa.xlsx
python3 refdata.py --probe IN            # show strategy + observed field names
```

## Routing codes

| Country | Routing code | Where it comes from |
|---|---|---|
| US | ACH / ABA routing number | `branchCodes` → `branch.branchId` (= `bank.abaCode`) |
| IN | IFSC | `branchCodes` per bank → `branch.branchId` |
| GB | Sort code | `branchCodes` → `branch.branchId` |
| AU | BSB | `branchCodes` → `branch.bsbCode` |
| CA | Transit number + institution number | `branch.branchId` + `bank.bankId` |
| JP, NZ, HK, ZA, TH, VN, PL | Bank code + branch code | `bank.bankId` + `branch.branchId` |
| ID | Clearing code (Sandi Kliring) | `bank.clearingCode` |
| SEPA / EEA | BIC / SWIFT | `swiftCodes` — SEPA routes on IBAN + BIC, not a domestic code |
| SG, AE, MY, MX, BR, PH, KR, CN, TR | BIC / SWIFT | bank-level only; these return no branch records |

### SWIFT alongside the domestic code

For every country that has a domestic routing code (everything above except the
SEPA / BIC-only rows), the exporter **also** pulls `/{country}/swiftCodes` and
merges the BIC into each row by bank name, so a row carries both — e.g.
`Sort Code 406231` *and* `BIC NOCUGB2L`. The `BIC Source` column says where it
came from:

| `BIC Source` | Meaning |
|---|---|
| `branch` / `bank` | The BIC was in the branch/bank payload itself |
| `swiftCodes (name)` | Matched on the exact normalised bank name |
| `swiftCodes (loose)` | Matched after stripping legal suffixes (PLC, LTD, NA…) and parentheticals |
| `swiftCodes (multiple)` | The bank has several BICs — all are listed, `/`-separated |
| *(blank)* | No BIC in the payload and no name match in swiftCodes |

Coverage is a property of Nium's data, not the tool: it is reported per country
in the Summary sheet's Notes column.

Any country not explicitly mapped falls back to a **presence-driven** rule that
picks whichever of `bsbCode` / `clearingCode` / `abaCode` / `branchCode` /
`branchId` / BIC the API actually returned, and labels it honestly rather than
guessing a scheme name.

## Notes and gotchas

These are all verified against the live API and are the reasons the code is
shaped the way it is:

- **`branchCodes` silently truncates to 100 rows** unless an explicit `limit` is
  sent. Every call passes `limit=1000000`. A country reporting exactly 100 rows
  is the tell that this regressed.
- **India cannot be fetched in one call** — `/IN/branchCodes?limit=1000000`
  returns HTTP 500 after ~59s. It is fetched per bank instead
  (`/IN/{bankId}/branchCodes`, 297 banks, 6 at a time): ~176,600 rows in ~2 min.
  Any other country that 5xxs on the whole-country call degrades to the same
  path automatically.
- **gzip matters** — it cuts GB from 8.6 MB to 0.9 MB and halves wall-clock.
  `urllib` does not decompress automatically, so `_get_json` does it manually.
- **The UK is a SEPA member but is *not* exported as BICs.** Strategy is decided
  per country by data, not by SEPA membership, so GB keeps its 20,567 sort codes.
- **SEPA membership has no API endpoint.** `SEPA_COUNTRIES` in `refdata.py` is a
  hand-maintained list of the 36 members (EU27 + IS LI NO CH GB MC SM AD VA).
- The workbook's first sheet is always **Summary** — row counts, endpoint used,
  elapsed time and any error — so a partial export is self-documenting. One
  country failing never blocks the rest.
- Excel sheet names are capped at 31 characters and cannot contain `[ ] : * ? / \`;
  names are prefixed with the ISO code so truncation stays unambiguous.

## Files

| File | Purpose |
|---|---|
| `refdata.py` | API client, per-country strategy, normalisation, Excel writer, CLI |
| `app.py` | Streamlit UI |
| `.claude/launch.json` | Dev-server config (port 8512) |
