# Kentucky American Seeds — Transaction Manager

A Streamlit web app that reads KAS sales-invoice PDFs (typed or handwritten),
extracts every line item with Anthropic Claude vision, lets you review and
edit the extraction, then writes the data into the existing
`2026_financing_transactions.xlsx` workbook on Google Drive.

Three tabs:

1. **📄 Upload & Process** — drag in PDFs, dual-pass extraction (extract +
   verify), editable side-by-side review with PDF preview, validation checks
   (math, required fields, low-confidence flags), duplicate detection, then
   one-click save.
2. **📊 Dashboard** — KPIs, filters (date / year / retailer / finance company /
   manufacturer / grower / product), trend charts, top-grower / top-product
   bars, finance mix, grower-summary table, exportable to CSV.
3. **🗺️ Grower Map** — Folium map of geocoded grower addresses. Click a
   marker to see total spend, invoice count, last purchase, and a full invoice
   history side panel.

## Architecture

- **Source of truth**: the `.xlsx` on Google Drive, accessed via a service
  account.
  - `Sheet1` (existing 21-column layout, ~2,290 historical rows) is preserved
    byte-for-byte. New rows are appended only.
  - `Finance Details` (new sheet, created lazily) holds loan #, batch, ACH date,
    amount-to-retailer, prepaid split, PDF source filename, Drive file ID, and
    a `Needs Review` flag for invoices missing CHS confirmation.
- **Processed PDFs** are uploaded to a Google Drive folder
  (`kas_processed_pdfs` next to the workbook) so both you and your collaborator
  can click through to the source document from the dashboard.
- **PDF rendering**: PyMuPDF (pure Python, no system dependencies — works on
  Streamlit Cloud without poppler).
- **Extraction**: Claude Sonnet 4.6 vision, dual-pass — first pass extracts
  with per-field confidence, second pass independently verifies and proposes
  corrections. The UI surfaces both versions and disagreements.
- **Auth**: shared password + signed cookie (30-day persistence).

## Local setup

```bash
pip install -r requirements.txt
```

Create `.streamlit/secrets.toml` (template in `.env.example`):

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
APP_PASSWORD = "kas2026!"
GOOGLE_DRIVE_FILE_ID = "1K_Ff4QCi_f_N6m5UqF1Xt2D2u4V6uwcC"
GCP_SERVICE_ACCOUNT_PATH = "./kas-transactions-cd7fa521e8a0.json"
COOKIE_SIGNING_KEY = "<32 random hex bytes>"
PROCESSED_PDFS_DRIVE_FOLDER_NAME = "kas_processed_pdfs"
TEST_MODE = true   # set to false to enable real writes
```

Generate a cookie key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

Run:

```bash
streamlit run app.py
```

## TEST_MODE

When `TEST_MODE = true`, the **Approve & Save** button performs all extraction,
validation, and dataclass conversion, but **does not write** to the workbook
or upload PDFs to Drive. Use this to walk through the UI with the test PDFs
without touching the source data. Flip to `false` to go live.

## Google Cloud service account setup

If you're starting from scratch (already done for this project):

1. <https://console.cloud.google.com/> → New Project (e.g. `kas-transactions`).
2. APIs & Services → Library → enable **Google Drive API**.
3. APIs & Services → Credentials → Create Credentials → Service Account.
   - Name: `kas-app`. No roles needed.
4. Open the new service account → Keys tab → Add Key → Create new key → JSON.
   Download the `.json` and place it in the project root.
5. In Google Drive, share `2026_financing_transactions.xlsx` with the service
   account's `client_email` (in the JSON), with **Editor** permission.

The app verifies access on start and will display a clear error if the share
is missing.

## Streamlit Community Cloud deploy

1. Push this directory to a private GitHub repo. The included `.gitignore`
   excludes the service-account JSON and `secrets.toml`.
2. <https://share.streamlit.io> → New app → connect to the repo, branch,
   `app.py`.
3. App **Settings → Secrets** — paste the same keys as in
   `.streamlit/secrets.toml`. For the service account, either:
   - Upload the JSON file alongside the repo (and set
     `GCP_SERVICE_ACCOUNT_PATH = "./kas-transactions-...json"`), **or**
   - Inline it as a single secret:
     ```toml
     GCP_SERVICE_ACCOUNT = '''
     { "type": "service_account", ... full JSON ... }
     '''
     ```
4. Set `TEST_MODE = false` once you've validated the flow.

The free Streamlit tier sleeps after inactivity but wakes on first hit; that's
fine for this use case.

## Key files

```
app.py                          # Streamlit entry, tab routing, write logic
lib/
  normalize.py                  # Unit/manufacturer/retailer cleanup tables
  workbook.py                   # Sheet1-preserving openpyxl reader/writer
  pdf_render.py                 # PyMuPDF page → base64 PNG for vision API
  extract.py                    # Dual-pass Claude extraction + adapter
  validate.py                   # Math + required-field validation rules
  auth.py                       # Password gate + signed-cookie persistence
  drive.py                      # Google Drive I/O + optimistic locking
  geocode.py                    # Nominatim with disk cache
  theme.py                      # KAS green/gold palette + CSS
.streamlit/
  config.toml                   # Theme + server settings
  secrets.toml                  # NOT committed
2026_financing_transactions.xlsx  # Local copy (Drive is source of truth)
logo.png                        # KAS logo, displayed in header
invoices/                       # Sample PDFs for testing
requirements.txt
packages.txt                    # Empty — PyMuPDF means no system deps needed
```

## Cost notes

- Each PDF extraction makes 2 Claude API calls (extract + verify), totaling
  ~10–20k input tokens and 2–6k output tokens depending on PDF length.
- A 2-page invoice runs ~$0.05; a 6-page multi-invoice handwritten ticket runs
  ~$0.15.
- The first pass alone (without verification) cuts cost in half — set
  `verify=False` in `lib/extract.py:extract_pdf` if needed.

## Known behaviors

- **Manufacturer ambiguity** on CHS rate strings like "Accolade CP BASF FMC"
  may resolve to either company. The editable review form lets you correct
  this before saving.
- **Handwritten tickets** without a CHS confirmation page are saved with
  `Finance Company` blank in Sheet1 and `needs_review = True` in the Finance
  Details sheet, so they show up in a "Pending CHS Match" filter later.
- **Concurrent edits**: the app reads the Drive file's `modifiedTime` before
  it generates a write, and refuses to upload if the file was changed by
  someone else in between (e.g. brother-in-law editing in Excel directly).
  You'll be prompted to reload and retry.
