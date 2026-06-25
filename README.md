# Script Hub — Flask App

An internal web app to run Python scripts through a browser. Upload an Excel file, click a button, download the result.

## Running locally

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start the app
python app.py

# 3. Open in your browser
http://localhost:5000
```

## Project structure

```
scripthub/
├── app.py                          ← Flask routes + in-memory job/result store
├── requirements.txt                ← flask, pandas, numpy, openpyxl, gunicorn, reportlab
├── Procfile                        ← gunicorn app:app
├── scripts/
│   ├── launch_check.py             ← Launch Check logic
│   ├── box_conversion.py           ← PC-per-Box conversion table (used by launch_check)
│   ├── garvis_export.py            ← Garvis Export logic
│   ├── promo_uplift_calc.py        ← Promo Uplift calculator logic
│   ├── uplift_applier.py           ← Uplift Applier logic
│   ├── cfr_orders.py               ← CFR Orders processor logic
│   └── cfr_pdf.py                  ← CFR Orders PDF generator
└── templates/
    ├── index.html                  ← Launch Check + Garvis Export (tabbed page)
    ├── promo_uplift_page.html      ← Promo Uplift calculator page
    ├── uplift_applier_page.html    ← Uplift Applier page
    ├── cfr_orders_page.html        ← CFR Orders page
    ├── promo_uplift.html           ← Promo Dashboard (hidden from nav, kept for reactivation)
    └── promo_detail.html           ← Promo Dashboard detail view (hidden from nav)
```

No database is used — all storage is in-memory (`_store = {}` in `app.py`). This means job status and temporary results are lost if the server restarts, but jobs finish in 1–3 minutes so this isn't an issue in practice.

## Adding a new script

1. Create a new file in `scripts/your_script.py` with a `run_your_script(file_obj, ...)` function that returns `(buf: io.BytesIO, stats: dict)`.
2. If the script is slow (large Excel files, more than ~20 seconds of processing), don't run it synchronously — Render's free tier has a 30-second request timeout. Instead, follow the background job pattern used by the other heavy modules:
   - The upload route saves the file to disk, starts a background thread, and immediately returns a `job_id`
   - The script itself accepts an optional `status_cb(msg)` callback to report progress
   - A `/your-module/status/<job_id>` route returns the current status
   - A `/your-module/download/<job_id>` route returns the finished file
   - The frontend polls the status route every 5–6 seconds and shows a progress message
3. Add the route(s) in `app.py`.
4. Add a page in `templates/` (or a tab in `index.html` for simple scripts), and a sidebar link in every template.

## Deploying (free)

### Render.com

1. Push to GitHub (a private repo is fine)
2. Go to render.com → New Web Service → connect your repo
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`
5. Deploy → you get a URL (e.g. `https://edgard-cooper.onrender.com`)
6. Every `git push` to the connected branch triggers an automatic redeploy (takes ~2–3 minutes)

Note: Render's free tier has 512MB RAM and a 30-second request timeout. Large Excel files (tens of thousands of rows) can hit both limits — read files with `dtype=str` where possible, free memory with `gc.collect()` after heavy steps, and use the background job pattern described above for anything that takes more than a few seconds.
