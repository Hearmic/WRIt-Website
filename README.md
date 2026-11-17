# Rent Reset

A one-page campaign site (WRIT 2202, York University) about Toronto's rent gap between new and sitting tenants.
Every number on the page is loaded at runtime from `data/*.json`, which `scripts/fetch_data.py` builds from
official sources. `data/SOURCES.md` lists each source, its exact table or endpoint, the filters, and when it was retrieved.

```
index.html                      the site (Chart.js from cdnjs)
data/*.json, data/*.csv         processed data the page reads
data/raw/                       untouched downloads that produced them
data/SOURCES.md                 generated: sources, filters, retrieval dates
data/manifest.json              generated: per-dataset hashes and retrieval dates
scripts/fetch_data.py           fetches and processes everything
scripts/check_expected.py       compares fetched values with the figures in the project brief
.github/workflows/refresh.yml   scheduled refresh on GitHub
```

## Run it locally

The page loads its data with `fetch()`, which doesn't work from a `file://` URL, so serve the folder:

```sh
python3 -m http.server 8000
# open http://localhost:8000
```

## Refresh the data locally

```sh
python3 -m venv .venv-data
.venv-data/bin/pip install -r requirements.txt
.venv-data/bin/python scripts/fetch_data.py --sets all     # or e.g. --sets F, or --sets A,B
.venv-data/bin/python scripts/check_expected.py            # optional: compare with the brief's figures
```

Datasets: **A** StatCan asking/paid rents · **B** CMHC turnover rents · **C** 2021 Census homeownership ·
**D** StatCan fertility · **E** LTB applications · **F** Toronto shelters.

If a dataset fails to download or doesn't look the way the script expects (a table renamed, a row missing),
the script prints an error, keeps that dataset's previous files, carries on with the others, and exits with status 1.
Files are only rewritten when the numbers change.

## Deploy to GitHub Pages

1. Create a GitHub repository and push this folder to its `main` branch.
   (`.gitignore` leaves out the virtual environments. The `app/` Django scaffold isn't used by the site;
   delete it or leave it out if you don't want it published.)
2. In the repository: **Settings → Pages → Build and deployment → Source: Deploy from a branch**,
   branch `main`, folder `/ (root)`. The site appears at `https://<user>.github.io/<repo>/`.
3. **Settings → Actions → General → Workflow permissions → Read and write permissions**, so the refresh job can commit.

## Automatic refresh

`.github/workflows/refresh.yml` runs every day at 10:00 UTC:

- every day: **F** (Toronto shelters)
- Mondays: **A, B, D, E** as well (these sources update quarterly or yearly)
- **C** (2021 Census) never changes, so it only runs when you ask for it

It commits `data/` only when something changed, and that commit redeploys the site.
If a fetch fails, the previous files stay in place, the other datasets still update, and the run is marked as failed
(GitHub emails the repository owner by default).

**Trigger a refresh by hand:** repository → **Actions → Refresh data → Run workflow**, and type `all`
(or a list like `A,B`). From the command line: `gh workflow run refresh.yml -f sets=all`.

Note: GitHub turns off scheduled workflows in repositories with no activity for 60 days. The daily data commits
count as activity, but if F stops changing, re-enable the workflow from the Actions tab.
