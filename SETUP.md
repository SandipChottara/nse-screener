# NSE Screener — Cloud + Phone App

Runs your screener automatically on GitHub's servers every weekday evening,
and serves the results to a phone web app you open on your S24 Ultra.

**Cost: free.** No credit card, no server to maintain, no cold starts.

## What's required for the cloud

Just a **free GitHub account** — nothing else. Specifically:

| Piece | What it does | Cost |
|---|---|---|
| GitHub Actions | Runs the screener on a schedule (Mon–Fri, 18:45 IST) | Free — 2,000 min/month on private repos, unlimited on public |
| GitHub Pages | Serves the phone app + data over HTTPS | Free |
| Storage | `data.json` is a few KB | Free |

I chose this over Railway/Render/PythonAnywhere deliberately: Render's free
tier sleeps after 15 min and has no free scheduled jobs, Railway removed its
free tier (now $5/mo minimum), and PythonAnywhere's free tier only allows
daily jobs *and* silently stops them every 3 months unless you log in and
click Renew. GitHub Actions has none of those problems for this use case.

## ⚠️ The one real risk, stated upfront

**NSE geo-blocks many datacentre IPs.** GitHub Actions runners are in
Microsoft Azure (mostly US). So:

- **yfinance price data → works reliably.** Swing (Tier A) and Investment
  (Tier B) should work fine.
- **NSE Bhavcopy (OI/PCR) → may be blocked.** If so, the F&O tab will be
  empty and the app will say so honestly, rather than pretending no stocks
  qualified.

You'll know after the first run. If F&O is blocked and you need it, the
options are: keep running the PC script for F&O, or move the backend to an
India-region host (costs money, but solves it).

## Setup (about 10 minutes, one time)

### 1. Create the repo
On github.com → **New repository** → name it `nse-screener` →
**Private** is fine → Create.

### 2. Upload these files
Keep the exact folder structure:
```
nse_momentum_screener.py
cloud_runner.py
requirements.txt
my-full-nse-universe.csv        ← YOUR screener.in export (add this yourself)
.github/workflows/screener.yml
docs/index.html
docs/manifest.json
docs/data.json
```
Easiest: on the repo page click **Add file → Upload files**, drag
everything in, Commit. (The `.github` folder may be hidden on Windows —
if drag-and-drop misses it, use **Add file → Create new file** and type
`.github/workflows/screener.yml` as the filename, then paste the contents.)

**Don't forget `my-full-nse-universe.csv`** — without it, Swing and
Investment can't run (F&O still will).

### 3. Turn on GitHub Pages
Repo **Settings → Pages** → Source: **Deploy from a branch** →
Branch: `main`, Folder: **`/docs`** → Save.
After a minute it shows your URL:
`https://<your-username>.github.io/nse-screener/`

### 4. Run it once manually
Repo **Actions** tab → **Run NSE Screener** → **Run workflow**.
Takes 30–90 minutes (it's downloading ~500 stocks' price history).
Watch the log — it prints exactly what worked and what didn't.

### 5. Add to your phone
Open that Pages URL in Chrome on your S24 Ultra →
**⋮ menu → Add to Home screen**. It installs like a real app: own icon,
full screen, no browser bars.

## Daily use

The workflow runs itself every weekday at 18:45 IST. Open the app on your
phone whenever you want — tap **Refresh** if you want to force-fetch the
latest.

Tabs: **F&O** / **Swing** / **Invest** — each shows the top 10 with entry
price, score, stop-loss, target, and why it was picked.

## Keeping fundamentals current

`my-full-nse-universe.csv` is a static file in the repo, so refresh it
about weekly: re-export from your screener.in screen, then on GitHub go to
the file → pencil icon → **Upload** the new version → Commit. Fundamentals
barely move day to day, so weekly is plenty.

## Changing the schedule

Edit `.github/workflows/screener.yml`, the `cron:` line. It's in **UTC**
(IST = UTC + 5:30). Current `15 13 * * 1-5` = 13:15 UTC = 18:45 IST,
Mon–Fri — set after NSE publishes Bhavcopy (~17:30 IST).

## Note on GitHub Actions scheduling

Scheduled runs on GitHub can be delayed during peak load, sometimes by
30+ minutes, and may be skipped entirely on very busy days. For a daily
evening screener that's fine, but it's not a guaranteed-to-the-minute
trigger — if a day looks missing, just trigger it manually from the
Actions tab.
