# Weekly Home Search Report

Runs every Saturday, pulls active single-family listings in Severna Park,
Arnold and Ellicott City from RentCast, ranks them, and emails you a
datestamped Excel report. Reports accumulate in `/reports`.

---

## 1. Get a RentCast API key

1. Go to **https://app.rentcast.io/app/api** and create an account.
2. Choose the **Free** plan — 50 requests/month, no card required.
3. Copy the key it generates. Keep it somewhere you can paste from once.

**Optional second key.** A second free account on a different email doubles
you to 100 requests/month. The script fails over automatically. If you skip
it, just leave `RENTCAST_API_KEY_2` empty — nothing breaks.

**Do not paste either key into a chat with me.** They live only in GitHub's
encrypted secrets.

## 2. Delivery — email is optional

**You do not need a Gmail App Password.** Every run delivers three ways
regardless:

- **A GitHub issue** with the ranked summary table. GitHub emails you when an
  issue opens in your own repo, so this is your notification with no SMTP
  credentials involved. Delete or close the issues as you read them.
- **A run artifact** — the .xlsx, downloadable from the Actions run page for
  90 days.
- **A commit to `/reports`** so history accumulates.

If you *do* want the workbook to land in your inbox as an attachment, add a
Gmail App Password. Two-factor must be on for the Google account first, and
Workspace admins can disable app passwords entirely — if the option isn't
there, that's why, and the three delivery paths above cover you.

1. Go to **https://myaccount.google.com/apppasswords**.
2. Name it something like `home-search` and create it.
3. **Google shows the 16-character code exactly once.** Copy it immediately —
   if you close the dialog you have to delete it and make a new one.

## 3. Create the repo

Make it **private**. Reports contain addresses and your search parameters.

```bash
mkdir home-search && cd home-search
# copy the files from this bundle in
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/home-search.git
git push -u origin main
```

## 4. Add the four secrets

Repo → **Settings** → **Secrets and variables** → **Actions** → **New
repository secret**. Add each of these:

| Name | Value |
|---|---|
| `RENTCAST_API_KEY` | your RentCast key |
| `RENTCAST_API_KEY_2` | second key, or leave blank |
| `EMAIL_ADDRESS` | your Gmail address — **optional**, leave unset to skip email |
| `EMAIL_APP_PASSWORD` | 16-character app password — **optional** |

## 5. Test it

Repo → **Actions** → **Weekly Home Search Report** → **Run workflow**.

Watch the log. You want to see listing counts per city, an `AVM calls spent`
line, and `REQUESTS THIS RUN`. Then check the repo's **Issues** tab — the
summary should be waiting there, and GitHub should have emailed you about it.

If you set up email and the run succeeds but nothing arrives, the app password
is the usual culprit — regenerate it. The issue and artifact still land either
way.

---

## Switching to biweekly

Cron cannot express "every other week", so the script handles it: with
`RUN_BIWEEKLY: "true"` it wakes up weekly and exits immediately on odd ISO
weeks. Edit that value in `.github/workflows/weekly-report.yml`.

To change the day, edit the cron. `0 7 * * 6` is 07:00 UTC Saturday
(3am ET / 4am EDT). The last field is the day: `0`=Sunday through `6`=Saturday.

## Quota management

Free tier is 50 requests/month per key. The script tracks usage in
`reports/quota.json` and stops at 48, so a manual re-run can't tip you over.

Sized for **one key**, worst case of a five-Saturday month:

| | |
|---|---|
| 3 listing calls × 5 runs | 15 |
| 4 AVM value calls × 5 runs | 20 |
| 3 property-record calls, once monthly | 3 |
| **Total** | **38** |

That leaves ~12 for manual test runs. `MAX_AVM_CALLS_PER_RUN` is the dial if
you want to trade breadth for depth — but on a single free key, 4 is about
the ceiling.

AVM calls are spent only on listings that already read MEETS ALL or NEAR MISS,
top of the ranking first, so the scarce requests go to houses you'd actually
consider. Everything else still gets an estimated value from the $/sqft
benchmark, which costs nothing.

If you ever add a second free key on a different email, raise
`MAX_AVM_CALLS_PER_RUN` back to 6 — the failover logic is already there.

## Things to know about the output

**Est. Value** is square footage × the seeded median $/sqft for that city.
It assumes average condition and no renovation. It's a screening number, not
an appraisal — the real one comes out of the workbook after you rate the
photos.

**Drive (min est)** comes from the `CITY_DRIVE_CALIBRATION` table at the top
of the script, not from a distance formula. Before your first real run, look
up each town center in Google Maps from 7765 Freetown Rd on a typical weekday
morning and replace my seeded numbers:

```python
CITY_DRIVE_CALIBRATION = {
    "Severna Park":  (16, 39.0704, -76.5455),
    "Arnold":        (22, 39.0334, -76.5033),
    "Ellicott City": (28, 39.2673, -76.7983),
}
```

Per-listing times are that baseline adjusted for how far the specific house
sits from its town center. So a north-Arnold listing reads ~18 and a west-side
Ellicott City listing on Rt 40 reads ~33 and drops to FAR — which is the
behavior you want. A city with no entry falls back to the old distance proxy.

Thresholds are 30 minutes preferred, 40 hard fail. Ellicott City at ~28 clears
as PASS but still scores lower on location, because that curve is continuous
and doesn't care about the PASS/FAR line.

**Subdivision** shows `—` when RentCast returns no subdivision for a record.
Verify it from the listing itself before trusting the city label. Municipal
boundaries are more reliable than the neighborhood polygons Zillow and Redfin
draw, but neither is a substitute for reading the listing.

`SUBDIVISION_ALLOWLIST` and `SUBDIVISION_BLOCKLIST` at the top of the script
are there for when you learn which subdivisions to chase or avoid.

## Files

| File | What it does |
|---|---|
| `home_search_report.py` | the whole pipeline |
| `.github/workflows/weekly-report.yml` | schedule, secrets, commit step |
| `requirements.txt` | requests, openpyxl |
| `reports/` | created on first run |
