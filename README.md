# Kuyy Crawler

A CLI tool that pulls tennis activities from [kuyy.id](https://kuyy.id) and exports them to Excel, so you can find the right court at the right time without fighting the website's filters.

## Why this exists

1. **See everything at once** -- Kuyy's web UI paginates and loads slowly. This tool dumps every available session for a date into a single spreadsheet you can scan in seconds.
2. **Sort by what matters** -- Sort by distance from wherever you'll be (home, office, anywhere), by price, or by time. The website only sorts by time.
3. **Filter precisely** -- Narrow down to an exact time window (e.g. 18:00-21:00 WIB) and get only the sessions that fit your schedule, across all venues.

## Setup

```bash
# Install dependencies
pip3 install -r requirements.txt
python3 -m playwright install chromium

# Configure credentials and locations
cp .env.example .env
# Edit .env with your kuyy.id email and preferred locations
```

### .env format

```
KUYY_EMAIL=you@example.com
LOC_HOME=-6.229942, 106.888373
LOC_WORK=-6.225897, 106.808559
LOC_TRC=-6.230666, 106.780111
LOC_KEMANG=-6.273797, 106.819399
```

Add as many `LOC_*` entries as you like. The name after `LOC_` becomes the `--from` value.

## First run (login)

Kuyy uses email OTP, so the first run requires two steps:

```bash
# Step 1: trigger the OTP email
python3 crawl.py --date 2026-05-26 --send-otp

# Step 2: enter the code and crawl
python3 crawl.py --date 2026-05-26 --otp 1234
```

The session is saved locally (`.session` file) and reused on future runs -- no OTP needed until it expires.

## Usage

```bash
# All tennis activities for a date
python3 crawl.py --date 2026-05-27

# Evening only, sorted by distance from home
python3 crawl.py --date 2026-05-27 --start-time 18:00 --end-time 23:59 --sort-by distance

# Sorted by distance from work
python3 crawl.py --date 2026-05-27 --start-time 18:00 --end-time 23:59 --sort-by distance --from work

# Cheapest first
python3 crawl.py --date 2026-05-27 --start-time 18:00 --end-time 23:59 --sort-by price

# Custom output path
python3 crawl.py --date 2026-05-27 -o my_file.xlsx
```

Output is saved to `output/kuyy_tennis_<date>.xlsx` by default.

## All options

See `FEATURES.md` for the full reference, or run:

```bash
python3 crawl.py --help
```
