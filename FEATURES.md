# Features & CLI Reference

## CLI Arguments

| Argument | Default | Description |
|---|---|---|
| `--date` | *(required)* | Target date in `YYYY-MM-DD` format |
| `--start-time` | `00:00` | Start of time window in WIB (HH:MM) |
| `--end-time` | `23:59` | End of time window in WIB (HH:MM) |
| `--sort-by` | `time` | Sort order: `time`, `distance`, or `price` |
| `--from` | `home` | Location name for distance calculation (must match a `LOC_*` key in `.env`) |
| `-o`, `--output` | `output/kuyy_tennis_<date>.xlsx` | Output file path |
| `--send-otp` | | Send OTP email without crawling (first-time login step 1) |
| `--otp` | | Provide OTP code to complete login (first-time login step 2) |
| `--no-headless` | | Show the browser window (useful for debugging) |

## Data source

Activities are fetched from the `kuyy.app/api/events` backend API. The script authenticates via Playwright (headless Chromium), then calls the API directly. All activities for the requested date are fetched in a single call.

## Authentication

Kuyy uses email-based OTP login (no password). On first run, the script:

1. Opens the login page
2. Submits your email (from `KUYY_EMAIL` in `.env`)
3. Waits for you to provide the OTP code

The browser session is saved to `.session` and reused until it expires.

## Time handling

All times in the API are in UTC. The script converts to WIB (UTC+7) for display and filtering. The `--start-time` and `--end-time` arguments are in WIB.

## Distance calculation

Distances are computed using the haversine formula (great-circle distance). Locations are defined in `.env` with the `LOC_` prefix:

```
LOC_HOME=<lat>, <lon>
LOC_WORK=<lat>, <lon>
```

When `--sort-by distance` is used, activities are sorted by proximity to the chosen location. The distance column is included in the Excel output whenever any location is configured.

## Excel output

Each row is one activity. Columns:

| Column | Description |
|---|---|
| Activity Name | Session title from kuyy |
| Host / Community | Organizer or community name |
| Venue | Court / venue name |
| Address | Full street address |
| Date | Activity date |
| Start Time (WIB) | Start time converted to WIB |
| End Time (WIB) | End time converted to WIB |
| Price (Rp) | Price in Indonesian Rupiah |
| Dist from *location* (km) | Distance from chosen location (only when locations are configured) |
| Type | Derived from the activity name (case-insensitive). Tags any combination of `Single` ("single"), `Double` ("double"), `Coaching` ("coach"). Multiple matches are joined with `, `. Empty when none match. |
| Link | Direct URL to the activity on kuyy.id |

The header row is styled, columns are auto-sized, and Excel auto-filter is enabled.

## Session persistence

The `.session` file stores browser cookies/state so you don't need to OTP every run. If the session expires, the script will tell you to re-authenticate via `--send-otp` + `--otp`.

## Dependencies

- **playwright** -- headless browser for authentication
- **openpyxl** -- Excel file generation
- **python-dotenv** -- `.env` file loading
