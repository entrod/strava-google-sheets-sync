Strava → Google Sheets Sync (Apps Script)

It syncs my Strava activities into Google Sheets using Apps Script.

## What it does
- Imports latest Strava activities into a `StravaData` sheet
- Computes pace, speed, elevation stats and HR zone minutes (Z1–Z5) per activity
- Imports per-km splits into a `Splits` sheet and appends new rows into `SplitsDataStored`
- Syncs new activities into a long-term `Data` sheet (append-only)
- Writes an `Execution Log` sheet for traceability and debugging
- Includes a setup dialog in Sheets for OAuth token setup + token test
- Sends workout sessions to `Plan` for comparison with coach's sessions

## Sheets created / used
- `StravaData` (refreshed on each import)
- `Data` (append-only history)
- `Splits` (refreshed on each import)
- `SplitsDataStored` (append-only splits history)
- `Execution Log` (latest ~100 runs)
- `Plan` (manually created - compare workouts done with plan ahead
- `DataSorted` (manually created - Sorting data and adding g sheets formulas for extensive analysis of runs
- `Polylines` (manually created - to export polylines in a csv file)

## Setup (high level)
1. Create a Strava API app and get Client ID + Client Secret
2. Open the Google Sheet → Extensions → Apps Script
3. Paste the code from this repo
4. Reload the Sheet to see the `🏃 Strava` menu
5. Run: `🏃 Strava → Setup (första gången)` and follow the dialog
6. Run: `🏃 Strava → Importera aktiviteter`

## Privacy & safety
This project will fetch location data (start/end latlng and polyline).  
If you publish your sheet or share exports, remove or disable location fields.

## Roadmap
- Pagination (import more than the latest page)
- Backfill support (import full history)
- Rate limit handling + retries
- Derived metrics (acute/chronic load, load ratio)

## Derived Metrics

The following metrics are calculated inside Google Sheets on top of the imported raw data.

### Aerobic Efficiency Index (AEI)

Formula:
distance_km / avg_heartrate

Used to monitor aerobic development over time.
Higher value at similar effort = improved efficiency.

### Training Load

Load = moving_time_min × intensity_factor

Used for:
- Acute load (7 days)
- Chronic load (28 days)
- Load ratio (injury proxy)

## Geospatial Export

The sheet contains a manually maintained `Polylines` tab used to export
Strava activity polylines into CSV format.

A small Python utility is included:

tools/convert_to_kepler.py

This converts polyline CSV exports into a format compatible with Kepler.gl
for route visualization and heatmap analysis.


