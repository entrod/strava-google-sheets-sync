# Geospatial Export Tools

This folder contains utilities for exporting and visualising Strava activity routes.

## Purpose

The Google Sheets integration stores activity metadata and (optionally) polylines.

Polylines are exported manually from the `Polylines` sheet into CSV format.
These CSV files can then be converted into a format compatible with Kepler.gl
for route visualisation and heatmap analysis.

## Workflow

1. Enable location export in `Code.js` (INCLUDE_LOCATION_DATA = true).
2. Run activity import in Google Sheets.
3. Copy polyline data into the `Polylines` sheet.
4. Export sheet as CSV.
5. Run:

   python convert_to_kepler.py input.csv output.csv

6. Import the resulting file into Kepler.gl.

## Why Kepler

Kepler.gl allows:
- Visualisation of all routes over time
- Density heatmaps
- Route overlap analysis
- Geographic training pattern analysis

This creates a full geospatial layer on top of Strava data.

## Privacy Note

Polylines may expose home/work locations.

Before sharing exported data:
- Remove first/last 500 meters of routes
- Or filter out sensitive start/end clusters

