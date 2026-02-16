import csv
import polyline
import json
from datetime import datetime

# -------------------------------
# Helpers
# -------------------------------

def decode_polyline_to_coordinates(encoded_polyline):
    """
    Decode a polyline string to [[lon, lat], ...] for GeoJSON.
    """
    if not encoded_polyline or encoded_polyline.strip() == "":
        return None

    try:
        coords = polyline.decode(encoded_polyline)  # [(lat, lon), ...]
        return [[lon, lat] for lat, lon in coords]
    except Exception as e:
        print(f"❌ Polyline decode error: {e}")
        return None


def parse_latlng(latlng_str):
    """
    Parse 'lat,lng' → [lon, lat]
    """
    if not latlng_str or latlng_str.strip() == "":
        return None
    try:
        lat, lng = latlng_str.split(",")
        return [float(lng.strip()), float(lat.strip())]
    except:
        return None


def detect_delimiter(file_path):
    """
    Auto-detect CSV delimiter.
    """
    with open(file_path, "r", encoding="utf-8") as f:
        sample = f.read(2048)

    if "\t" in sample:
        return "\t"
    elif ";" in sample:
        return ";"
    else:
        return ","


# -------------------------------
# Main converter
# -------------------------------

def convert_to_kepler_format(input_file, output_file):

    delimiter = detect_delimiter(input_file)
    print(f"✅ Detected delimiter: '{delimiter}'")

    runs_data = []

    with open(input_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=delimiter)

        print("✅ Detected columns:", reader.fieldnames)

        for row in reader:
            date = row.get("date", "").strip()
            distance = row.get("distance_km", "").strip()
            polyline_encoded = row.get("polyline", "").strip()
            start_latlng = row.get("start_latlng", "").strip()
            end_latlng = row.get("end_latlng", "").strip()

            # Skip empty rows
            if not date and not distance:
                continue

            # Decode route
            coordinates = decode_polyline_to_coordinates(polyline_encoded)

            # Start / end points
            start_point = parse_latlng(start_latlng)
            end_point = parse_latlng(end_latlng)

            # Proper GeoJSON LineString
            geojson_geometry = (
                json.dumps({
                    "type": "LineString",
                    "coordinates": coordinates
                }) if coordinates else ""
            )

            run_entry = {
                "date": date,
                "distance_km": distance,
                "start_lng": start_point[0] if start_point else "",
                "start_lat": start_point[1] if start_point else "",
                "end_lng": end_point[0] if end_point else "",
                "end_lat": end_point[1] if end_point else "",
                "has_route": "yes" if coordinates else "no",
                "num_points": len(coordinates) if coordinates else 0,
                "geometry": geojson_geometry
            }

            runs_data.append(run_entry)

    # -------------------------------
    # Write output
    # -------------------------------

    with open(output_file, "w", newline="", encoding="utf-8") as f:
        fieldnames = [
            "date",
            "distance_km",
            "start_lng",
            "start_lat",
            "end_lng",
            "end_lat",
            "has_route",
            "num_points",
            "geometry",
        ]

        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(runs_data)

    print(f"\n✅ Created {output_file}")
    print(f"📊 Processed {len(runs_data)} runs")
    print(f"🗺️  Runs with routes: {sum(1 for r in runs_data if r['has_route'] == 'yes')}")
    print(f"📍 Runs without routes: {sum(1 for r in runs_data if r['has_route'] == 'no')}")


# -------------------------------
# Execution
# -------------------------------

if __name__ == "__main__":

    input_file = "running_data.csv"
    output_file = "running_kepler.csv"

    convert_to_kepler_format(input_file, output_file)

    print("\n📋 Next steps in Kepler.gl:")
    print("1. Go to https://kepler.gl/demo")
    print("2. Click 'Add Data'")
    print("3. Upload running_kepler.csv")
    print("4. Add Layer → Trip")
    print("5. Under GeoJSON → select column: geometry")
    print("\n✅ Now your FULL routes will appear as lines!")

