import openpyxl, requests, json, time

API_KEY = "AIzaSyA0Z8_GzoG68yaI8pAEkA_03Ig1N-6qki8"
GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"

# Load existing geocoded data to reuse coordinates
with open("hotels_geocoded.json", "r", encoding="utf-8") as f:
    existing = json.load(f)

# Build lookup by hotel name
existing_lookup = {}
for h in existing:
    existing_lookup[h["name"]] = h

# Read the new Excel file
wb = openpyxl.load_workbook("hotel_sic_summary_revised actual.xlsx")
ws = wb["SIC Hotels Summary"]

hotels = []
for row in ws.iter_rows(min_row=4, values_only=True):
    city, name, cat, area, sic, address = row[0], row[1], row[2], row[3], row[4], row[5]
    if not name:
        continue
    city_str = str(city or "").strip()
    name_str = str(name).strip()
    cat_str = str(cat or "").strip()
    area_str = str(area or "").strip()
    sic_str = str(sic or "").strip()
    addr_str = str(address or "").replace('\n', ' ').strip()

    hotels.append({
        "city": city_str,
        "name": name_str,
        "cat": cat_str,
        "area": area_str,
        "sic": sic_str,
        "address": addr_str,
    })

print(f"Total hotels from Sheet 2: {len(hotels)}")

results = []
geocoded_count = 0
reused_count = 0
failed_count = 0

for i, h in enumerate(hotels):
    # Try to reuse existing coordinates
    if h["name"] in existing_lookup:
        ex = existing_lookup[h["name"]]
        results.append({
            "city": h["city"],
            "name": h["name"],
            "cat": h["cat"],
            "area": h["area"],
            "sic": h["sic"],
            "address": h["address"],
            "lat": ex["lat"],
            "lng": ex["lng"]
        })
        reused_count += 1
        print(f"[{i+1}/{len(hotels)}] REUSED: {h['name']}")
        continue

    # Geocode new hotels
    search_addr = f"{h['address']}, {h['city']}, Vietnam"
    params = {"address": search_addr, "key": API_KEY}
    resp = requests.get(GEOCODE_URL, params=params)
    data = resp.json()
    lat, lng = None, None
    if data["status"] == "OK" and data["results"]:
        loc = data["results"][0]["geometry"]["location"]
        lat = loc["lat"]
        lng = loc["lng"]
        geocoded_count += 1
        print(f"[{i+1}/{len(hotels)}] GEOCODED: {h['name']} -> {lat}, {lng}")
    else:
        failed_count += 1
        print(f"[{i+1}/{len(hotels)}] FAIL: {h['name']} ({data['status']})")

    results.append({
        "city": h["city"],
        "name": h["name"],
        "cat": h["cat"],
        "area": h["area"],
        "sic": h["sic"],
        "address": h["address"],
        "lat": lat,
        "lng": lng
    })
    time.sleep(0.1)

with open("hotels_sheet2.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print(f"\nDone! Reused: {reused_count}, Geocoded: {geocoded_count}, Failed: {failed_count}")
print(f"Total: {len(results)} hotels saved to hotels_sheet2.json")
