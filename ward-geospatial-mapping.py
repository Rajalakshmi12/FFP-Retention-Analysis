import os
import pandas as pd
import geopandas as gpd
import folium
import webbrowser

# -----------------------------
# 1. File configuration
# -----------------------------
file_path = "Mar24_Mar25.xlsx"         # Input Excel
geojson_file = "Ward_UK.geojson"       # ONS Wards GeoJSON (Dec 2024 UK BGC)
output_file = "Documents/Ward_Info.xlsx"  # Output Excel

# -----------------------------
# 2. Read Excel input (must have Ward & District)
# -----------------------------
df = pd.read_excel(file_path)

required_cols = {"Ward", "District"}
missing = required_cols - set(df.columns)
if missing:
    raise ValueError(f"Missing required columns in Excel: {missing}")

# Group by Ward + District and count unique attendees (if exists)
count_col = "Attendee ID" if "Attendee ID" in df.columns else None
if count_col:
    ward_counts = (
        df.groupby(["Ward", "District"])[count_col]
          .nunique()
          .reset_index(name="count")
          .sort_values(by="count", ascending=False)
    )
else:
    ward_counts = (
        df.groupby(["Ward", "District"])
          .size()
          .reset_index(name="count")
          .sort_values(by="count", ascending=False)
    )

# Save base data
with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as writer:
    ward_counts.to_excel(writer, sheet_name="Ward", index=False)
print("✅ Ward + District counts written to Ward_Info.xlsx")

# -----------------------------
# 3. Merge with GeoJSON
# -----------------------------
if not os.path.exists(geojson_file):
    raise FileNotFoundError("⚠️ GeoJSON file not found. Please download 'Ward_UK.geojson' and place it in the folder.")

gdf = gpd.read_file(geojson_file)
merged_rows = []

for _, row in ward_counts.iterrows():
    ward_name = str(row["Ward"]).strip()
    district_name = str(row["District"]).strip()
    match_row = None

    # Exact match on Ward + District
    matches = gdf[
        (gdf["WD24NM"].str.lower() == ward_name.lower()) &
        (gdf["LAD24NM"].str.lower() == district_name.lower())
    ]

    # Fuzzy handling for "St."/"Saint"
    if matches.empty:
        cleaned = (
            ward_name.replace("St.", "Saint")
                     .replace("St ", "Saint ")
                     .strip()
        )
        matches = gdf[
            (gdf["WD24NM"].str.contains(cleaned, case=False, na=False)) &
            (gdf["LAD24NM"].str.lower() == district_name.lower())
        ]

    if not matches.empty:
        match_row = matches.iloc[0]

    if match_row is not None:
        merged_rows.append({
            "Ward": ward_name,
            "District": district_name,
            "count": row["count"],
            "latitude": match_row.get("LAT"),
            "longitude": match_row.get("LONG"),
            "WD24NM": match_row.get("WD24NM"),
            "LAD24NM": match_row.get("LAD24NM"),
            "WD24CD": match_row.get("WD24CD")
        })
    else:
        print(f"⚠️ No match found for Ward: {ward_name} ({district_name})")

merged_df = pd.DataFrame(merged_rows)

# Drop rows with no coordinates before map
before_len = len(merged_df)
merged_df = merged_df.dropna(subset=["latitude", "longitude"])
removed = before_len - len(merged_df)
if removed > 0:
    print(f"⚠️ Removed {removed} rows without coordinates before mapping.")

# Save merged data
with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    merged_df.to_excel(writer, sheet_name="Ward_with_geo", index=False)
print("✅ Ward_with_geo sheet updated with coordinates.")

# -----------------------------
# 4. Generate Geospatial Map
# -----------------------------
if not merged_df.empty:
    merged_sorted = merged_df.sort_values(by="count", ascending=False)
    top10 = set(merged_sorted.head(10)["Ward"])
    max_count = merged_sorted["count"].max()

    center_lat = merged_df["latitude"].mean()
    center_lon = merged_df["longitude"].mean()

    m = folium.Map(location=[center_lat, center_lon], zoom_start=6)

    for _, r in merged_df.iterrows():
        count = r["count"]
        if count == 1:
            color = "green"
            radius = 6
        elif r["Ward"] in top10:
            color = "red"
            radius = 8 + (count / max_count) * 14
        else:
            color = "orange"
            radius = 6 + (count / max_count) * 10

        popup_html = (
            f"<b>Ward:</b> {r['Ward']}<br>"
            f"<b>District:</b> {r['District']}<br>"
            f"<b>Count:</b> {count}"
        )

        folium.CircleMarker(
            location=[r["latitude"], r["longitude"]],
            radius=radius,
            popup=popup_html,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.85
        ).add_to(m)

    # Add legend
    legend_html = """
    <div style="position: fixed; bottom: 30px; left: 30px; width: 190px; height: 120px; 
                background-color: white; z-index:9999; font-size:14px; 
                border:2px solid grey; padding:10px;">
    <b>Legend</b><br>
    🔴 Top 10 Wards<br>
    🟠 Other Wards<br>
    🟢 Count = 1
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    map_file = "ward_map.html"
    m.save(map_file)
    print(f"✅ Geospatial map generated: {map_file}")
    webbrowser.open('file://' + os.path.realpath(map_file))
else:
    print("⚠️ No valid coordinates found for mapping.")
