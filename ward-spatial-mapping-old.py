import os
import pandas as pd
import geopandas as gpd
import folium
import webbrowser

# -----------------------------
# 1. Read Excel input file
# -----------------------------
file_path = "Mar24_Mar25.xlsx"       # Input file
geojson_file = "Ward_UK.geojson"     # Ward boundaries file
output_file = "Documents/Ward_Info.xlsx"       # New Excel output

# Load Excel (must have Ward and Constituency columns)
df = pd.read_excel(file_path)

# Group by Ward + Constituency name
ward_counts = (
    df.groupby(["Ward", "Constituency"])["Attendee ID"].nunique().reset_index(name="count").sort_values(by="count", ascending=False)
)

# Save Ward + Constituency counts into new Excel file
with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as writer:
    ward_counts.to_excel(writer, sheet_name="Ward", index=False)
print("Ward + Constituency counts written to Ward_Info.xlsx")

# -----------------------------
# 2. Merge with Ward GeoJSON using Constituency + cleaning logic
# -----------------------------
if os.path.exists(geojson_file):
    gdf = gpd.read_file(geojson_file)

    merged_rows = []

    for _, row in ward_counts.iterrows():
        ward_name = str(row["Ward"]).strip()
        constituency_name = str(row["Constituency"]).strip()

        # 1️⃣ Exact Ward name match
        matches = gdf[gdf["WD24NM"].str.lower() == ward_name.lower()]

        if len(matches) == 1:
            match_row = matches.iloc[0]

        elif len(matches) > 1:
            # 2️⃣ Multiple matches → check LAD24NM contains constituency
            lad_matches = matches[
                matches["LAD24NM"].str.contains(constituency_name, case=False, na=False)
            ]
            if len(lad_matches) >= 1:
                match_row = lad_matches.iloc[0]
            else:
                match_row = matches.iloc[0]

        else:
            # 3️⃣ Fuzzy-clean ward name (St./Saint)
            cleaned = (
                ward_name.replace("St.", "")
                .replace("St ", "")
                .replace("Saint", "")
                .strip()
            )

            matches = gdf[gdf["WD24NM"].str.contains(cleaned, case=False, na=False)]

            if not matches.empty:
                lad_matches = matches[
                    matches["LAD24NM"].str.contains(constituency_name, case=False, na=False)
                ]
                if len(lad_matches) >= 1:
                    match_row = lad_matches.iloc[0]
                else:
                    match_row = matches.iloc[0]
            else:
                match_row = None

        if match_row is not None:
            merged_rows.append({
                "Ward": ward_name,
                "Constituency": constituency_name,
                "count": row["count"],
                "latitude": match_row.get("LAT"),
                "longitude": match_row.get("LONG")
            })
        else:
            print(f"⚠️ No match found for Ward: {ward_name} ({constituency_name})")

    merged_df = pd.DataFrame(merged_rows)

    # Remove rows with empty lat/long before map generation
    before_len = len(merged_df)
    merged_df = merged_df.dropna(subset=["latitude", "longitude"])
    after_len = len(merged_df)
    removed = before_len - after_len
    if removed > 0:
        print(f"Removed {removed} rows with empty latitude/longitude before map generation.")

    # Save merged Ward info with Constituency
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        merged_df.to_excel(writer, sheet_name="Ward", index=False)
    print("✅ Ward sheet updated with Constituency info and geo coordinates")

    # -----------------------------
    # 3. Generate Geospatial Map
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

            folium.CircleMarker(
                location=[r["latitude"], r["longitude"]],
                radius=radius,
                popup=f"{r['Ward']}<br>{r['Constituency']}<br>Count: {count}",
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.85
            ).add_to(m)

        legend_html = """
        <div style="position: fixed; 
                    bottom: 30px; left: 30px; width: 180px; height: 120px; 
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
        print(f"Geospatial map generated with Ward + Constituency matching: {map_file}")

        webbrowser.open('file://' + os.path.realpath(map_file))

    else:
        print("No valid coordinates available to generate map.")

else:
    print("Ward GeoJSON not found. Please download and place in project folder.")
