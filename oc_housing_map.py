"""Standalone fullscreen Orange County housing need map (Plotly version).

Run with: python3 oc_housing_map.py
The generated map is saved beside this script and opened in the default browser.
"""

import hashlib
import json
import math
import os
import webbrowser
from collections import defaultdict

import openpyxl
import plotly.graph_objects as go


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# All data files are in the same directory as this script
XLSX_PATH = os.path.join(BASE_DIR, "HousingCentersOCFinalData.xlsx")
GEOJSON_PATH = os.path.join(BASE_DIR, "oc_tracts.geojson")
CES_PATH = os.path.join(BASE_DIR, "ces50_oc.json")
POVERTY_PATH = os.path.join(BASE_DIR, "poverty_oc.json")
TRACT_CENTROIDS_PATH = os.path.join(BASE_DIR, "tract_centroids.json")
OUTPUT_PATH = os.path.join(BASE_DIR, "oc_housing_map.html")


def clean_tract(value):
    text = str(value).strip()
    if "." in text:
        text = text.split(".")[0]
    return text.zfill(6)


def number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def need_color(score):
    if score >= 0.80:
        return "#D95A00"
    if score >= 0.70:
        return "#E8870A"
    if score >= 0.60:
        return "#F0A500"
    return "#2E86C1"


def enviro_color(value):
    if value is None:
        return "#AAAAAA"
    stops = [
        (0.00, (0, 80, 0)), (0.20, (64, 160, 0)),
        (0.40, (208, 208, 0)), (0.60, (240, 160, 0)),
        (0.80, (224, 64, 0)), (1.00, (160, 0, 0)),
    ]
    for (low, start), (high, end) in zip(stops, stops[1:]):
        if low <= value <= high:
            fraction = (value - low) / (high - low)
            rgb = [int(start[i] + (end[i] - start[i]) * fraction) for i in range(3)]
            return "#{:02X}{:02X}{:02X}".format(*rgb)
    return "#AAAAAA"


def build_bivariate_colors():
    """Create a 3x3 bivariate color matrix.

    Rows = CES 5.0 (Low to High)
    Cols = Housing Need Score (Low to High)

    Color scheme:
    - Low CES (green): Light green -> Medium green -> Dark green
    - Medium CES (yellow): Light yellow -> Yellow -> Gold
    - High CES (red): Light coral -> Coral -> Red
    """
    # 3x3 color matrix: [ces_bin][need_bin]
    return [
        # CES Low (green-ish)
        ["#e8f5e9", "#a5d6a7", "#66bb6a"],   # Need: Low, Med, High
        # CES Medium (yellow-ish)
        ["#fff9c4", "#ffeb3b", "#fdd835"],   # Need: Low, Med, High
        # CES High (red-ish)
        ["#ffccbc", "#ff8a65", "#ef5350"],   # Need: Low, Med, High
    ]


def build_bivariate_legend_html(score_t1, score_t2, ces_t1, ces_t2, colors):
    """Generate HTML for a 3x3 bivariate legend grid."""
    html = '''
    <div style="font-family: Arial, sans-serif; line-height: 1.2;">
        <div style="font-weight: bold; margin-bottom: 4px; font-size: 11px;">
            Housing Need Score →
        </div>
        <table style="border-collapse: collapse; font-size: 10px;">
            <tr>
                <td style="width: 40px; text-align: center; padding: 2px;">Low<br>({:.2f})</td>
                <td style="width: 40px; text-align: center; padding: 2px;">Med<br>({:.2f})</td>
                <td style="width: 40px; text-align: center; padding: 2px;">High<br>({:.2f})</td>
            </tr>
    '''.format(score_t1, score_t2, 1.0)

    # CES row labels
    ces_labels = [
        ("Low", ces_t1),
        ("Med", ces_t2),
        ("High", 1.0)
    ]

    for row_idx, (ces_label, ces_val) in enumerate(ces_labels):
        html += '            <tr>\n'
        html += '                <td style="width: 50px; text-align: right; padding: 2px 4px 2px 2px; vertical-align: middle;">{}<br>({:.2f})</td>\n'.format(ces_label, ces_val)
        for col_idx in range(3):
            color = colors[row_idx][col_idx]
            html += '                <td style="width: 36px; height: 36px; background-color: {}; border: 1px solid #999; border-radius: 3px;"></td>\n'.format(color)
        html += '            </tr>\n'

    html += '''        </table>
        <div style="margin-top: 4px; font-size: 9px; color: #666;">
            CES 5.0 Percentile ↓
        </div>
    </div>
    '''
    return html


def load_housing_scores():
    workbook = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    sheet = workbook["Short"]
    raw_headers = [cell.value for cell in sheet[1]]
    headers = []
    project_type_count = 0
    for header in raw_headers:
        if header == "Project Type":
            project_type_count += 1
            headers.append("Project Type" if project_type_count == 1 else "Housing Subtype")
        else:
            headers.append(header)

    with open(POVERTY_PATH, "r") as stream:
        poverty_data = json.load(stream)

    rows = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, values))
        if not row.get("Census Tract"):
            continue
        rent = number(row.get("Rent Burden %"))
        income = number(row.get("Median HH Income Percentile"))
        poverty = number(row.get("Poverty Line Percentile"))
        if rent is None or income is None or poverty is None:
            continue
        tract = clean_tract(row["Census Tract"])
        score = rent * 0.40 + (1 - income) * 0.25 + poverty * 0.35
        rows.append({
            "tract": tract,
            "city": str(row.get("City") or "Unknown"),
            "score": score,
            "units": number(row.get("Affordable Units")) or 0,
        })

    grouped = {}
    for row in rows:
        tract = row["tract"]
        if tract not in grouped:
            grouped[tract] = {
                "tract": tract,
                "city": row["city"],
                "score_total": 0,
                "units": 0,
                "programs": 0,
            }
        grouped[tract]["score_total"] += row["score"]
        grouped[tract]["units"] += row["units"]
        grouped[tract]["programs"] += 1

    scores = {}
    for tract, row in grouped.items():
        geoid = "06059" + tract
        poverty_info = poverty_data.get(geoid, {})
        below_poverty = poverty_info.get("below_poverty_count", 0)
        units_per_1k = (row["units"] / below_poverty * 1000) if below_poverty > 0 else 0
        scores[tract] = {
            "tract": tract,
            "city": row["city"],
            "score": round(row["score_total"] / row["programs"], 3),
            "units": int(row["units"]),
            "programs": row["programs"],
            "below_poverty": below_poverty,
            "units_per_1k_poverty": round(units_per_1k, 2),
        }

    median_score = sorted(item["score"] for item in scores.values())[len(scores) // 2]
    median_units_per_1k = sorted(item["units_per_1k_poverty"] for item in scores.values())[len(scores) // 2]
    for item in scores.values():
        item["underserved"] = (
            item["score"] > median_score and
            item["units_per_1k_poverty"] < median_units_per_1k
        )
    return scores, median_score, median_units_per_1k


def load_short_tab_properties():
    if not os.path.exists(TRACT_CENTROIDS_PATH):
        return []

    with open(TRACT_CENTROIDS_PATH, "r") as stream:
        tract_centroids = json.load(stream)

    workbook = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    sheet = workbook["Short"]
    raw_headers = [cell.value for cell in sheet[1]]
    headers = []
    project_type_count = 0
    for header in raw_headers:
        if header == "Project Type":
            project_type_count += 1
            headers.append("Project Type" if project_type_count == 1 else "Housing Subtype")
        else:
            headers.append(header)

    properties = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        row = dict(zip(headers, values))
        if not row.get("Property Name") or not row.get("Census Tract"):
            continue

        tract = clean_tract(row["Census Tract"])
        lat, lon = tract_centroids.get(tract, (None, None))
        if lat is None or lon is None:
            continue

        affordable = number(row.get("Affordable Units")) or 0
        total = number(row.get("Total Units")) or 0

        properties.append({
            "name": str(row.get("Property Name") or "Unknown").strip(),
            "city": str(row.get("City") or "Unknown").strip(),
            "tract": tract,
            "address": str(row.get("Address") or "").strip() if row.get("Address") else None,
            "project_type": str(row.get("Project Type") or "").strip() if row.get("Project Type") else None,
            "housing_subtype": str(row.get("Housing Subtype") or "").strip() if row.get("Housing Subtype") else None,
            "total_units": total,
            "affordable_units": affordable,
            "lat": lat,
            "lon": lon,
            "rent_burden": number(row.get("Rent Burden %")),
            "income_percentile": number(row.get("Median HH Income Percentile")),
            "poverty_percentile": number(row.get("Poverty Line Percentile")),
            "ces4_percentile": number(row.get("CalEnviroScreen 4.0 Percentile")),
        })

    return properties


def build_bivariate_colors():
    """Create a 3x3 bivariate color matrix.

    Rows = CES 5.0 (Low to High)
    Cols = Housing Need Score (Low to High)

    Color scheme: Blue (low need) -> Purple -> Red (high need)
    combined with Green (low CES) -> Yellow -> Red (high CES)
    """
    # 3x3 color matrix: [ces_bin][need_bin]
    # Using a blue-yellow-red diverging scheme for bivariate
    return [
        # CES Low (green-ish)
        ["#e8f5e9", "#a5d6a7", "#66bb6a"],   # Need: Low, Med, High
        # CES Medium (yellow-ish)
        ["#fff9c4", "#ffeb3b", "#fdd835"],   # Need: Low, Med, High
        # CES High (red-ish)
        ["#ffccbc", "#ff8a65", "#ef5350"],   # Need: Low, Med, High
    ]


CITY_BOUNDARIES_PATH = os.path.join(BASE_DIR, "oc_cities.geojson")


def load_city_boundaries():
    """Load official Orange County city boundaries from GeoJSON."""
    if not os.path.exists(CITY_BOUNDARIES_PATH):
        return None

    with open(CITY_BOUNDARIES_PATH, "r") as stream:
        city_geojson = json.load(stream)

    # Filter to only features with valid geometry and city name
    features = []
    for feature in city_geojson.get("features", []):
        props = feature.get("properties", {})
        city = props.get("CITY") or props.get("City_Capwords")
        geom = feature.get("geometry")
        if city and geom and geom["type"] in ("Polygon", "MultiPolygon"):
            features.append({
                "type": "Feature",
                "properties": {"city": city},
                "geometry": geom
            })

    return {
        "type": "FeatureCollection",
        "features": features
    }


def build_map():
    with open(GEOJSON_PATH, "r") as stream:
        tract_geojson = json.load(stream)

    with open(CES_PATH, "r") as stream:
        ces50 = json.load(stream)

    scores, median_score, median_units_per_1k = load_housing_scores()
    short_properties = load_short_tab_properties()

    # Load official city boundaries
    city_geojson = load_city_boundaries()
    if city_geojson is None:
        # Fallback to tract-dissolved boundaries
        city_geojson = {"type": "FeatureCollection", "features": []}

    # Build tract arrays for bivariate choropleth
    tract_ids = []
    tract_scores = []
    tract_cities = []
    tract_units = []
    tract_below_pov = []
    tract_units_per_1k = []
    tract_programs = []
    tract_underserved = []
    tract_ces50 = []

    for feature in tract_geojson.get("features", []):
        props = feature.get("properties", {})
        tract = clean_tract(props.get("TRACTCE", props.get("TRACT", "")))
        geoid = "06059" + tract

        score_data = scores.get(tract)
        if score_data is None:
            continue

        ces_val = ces50.get(geoid)

        tract_ids.append(geoid)
        tract_scores.append(score_data["score"])
        tract_cities.append(score_data["city"])
        tract_units.append(score_data["units"])
        tract_below_pov.append(score_data["below_poverty"])
        tract_units_per_1k.append(score_data["units_per_1k_poverty"])
        tract_programs.append(score_data["programs"])
        tract_underserved.append("Yes" if score_data["underserved"] else "No")
        tract_ces50.append(ces_val if ces_val is not None else None)

    # Calculate tertile breaks for both variables (only for tracts with both values)
    valid_pairs = [(s, c) for s, c in zip(tract_scores, tract_ces50) if c is not None]
    if valid_pairs:
        score_vals = [p[0] for p in valid_pairs]
        ces_vals = [p[1] for p in valid_pairs]

        score_sorted = sorted(score_vals)
        ces_sorted = sorted(ces_vals)
        n = len(score_sorted)

        score_t1 = score_sorted[n // 3]
        score_t2 = score_sorted[2 * n // 3]
        ces_t1 = ces_sorted[n // 3]
        ces_t2 = ces_sorted[2 * n // 3]
    else:
        score_t1 = score_t2 = 0.33
        ces_t1 = ces_t2 = 0.33

    # Assign each tract to a bivariate bin (0-8)
    bivariate_colors = build_bivariate_colors()
    tract_colors = []
    tract_bins = []  # 0-8 for legend reference

    for score, ces in zip(tract_scores, tract_ces50):
        if ces is None:
            tract_colors.append("#AAAAAA")
            tract_bins.append(-1)
            continue

        # Determine need bin (0, 1, 2)
        if score <= score_t1:
            need_bin = 0
        elif score <= score_t2:
            need_bin = 1
        else:
            need_bin = 2

        # Determine CES bin (0, 1, 2) - row in matrix
        if ces <= ces_t1:
            ces_bin = 0
        elif ces <= ces_t2:
            ces_bin = 1
        else:
            ces_bin = 2

        bin_idx = ces_bin * 3 + need_bin
        tract_colors.append(bivariate_colors[ces_bin][need_bin])
        tract_bins.append(bin_idx)

    # Create figure
    fig = go.Figure()

    # ---- 1. Bivariate Choropleth (Housing Need x CES 5.0) ----
    fig.add_trace(go.Choroplethmap(
        geojson=tract_geojson,
        locations=tract_ids,
        z=tract_bins,  # Use bin index for coloring
        featureidkey="properties.GEOID",
        colorscale=[[i/8, bivariate_colors[i//3][i%3]] for i in range(9)],
        zmin=0,
        zmax=8,
        marker_opacity=0.75,
        marker_line_width=0.5,
        marker_line_color="#555555",
        showscale=False,  # We'll use custom legend
        customdata=list(zip(tract_cities, tract_scores, tract_ces50, tract_units, tract_below_pov, tract_units_per_1k, tract_programs, tract_underserved)),
        hovertemplate=(
            "<b>Tract: %{location}</b><br>"
            "City: %{customdata[0]}<br>"
            "Housing Need Score: %{customdata[1]:.3f}<br>"
            "CES 5.0: %{customdata[2]:.1%}<br>"
            "Affordable Units: %{customdata[3]}<br>"
            "Below Poverty: %{customdata[4]:,}<br>"
            "Units/1k Poverty: %{customdata[5]:.1f}<br>"
            "Programs: %{customdata[6]}<br>"
            "Underserved: %{customdata[7]}"
            "<extra></extra>"
        ),
        name="Housing Need × CES 5.0",
        visible=True,
        showlegend=False,
    ))

    # Store tertile values for legend
    fig._score_t1 = score_t1
    fig._score_t2 = score_t2
    fig._ces_t1 = ces_t1
    fig._ces_t2 = ces_t2
    fig._bivariate_colors = bivariate_colors

    # ---- 2. City Boundaries (toggleable) ----
    # Extract line coordinates from city boundaries for Scattermap lines
    city_line_lats = []
    city_line_lons = []
    city_line_names = []

    for feature in city_geojson["features"]:
        city = feature["properties"]["city"]
        geom = feature["geometry"]
        if geom["type"] == "Polygon":
            # Polygon: coordinates = [exterior_ring, interior_ring1, ...]
            for ring in geom["coordinates"]:
                lons = [c[0] for c in ring]
                lats = [c[1] for c in ring]
                city_line_lons.extend(lons)
                city_line_lats.extend(lats)
                city_line_names.extend([city] * len(lons))
                # Add None to break the line between rings
                city_line_lons.append(None)
                city_line_lats.append(None)
                city_line_names.append(None)
        elif geom["type"] == "MultiPolygon":
            # MultiPolygon: coordinates = [[[exterior, interior...], [exterior, interior...], ...]]
            for polygon in geom["coordinates"]:
                for ring in polygon:
                    lons = [c[0] for c in ring]
                    lats = [c[1] for c in ring]
                    city_line_lons.extend(lons)
                    city_line_lats.extend(lats)
                    city_line_names.extend([city] * len(lons))
                    # Add None to break the line between rings
                    city_line_lons.append(None)
                    city_line_lats.append(None)
                    city_line_names.append(None)

    fig.add_trace(go.Scattermap(
        lat=city_line_lats,
        lon=city_line_lons,
        mode="lines",
        line=dict(width=1.5, color="#1A5276"),
        name="City Boundaries",
        visible=True,
        showlegend=False,
        hoverinfo="skip",
        hovertemplate="%{customdata}<extra></extra>",
        customdata=city_line_names,
    ))

    # ---- 3. Affordable Housing Locations (always visible) ----
    # Group properties by tract and aggregate totals (one bubble per tract)
    props_by_tract = defaultdict(list)
    for prop in short_properties:
        if prop["lat"] is not None and prop["lon"] is not None:
            props_by_tract[prop["tract"]].append(prop)

    # Build tract centroids from GeoJSON polygons (true geometric centroids)
    tract_centroids = {}
    for feature in tract_geojson.get("features", []):
        props = feature.get("properties", {})
        tract = clean_tract(props.get("TRACTCE", props.get("TRACT", "")))
        geom = feature.get("geometry")
        if geom and geom["type"] == "Polygon":
            coords = geom["coordinates"][0]
            lons = [c[0] for c in coords]
            lats = [c[1] for c in coords]
            centroid_lat = sum(lats) / len(lats)
            centroid_lon = sum(lons) / len(lons)
            tract_centroids[tract] = (centroid_lat, centroid_lon)

    tract_lats = []
    tract_lons = []
    tract_names = []  # Will hold "Tract XXXXXX (N properties)"
    tract_cities = []
    tract_tract_ids = []
    tract_total_units = []
    tract_affordable_units = []
    tract_property_counts = []
    tract_avg_rent_burden = []
    tract_avg_income_pct = []
    tract_avg_poverty_pct = []
    tract_avg_ces4 = []
    tract_radii = []
    tract_colors = []

    for tract, props in props_by_tract.items():
        # Use the tract's geometric centroid from GeoJSON
        lat, lon = tract_centroids.get(tract, (props[0]["lat"], props[0]["lon"]))

        # Aggregate totals
        total_props = len(props)
        total_units = sum(p["total_units"] for p in props)
        total_affordable = sum(p["affordable_units"] for p in props)
        avg_rent_burden = sum(p["rent_burden"] for p in props if p["rent_burden"] is not None) / max(1, sum(1 for p in props if p["rent_burden"] is not None))
        avg_income = sum(p["income_percentile"] for p in props if p["income_percentile"] is not None) / max(1, sum(1 for p in props if p["income_percentile"] is not None))
        avg_poverty = sum(p["poverty_percentile"] for p in props if p["poverty_percentile"] is not None) / max(1, sum(1 for p in props if p["poverty_percentile"] is not None))
        avg_ces4 = sum(p["ces4_percentile"] for p in props if p["ces4_percentile"] is not None) / max(1, sum(1 for p in props if p["ces4_percentile"] is not None))

        tract_lats.append(lat)
        tract_lons.append(lon)
        tract_names.append(f"Tract {tract} ({total_props} propert{'y' if total_props == 1 else 'ies'})")
        tract_cities.append(props[0]["city"])
        tract_tract_ids.append(tract)
        tract_total_units.append(total_units)
        tract_affordable_units.append(total_affordable)
        tract_property_counts.append(total_props)
        tract_avg_rent_burden.append(avg_rent_burden if any(p["rent_burden"] is not None for p in props) else None)
        tract_avg_income_pct.append(avg_income if any(p["income_percentile"] is not None for p in props) else None)
        tract_avg_poverty_pct.append(avg_poverty if any(p["poverty_percentile"] is not None for p in props) else None)
        tract_avg_ces4.append(avg_ces4 if any(p["ces4_percentile"] is not None for p in props) else None)
        # Radius: min 8, max 30, scaled by total affordable units in tract
        radius = max(8, min(30, 8 + total_affordable * 0.10))
        tract_radii.append(radius)
        tract_colors.append("#D95A00")

    def fmt_pct(val):
        if val is None:
            return "N/A"
        return f"{val:.1%}"

    fig.add_trace(go.Scattermap(
        lat=tract_lats,
        lon=tract_lons,
        mode="markers",
        marker=dict(
            size=tract_radii,
            color=tract_colors,
            opacity=0.85,
            sizemode="diameter",
            sizeref=1.0,
        ),
        customdata=list(zip(
            tract_names, tract_cities, tract_tract_ids,
            tract_total_units, tract_affordable_units, tract_property_counts,
            tract_avg_rent_burden, tract_avg_income_pct, tract_avg_poverty_pct, tract_avg_ces4
        )),
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "City: %{customdata[1]}<br>"
            "Tract: %{customdata[2]}<br>"
            "Total Units: %{customdata[3]:,}<br>"
            "Total Affordable Units: %{customdata[4]:,}<br>"
            "Property Count: %{customdata[5]}<br>"
            "Avg Rent Burden: %{customdata[6]:.1%}<br>"
            "Avg Income Percentile: %{customdata[7]:.1%}<br>"
            "Avg Poverty Percentile: %{customdata[8]:.1%}<br>"
            "Avg CES 4.0: %{customdata[9]:.1%}"
            "<extra></extra>"
        ),
        name="Affordable Housing Locations",
        visible=True,
        showlegend=False,
    ))

    # ---- 4. Underserved Tract Labels ----
    underserved_lats = []
    underserved_lons = []
    underserved_texts = []

    for feature in tract_geojson.get("features", []):
        props = feature.get("properties", {})
        tract = clean_tract(props.get("TRACTCE", props.get("TRACT", "")))
        score_data = scores.get(tract)
        if score_data and score_data["underserved"]:
            geom = feature.get("geometry")
            if geom and geom["type"] == "Polygon":
                coords = geom["coordinates"][0]
                lons = [c[0] for c in coords]
                lats = [c[1] for c in coords]
                centroid_lat = sum(lats) / len(lats)
                centroid_lon = sum(lons) / len(lons)
                underserved_lats.append(centroid_lat)
                underserved_lons.append(centroid_lon)
                underserved_texts.append(f"⚠ {tract}")

    fig.add_trace(go.Scattermap(
        lat=underserved_lats,
        lon=underserved_lons,
        mode="text",
        text=underserved_texts,
        textfont=dict(
            size=14,
            color="#D95A00",
            family="Arial Black",
        ),
        textposition="middle center",
        name="Underserved Tracts",
        visible=True,
        showlegend=False,
        hoverinfo="skip",
    ))

    # Layout with fullscreen map (using MapLibre - no API key needed)
    fig.update_layout(
        map=dict(
            style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
            center=dict(lat=33.74, lon=-117.87),
            zoom=9.5,
        ),
        margin=dict(l=0, r=0, t=0, b=0),
        height=900,
        showlegend=False,
        updatemenus=[
            dict(
                type="buttons",
                direction="down",
                x=0.01,
                y=0.99,
                xanchor="left",
                yanchor="top",
                bgcolor="white",
                bordercolor="#ccc",
                borderwidth=1,
                font=dict(size=12, color="#333"),
                buttons=[
                    dict(
                        label="Show All Layers",
                        method="update",
                        args=[{"visible": [True, True, True, True]}],
                    ),
                    dict(
                        label="Hide City Boundaries",
                        method="update",
                        args=[{"visible": [True, False, True, True]}],
                    ),
                ],
                showactive=True,
            ),
        ],
        annotations=[
            dict(
                text="Orange County Housing Justice Map — Bivariate: Housing Need × CES 5.0",
                x=0.5,
                y=1.02,
                xref="paper",
                yref="paper",
                xanchor="center",
                yanchor="bottom",
                showarrow=False,
                font=dict(size=16, color="#1A5276", family="Arial Black"),
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="#1A5276",
                borderwidth=1,
                borderpad=8,
            ),
            dict(
                text=(
                    f"Tracts Analyzed: {len(scores)} | "
                    f"Underserved: {sum(1 for v in scores.values() if v['underserved'])} | "
                    f"Median Need: {median_score:.3f} | "
                    f"Median Units/1k Poverty: {median_units_per_1k:.1f}"
                ),
                x=0.5,
                y=1.08,
                xref="paper",
                yref="paper",
                xanchor="center",
                yanchor="bottom",
                showarrow=False,
                font=dict(size=11, color="#333"),
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="#ccc",
                borderwidth=1,
                borderpad=6,
            ),
        ],
    )

    # Save and open
    fig.write_html(
        OUTPUT_PATH,
        config={"scrollZoom": True, "displayModeBar": True},
        include_plotlyjs="cdn",
        full_html=True,
        auto_open=False,
    )

    # Inject bivariate legend HTML into the saved file (Plotly annotations don't render HTML)
    legend_html = build_bivariate_legend_html(
        fig._score_t1, fig._score_t2, fig._ces_t1, fig._ces_t2, fig._bivariate_colors
    )
    with open(OUTPUT_PATH, "r") as f:
        html_content = f.read()

    # Insert legend before closing </body> tag
    legend_div = f'''
    <div id="bivariate-legend" style="
        position: fixed;
        bottom: 20px;
        left: 20px;
        z-index: 9999;
        background: rgba(255,255,255,0.95);
        border: 1px solid #999;
        border-radius: 6px;
        padding: 12px;
        font-family: Arial, sans-serif;
        font-size: 11px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        max-width: 300px;
    ">
        {legend_html}
    </div>
    '''
    html_content = html_content.replace('</body>', legend_div + '\n</body>')

    with open(OUTPUT_PATH, "w") as f:
        f.write(html_content)

    return fig, len(tract_ids), len(short_properties)


def main():
    fig, tract_count, property_count = build_map()
    webbrowser.open("file://" + os.path.abspath(OUTPUT_PATH))
    print(f"Opened {OUTPUT_PATH} with {tract_count} housing-score tracts and {property_count} affordable housing properties.")


if __name__ == "__main__":
    main()