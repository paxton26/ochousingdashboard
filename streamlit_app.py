"""Streamlit app for Orange County Affordable Housing Need Map.

Run with: streamlit run streamlit_app.py
"""

import hashlib
import json
import math
import os
from collections import defaultdict

import openpyxl
import plotly.graph_objects as go
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Orange County Affordable Housing Need Map",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Data file paths
XLSX_PATH = os.path.join(BASE_DIR, "HousingCentersOCFinalData.xlsx")
GEOJSON_PATH = os.path.join(BASE_DIR, "oc_tracts.geojson")
CES_PATH = os.path.join(BASE_DIR, "ces50_oc.json")
POVERTY_PATH = os.path.join(BASE_DIR, "poverty_oc.json")
TRACT_CENTROIDS_PATH = os.path.join(BASE_DIR, "tract_centroids.json")
CITY_BOUNDARIES_PATH = os.path.join(BASE_DIR, "oc_cities.geojson")


# ──────────────────────────────────────────────
# Utility functions
# ──────────────────────────────────────────────
def clean_tract(value: str) -> str:
    """Clean tract code to 6-digit format."""
    text = str(value).strip()
    if "." in text:
        text = text.split(".")[0]
    return text.zfill(6)


def number(value):
    """Safely convert to float."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def need_color(score: float) -> str:
    """Color for housing need score (standalone legend)."""
    if score >= 0.80:
        return "#D95A00"
    if score >= 0.70:
        return "#E8870A"
    if score >= 0.60:
        return "#F0A500"
    return "#2E86C1"


# ──────────────────────────────────────────────
# Data loading (cached)
# ──────────────────────────────────────────────
@st.cache_data(show_spinner="Loading housing data...")
def load_housing_scores():
    """Load and compute housing need scores from Excel + poverty data."""
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

    with open(POVERTY_PATH, "r") as f:
        poverty_data = json.load(f)

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

    # Aggregate by tract
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

    # Determine underserved tracts (high need, low units per poverty)
    median_score = sorted(item["score"] for item in scores.values())[len(scores) // 2]
    median_units_per_1k = sorted(item["units_per_1k_poverty"] for item in scores.values())[len(scores) // 2]
    for item in scores.values():
        item["underserved"] = (
            item["score"] > median_score and
            item["units_per_1k_poverty"] < median_units_per_1k
        )
    return scores, median_score, median_units_per_1k


@st.cache_data(show_spinner="Loading property data...")
def load_short_tab_properties():
    """Load individual property locations from Excel."""
    if not os.path.exists(TRACT_CENTROIDS_PATH):
        return []

    with open(TRACT_CENTROIDS_PATH, "r") as f:
        tract_centroids = json.load(f)

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


@st.cache_data(show_spinner="Loading geographic data...")
def load_geojson_data():
    """Load tract and city boundary GeoJSON files."""
    with open(GEOJSON_PATH, "r") as f:
        tract_geojson = json.load(f)
    with open(CES_PATH, "r") as f:
        ces50 = json.load(f)
    with open(CITY_BOUNDARIES_PATH, "r") as f:
        city_geojson = json.load(f)
    return tract_geojson, ces50, city_geojson


# ──────────────────────────────────────────────
# Bivariate color scheme & legend
# ──────────────────────────────────────────────
def build_bivariate_colors():
    """3×3 bivariate color matrix: CES (rows) × Housing Need (cols)."""
    return [
        # CES Low (green-ish)
        ["#e8f5e9", "#a5d6a7", "#66bb6a"],   # Need: Low, Med, High
        # CES Medium (yellow-ish)
        ["#fff9c4", "#ffeb3b", "#fdd835"],   # Need: Low, Med, High
        # CES High (red-ish)
        ["#ffccbc", "#ff8a65", "#ef5350"],   # Need: Low, Med, High
    ]


def build_bivariate_legend_html(score_t1, score_t2, ces_t1, ces_t2, colors):
    """Generate HTML for a 3×3 bivariate legend grid."""
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


# ──────────────────────────────────────────────
# Map building
# ──────────────────────────────────────────────
def build_map(tract_geojson, ces50, scores, short_properties, city_geojson,
              show_city_boundaries=True, show_properties=True, show_underserved_labels=True):
    """Build the Plotly figure with all layers."""
    # ── Prepare tract-level arrays ──
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

    # ── Calculate tertile breaks for bivariate choropleth ──
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

    # ── Assign each tract to a bivariate bin (0-8) ──
    bivariate_colors = build_bivariate_colors()
    tract_bin_indices = []

    for score, ces in zip(tract_scores, tract_ces50):
        if ces is None:
            tract_bin_indices.append(-1)
            continue

        # Need bin (0, 1, 2)
        if score <= score_t1:
            need_bin = 0
        elif score <= score_t2:
            need_bin = 1
        else:
            need_bin = 2

        # CES bin (0, 1, 2)
        if ces <= ces_t1:
            ces_bin = 0
        elif ces <= ces_t2:
            ces_bin = 1
        else:
            ces_bin = 2

        bin_idx = ces_bin * 3 + need_bin
        tract_bin_indices.append(bin_idx)

    # ── Create figure ──
    fig = go.Figure()

    # ── 1. Bivariate Choropleth (Housing Need × CES 5.0) ──
    fig.add_trace(go.Choroplethmap(
        geojson=tract_geojson,
        locations=tract_ids,
        z=tract_bin_indices,
        featureidkey="properties.GEOID",
        colorscale=[[i/8, bivariate_colors[i//3][i%3]] for i in range(9)],
        zmin=0,
        zmax=8,
        marker_opacity=0.75,
        marker_line_width=0.5,
        marker_line_color="#555555",
        showscale=False,
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

    # ── 2. City Boundaries (toggleable) ──
    if show_city_boundaries:
        city_line_lats = []
        city_line_lons = []
        city_line_names = []

        for feature in city_geojson.get("features", []):
            city = feature["properties"].get("city", "Unknown")
            geom = feature.get("geometry")
            if geom and geom["type"] in ("Polygon", "MultiPolygon"):
                if geom["type"] == "Polygon":
                    for ring in geom["coordinates"]:
                        lons = [c[0] for c in ring]
                        lats = [c[1] for c in ring]
                        city_line_lons.extend(lons)
                        city_line_lats.extend(lats)
                        city_line_names.extend([city] * len(lons))
                        city_line_lons.append(None)
                        city_line_lats.append(None)
                        city_line_names.append(None)
                elif geom["type"] == "MultiPolygon":
                    for polygon in geom["coordinates"]:
                        for ring in polygon:
                            lons = [c[0] for c in ring]
                            lats = [c[1] for c in ring]
                            city_line_lons.extend(lons)
                            city_line_lats.extend(lats)
                            city_line_names.extend([city] * len(lons))
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

    # ── 3. Affordable Housing Locations ──
    if show_properties and short_properties:
        props_by_tract = defaultdict(list)
        for prop in short_properties:
            if prop["lat"] is not None and prop["lon"] is not None:
                props_by_tract[prop["tract"]].append(prop)

        # Build tract centroids from GeoJSON
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
        tract_names = []
        tract_cities_list = []
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
            lat, lon = tract_centroids.get(tract, (props[0]["lat"], props[0]["lon"]))

            total_props = len(props)
            total_units = sum(p["total_units"] for p in props)
            total_affordable = sum(p["affordable_units"] for p in props)

            rb_vals = [p["rent_burden"] for p in props if p["rent_burden"] is not None]
            inc_vals = [p["income_percentile"] for p in props if p["income_percentile"] is not None]
            pov_vals = [p["poverty_percentile"] for p in props if p["poverty_percentile"] is not None]
            ces4_vals = [p["ces4_percentile"] for p in props if p["ces4_percentile"] is not None]

            avg_rent_burden = sum(rb_vals) / len(rb_vals) if rb_vals else None
            avg_income = sum(inc_vals) / len(inc_vals) if inc_vals else None
            avg_poverty = sum(pov_vals) / len(pov_vals) if pov_vals else None
            avg_ces4 = sum(ces4_vals) / len(ces4_vals) if ces4_vals else None

            tract_lats.append(lat)
            tract_lons.append(lon)
            tract_names.append(f"Tract {tract} ({total_props} propert{'y' if total_props == 1 else 'ies'})")
            tract_cities_list.append(props[0]["city"])
            tract_tract_ids.append(tract)
            tract_total_units.append(total_units)
            tract_affordable_units.append(total_affordable)
            tract_property_counts.append(total_props)
            tract_avg_rent_burden.append(avg_rent_burden)
            tract_avg_income_pct.append(avg_income)
            tract_avg_poverty_pct.append(avg_poverty)
            tract_avg_ces4.append(avg_ces4)
            radius = max(8, min(30, 8 + total_affordable * 0.10))
            tract_radii.append(radius)
            tract_colors.append("#D95A00")

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
                tract_names, tract_cities_list, tract_tract_ids,
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

    # ── 4. Underserved Tract Labels ──
    if show_underserved_labels:
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

        if underserved_lats:
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

    # ── Layout ──
    fig.update_layout(
        map=dict(
            style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
            center=dict(lat=33.74, lon=-117.87),
            zoom=9.5,
        ),
        margin=dict(l=0, r=0, t=0, b=0),
        height=800,
        showlegend=False,
    )

    return fig, score_t1, score_t2, ces_t1, ces_t2, bivariate_colors


# ──────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────
def main():
    st.title("🏠 Orange County Affordable Housing Need Map")
    st.caption("Bivariate choropleth: Housing Need Score × CalEnviroScreen 5.0 | Orange County, CA")

    # Load data
    with st.spinner("Loading data..."):
        scores, median_score, median_units_per_1k = load_housing_scores()
        short_properties = load_short_tab_properties()
        tract_geojson, ces50, city_geojson = load_geojson_data()

    # Sidebar controls
    with st.sidebar:
        st.header("🎛️ Map Controls")

        show_city_boundaries = st.checkbox("Show City Boundaries", value=True)
        show_properties = st.checkbox("Show Affordable Housing Locations", value=True)
        show_underserved_labels = st.checkbox("Show Underserved Tract Labels (⚠)", value=True)

        st.divider()

        # Filters
        st.subheader("🔍 Filters")

        # City filter
        all_cities = sorted(set(s["city"] for s in scores.values()))
        selected_cities = st.multiselect(
            "Filter by City",
            options=all_cities,
            default=all_cities,
            help="Select cities to display"
        )

        # Score range filter
        score_min = min(s["score"] for s in scores.values())
        score_max = max(s["score"] for s in scores.values())
        score_range = st.slider(
            "Housing Need Score Range",
            min_value=float(score_min),
            max_value=float(score_max),
            value=(float(score_min), float(score_max)),
            step=0.01,
            format="%.2f"
        )

        # Underserved only
        show_underserved_only = st.checkbox("Show Only Underserved Tracts", value=False)

        st.divider()

        # Stats
        st.subheader("📊 Summary Stats")
        total_tracts = len(scores)
        underserved_count = sum(1 for v in scores.values() if v["underserved"])
        total_units = sum(v["units"] for v in scores.values())
        total_props = len(short_properties)

        st.metric("Tracts Analyzed", total_tracts)
        st.metric("Underserved Tracts", underserved_count, f"{underserved_count/total_tracts*100:.1f}%")
        st.metric("Median Need Score", f"{median_score:.3f}")
        st.metric("Total Affordable Units", f"{total_units:,}")
        st.metric("Properties Mapped", total_props)

    # Filter scores based on sidebar selections
    filtered_scores = {
        tract: data for tract, data in scores.items()
        if data["city"] in selected_cities
        and score_range[0] <= data["score"] <= score_range[1]
        and (not show_underserved_only or data["underserved"])
    }

    # Filter properties
    filtered_properties = [
        p for p in short_properties
        if p["city"] in selected_cities
    ]

    # Build map
    fig, score_t1, score_t2, ces_t1, ces_t2, bivariate_colors = build_map(
        tract_geojson=tract_geojson,
        ces50=ces50,
        scores=filtered_scores,
        short_properties=filtered_properties,
        city_geojson=city_geojson,
        show_city_boundaries=show_city_boundaries,
        show_properties=show_properties,
        show_underserved_labels=show_underserved_labels,
    )

    # Display map
    st.plotly_chart(fig, use_container_width=True, config={"scrollZoom": True, "displayModeBar": True})

    # Bivariate legend (as HTML in an expander)
    with st.expander("📖 Bivariate Legend: Housing Need × CES 5.0", expanded=False):
        legend_html = build_bivariate_legend_html(score_t1, score_t2, ces_t1, ces_t2, bivariate_colors)
        st.markdown(legend_html, unsafe_allow_html=True)

        st.markdown("""
        **How to read this legend:**
        - **Rows (↓)**: CalEnviroScreen 5.0 percentile (environmental burden) — Low → High
        - **Columns (→)**: Housing Need Score (rent burden, income, poverty) — Low → High
        - **Top-right (green)**: Low environmental burden, low housing need
        - **Bottom-right (red)**: High environmental burden, high housing need — *priority areas*
        """)

    # Data table
    with st.expander("📋 Tract Data Table", expanded=False):
        import pandas as pd
        df = pd.DataFrame([
            {
                "Tract": v["tract"],
                "City": v["city"],
                "Housing Need Score": v["score"],
                "Affordable Units": v["units"],
                "Programs": v["programs"],
                "Below Poverty": v["below_poverty"],
                "Units/1k Poverty": v["units_per_1k_poverty"],
                "Underserved": "Yes" if v["underserved"] else "No",
            }
            for v in filtered_scores.values()
        ])
        st.dataframe(df.sort_values("Housing Need Score", ascending=False), use_container_width=True, hide_index=True)

    # Footer
    st.divider()
    st.caption(
        "Data sources: HUD Affordable Housing Data, ACS 5-Year Estimates, CalEnviroScreen 5.0 | "
        "Built with Streamlit + Plotly"
    )


if __name__ == "__main__":
    main()