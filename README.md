# Orange County Affordable Housing Need Map

Interactive Streamlit app visualizing affordable housing need across Orange County census tracts using a bivariate choropleth (Housing Need Score × CalEnviroScreen 5.0).

## Features

- **Bivariate Choropleth**: Housing Need Score (rent burden, income, poverty) × CES 5.0 (environmental burden)
- **Interactive Filters**: Filter by city, score range, and underserved status
- **Layer Toggles**: City boundaries, affordable housing locations, underserved tract labels
- **Hover Details**: Rich tooltips with tract-level metrics
- **Data Table**: Sortable table of all tract scores
- **Cached Data Loading**: Fast reloads with `@st.cache_data`

## Quick Start (Local)

```bash
cd mapv1
pip install -r requirements.txt
streamlit run streamlit_app.py
```

The app will open at `http://localhost:8501`

## Deploy to Streamlit Cloud

1. **Push to GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/yourusername/oc-housing-map.git
   git push -u origin main
   ```

2. **Deploy on Streamlit Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Click "New app"
   - Connect your GitHub repo
   - Main file path: `mapv1/streamlit_app.py`
   - Click "Deploy!"

3. **No secrets needed** — all data files are included in the repo.

## Data Files (included in repo)

| File | Description |
|------|-------------|
| `HousingCentersOCFinalData.xlsx` | HUD affordable housing inventory (Sheet: "Short") |
| `oc_tracts.geojson` | Orange County census tract boundaries |
| `oc_cities.geojson` | Orange County city boundaries |
| `ces50_oc.json` | CalEnviroScreen 5.0 percentiles by tract |
| `poverty_oc.json` | Poverty counts by tract (ACS) |
| `tract_centroids.json` | Tract centroid coordinates for property markers |

## Methodology

**Housing Need Score** (per tract, averaged across programs):
```
Score = Rent_Burden_% × 0.40 + (1 - Median_HH_Income_Percentile) × 0.25 + Poverty_Percentile × 0.35
```

**Underserved Tract**: High need (above median score) AND low affordable units per 1,000 people in poverty (below median)

**Bivariate Classification**: Tertile splits on both Housing Need Score and CES 5.0 → 3×3 grid

## Project Structure

```
mapv1/
├── streamlit_app.py          # Main Streamlit app
├── oc_housing_map.py         # Original standalone script (generates HTML)
├── requirements.txt          # Python dependencies
├── .streamlit/
│   └── config.toml           # Streamlit config
├── HousingCentersOCFinalData.xlsx
├── oc_tracts.geojson
├── oc_cities.geojson
├── ces50_oc.json
├── poverty_oc.json
└── tract_centroids.json
```

## Customization

- **Colors**: Edit `build_bivariate_colors()` in `streamlit_app.py`
- **Weights**: Modify score formula in `load_housing_scores()`
- **Map Style**: Change `map.style` in `fig.update_layout()` (any MapLibre style URL)
- **Default View**: Adjust `center` and `zoom` in map layout

## License

Data from public sources (HUD, Census ACS, CalEPA). Code: MIT.