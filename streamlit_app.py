import os

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="OC Housing Map", layout="wide")

HERE = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(HERE, "oc_housing_map.html")

st.title("Orange County Affordable Housing Need Map")

if os.path.exists(HTML_PATH):
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        components.html(f.read(), height=800, scrolling=False)
else:
    st.error(
        "oc_housing_map.html is missing. Run `python oc_housing_map.py` locally "
        "and commit the generated oc_housing_map.html to the repo."
    )
