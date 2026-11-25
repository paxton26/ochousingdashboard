import os

import streamlit as st
import streamlit.components.v1 as components

import oc_housing_map

st.set_page_config(page_title="OC Housing Map", layout="wide")


@st.cache_data(show_spinner="Building map...")
def get_map_html():
    # build_map() writes the finished HTML (including the injected legend) to OUTPUT_PATH
    oc_housing_map.build_map()
    with open(oc_housing_map.OUTPUT_PATH, "r") as f:
        return f.read()


st.title("Orange County Affordable Housing Need Map")
components.html(get_map_html(), height=800, scrolling=False)
