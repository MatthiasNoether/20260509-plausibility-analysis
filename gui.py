import streamlit as st
import os

# 1. Get the exact folder path where this gui.py file lives
base_dir = os.path.dirname(os.path.abspath(__file__))

# Page configuration
st.set_page_config(page_title="Plausibility Dashboard", layout="wide")

# Sidebar Navigation
st.sidebar.title("Dashboard")
pages = [
    "Populations",
    "Births",
    "Deaths",
    "Fertility",
    "Life Expectancy",
    "Migrations",
    "Emigrations",
]
page = st.sidebar.radio("Navigation", pages)
st.title("Plausibilitäts-Check: Demographie")
st.info("Diese Übersicht zeigt die Kennzahlen im Vergleich zum Durchschnitt (rote Linie).")
st.caption("Die Werte sollten den roten Zellen in den Excel-Sheets entsprechen.")

# Content logic for Migrants
# Content logic for Migrants: show migrant and native plots side-by-side (smaller)
page_map = {
    "Populations": ("populations", "Populations"),
    "Births": ("births", "Births"),
    "Deaths": ("deaths", "Deaths"),
    "Fertility": ("fertility_calculated", "Fertility (calculated)"),
    "Life Expectancy": ("life_expectancy", "Life Expectancy"),
    "Migrations": ("migrations", "Migrations"),
    "Emigrations": ("emigrations", "Emigrations"),
}

def show_side_by_side(safe_label: str, title: str, img_width: int = 450) -> None:
    left_img = os.path.join(base_dir, "plots", f"{safe_label}_migrant.png")
    right_img = os.path.join(base_dir, "plots", f"{safe_label}_native.png")
    st.subheader(f"{title} — Migrant vs Native")
    col1, col2 = st.columns(2)
    if os.path.exists(left_img):
        col1.caption("Migrant")
        col1.image(left_img, use_column_width=False, width=img_width)
    else:
        col1.info(f"Migrant plot not found: {left_img}")
    if os.path.exists(right_img):
        col2.caption("Native")
        col2.image(right_img, use_column_width=False, width=img_width)
    else:
        col2.info(f"Native plot not found: {right_img}")

if page in page_map:
    safe_label, title = page_map[page]
    show_side_by_side(safe_label, title, img_width=450)

# Native-only displays removed — both migrant and native plots are shown side-by-side above
# Native-only displays removed — both plots are shown side-by-side above
st.divider()
st.caption("Erstellt am 09. Mai 2026")