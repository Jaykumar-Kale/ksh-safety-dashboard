import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="KSH Safety Dashboard", layout="wide")
st.title("KSH – Safety Performance Dashboard")
st.caption("Based on cleaned, auditable incident data")

@st.cache_data
def load_data():
    incidents = pd.read_excel("data/processed/incident_fact.xlsx")
    manhours = pd.read_excel("data/processed/manhours_fact.xlsx")
    return incidents, manhours

df, mh = load_data()

df["Incident_Date"] = pd.to_datetime(df["Incident_Date"])
df["Year"] = df["Incident_Date"].dt.year
df["Month"] = df["Incident_Date"].dt.month

# ---------------- Filters ----------------
st.sidebar.header("Filters")

years = sorted(df["Year"].dropna().unique())
warehouses = sorted(df["Warehouse"].unique())

year_f = st.sidebar.multiselect("Year", years, default=years)
wh_f = st.sidebar.multiselect("Warehouse", warehouses, default=warehouses)

fdf = df[
    df["Year"].isin(year_f) &
    df["Warehouse"].isin(wh_f)
]

# ---------------- KPIs ----------------
total_incidents = len(fdf)
accident_free_days = (
    (datetime.today() - fdf["Incident_Date"].max()).days
    if not fdf.empty else 0
)

k1, k2, k3 = st.columns(3)
k1.metric("Total Cases", total_incidents)
k2.metric("Accident Free Days", accident_free_days)
k3.metric("Warehouses", fdf["Warehouse"].nunique())

# ---------------- Layout ----------------
left, right = st.columns((2, 1))

with left:
    st.subheader("Incident Trend")
    trend = fdf.groupby(["Year", "Month"]).size().reset_index(name="Count")
    st.plotly_chart(
        px.line(trend, x="Month", y="Count", color="Year", markers=True),
        use_container_width=True
    )

    st.subheader("Injuries by Body Part")
    st.plotly_chart(
        px.bar(fdf, x="Body_Part", color="Body_Part"),
        use_container_width=True
    )

with right:
    st.subheader("Incident Classification")
    st.plotly_chart(
        px.pie(
            fdf,
            names="Case_Type",
            hole=0.45,
            color_discrete_sequence=["#ffbf00","#ff7f0e","#d62728","#2ca02c","#9e9e9e"]
        ),
        use_container_width=True
    )

st.success("Dashboard built on analytically correct data")
