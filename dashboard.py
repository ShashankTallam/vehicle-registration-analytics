import streamlit as st
import pandas as pd
import plotly.express as px

# ===== LOAD DATA =====
@st.cache_data
def load_data():
    df_month = pd.read_excel("vehicle_growth_metrics.xlsx", sheet_name="Monthly_with_Growth")  # monthly file
    df_quarter = pd.read_excel("vehicle_growth_metrics_quarterly.xlsx", sheet_name="Quarterly_Growth")  # quarterly file
    return df_month, df_quarter

df_month, df_quarter = load_data()

# ===== SIDEBAR FILTERS =====
st.sidebar.header("Filters")
view_type = st.sidebar.radio("Select view", ["Monthly", "Quarterly"])
vehicle_types = st.sidebar.multiselect("Vehicle Type", options=df_month['vehicle_type'].unique(), default=df_month['vehicle_type'].unique())
manufacturers = st.sidebar.multiselect("Manufacturer", options=df_month['Maker'].unique(), default=df_month['Maker'].unique())

if view_type == "Monthly":
    # Convert pandas Timestamp to Python datetime
    min_date = df_month['date'].min().to_pydatetime().date()
    max_date = df_month['date'].max().to_pydatetime().date()
    date_range = st.sidebar.date_input(
        "Date Range", 
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
    # Handle single date selection
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = date_range
else:
    min_year, max_year = int(df_quarter['year'].min()), int(df_quarter['year'].max())
    year_range = st.sidebar.slider("Year Range", min_year, max_year, (min_year, max_year))

# ===== FILTER DATA =====
if view_type == "Monthly":
    # Convert date_range back to pandas datetime for filtering
    start_date_pd = pd.to_datetime(start_date)
    end_date_pd = pd.to_datetime(end_date)
    
    df_filtered = df_month[
        (df_month['vehicle_type'].isin(vehicle_types)) &
        (df_month['Maker'].isin(manufacturers)) &
        (df_month['date'] >= start_date_pd) &
        (df_month['date'] <= end_date_pd)
    ]
else:
    df_filtered = df_quarter[
        (df_quarter['vehicle_type'].isin(vehicle_types)) &
        (df_quarter['Maker'].isin(manufacturers)) &
        (df_quarter['year'].between(year_range[0], year_range[1]))
    ]

# ===== KPI METRICS =====
total_regs = df_filtered['registrations'].sum()
avg_yoy = df_filtered['YoY_growth_%'].replace("N/A", None).astype(float).mean() if 'YoY_growth_%' in df_filtered.columns else None
avg_qoq = df_filtered['QoQ_growth_%'].replace("N/A", None).astype(float).mean() if 'QoQ_growth_%' in df_filtered.columns else None

col1, col2, col3 = st.columns(3)
col1.metric("Total Registrations", f"{total_regs:,.0f}")
if avg_yoy is not None:
    col2.metric("Avg YoY Growth %", f"{avg_yoy:.2f}%")
if avg_qoq is not None:
    col3.metric("Avg QoQ Growth %", f"{avg_qoq:.2f}%")

# ===== CHARTS =====
st.subheader(f"{view_type} Trends")

if view_type == "Monthly":
    fig1 = px.line(df_filtered, x="date", y="registrations", color="vehicle_type", title="Registrations Over Time")
else:
    fig1 = px.line(df_filtered, x="year_quarter", y="registrations", color="vehicle_type", title="Registrations Over Time")
st.plotly_chart(fig1, use_container_width=True)

# YoY Chart (if available)
if 'YoY_growth_%' in df_filtered.columns:
    # Filter out non-numeric values before grouping
    yoy_numeric = df_filtered[df_filtered['YoY_growth_%'] != "N/A"].copy()
    if not yoy_numeric.empty:
        yoy_numeric['YoY_growth_%'] = pd.to_numeric(yoy_numeric['YoY_growth_%'], errors='coerce')
        yoy_df = yoy_numeric.groupby("Maker")['YoY_growth_%'].mean().reset_index()
        yoy_df = yoy_df.dropna()  # Remove any NaN values
        if not yoy_df.empty:
            fig2 = px.bar(yoy_df.sort_values("YoY_growth_%", ascending=False).head(15), 
                         x="Maker", y="YoY_growth_%", title="Top 15 Manufacturers by Avg YoY Growth %")
            st.plotly_chart(fig2, use_container_width=True)

# QoQ Chart (if available)
if 'QoQ_growth_%' in df_filtered.columns:
    # Filter out non-numeric values before grouping
    qoq_numeric = df_filtered[df_filtered['QoQ_growth_%'] != "N/A"].copy()
    if not qoq_numeric.empty:
        qoq_numeric['QoQ_growth_%'] = pd.to_numeric(qoq_numeric['QoQ_growth_%'], errors='coerce')
        qoq_df = qoq_numeric.groupby("Maker")['QoQ_growth_%'].mean().reset_index()
        qoq_df = qoq_df.dropna()  # Remove any NaN values
        if not qoq_df.empty:
            fig3 = px.bar(qoq_df.sort_values("QoQ_growth_%", ascending=False).head(15), 
                         x="Maker", y="QoQ_growth_%", title="Top 15 Manufacturers by Avg QoQ Growth %")
            st.plotly_chart(fig3, use_container_width=True)

# ===== MARKET SHARE =====
market_share = df_filtered.groupby("Maker")['registrations'].sum().reset_index()
fig4 = px.pie(market_share, values='registrations', names='Maker', title='Market Share by Registrations')
st.plotly_chart(fig4, use_container_width=True)

# ===== DATA TABLE =====
st.subheader("Filtered Data")
st.dataframe(df_filtered)