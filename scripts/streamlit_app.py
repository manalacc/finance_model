# ======================
# import libraries
import streamlit as st
import pandas as pd
from data_analysis import DataAnalysis

# ======================
# Page configuration
st.set_page_config(
    page_title="Key Financials Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================
# CSS styles
st.markdown("""
<style>
[data-testid="stMetric"] {
    background-color: #393939;
    text-align: center;
    width: 75%;
    padding: 2.5px 0;
    border-radius: 7px;
}

[data-testid="stMetricLabel"] {
    display: flex;
    justify-content: center;
    align-items: center;
}

</style> 
""", unsafe_allow_html=True)

# ======================
# Load data
df = pd.read_excel("../data/Budget_vs_Actuals_Template.xlsx", sheet_name="Actuals & Variance")
analysis = DataAnalysis(df)

# Define proper month order
month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

# ======================
# Prepare data with proper month ordering
df_chart = df.copy()
df_chart['Month'] = pd.Categorical(df_chart['Month'], categories=month_order, ordered=True)
df_chart = df_chart.sort_values('Month')  # Ensure it's sorted by the categorical order
df_chart['Cost Variance'] = df_chart['Actual Cost'] - df_chart['Budget Cost']
df_chart['Gross Margin (%)'] = (df_chart['Actual Revenue'] - df_chart['Actual Cost']) / df_chart['Actual Revenue'] * 100
df_chart['Monthly Cost Variance'] = df_chart['Cost Variance'].diff()

# Calculate variances
revenue_variance = df_chart['Actual Revenue'].sum() - df_chart['Budget Revenue'].sum()
cost_variance = df_chart['Actual Cost'].sum() - df_chart['Budget Cost'].sum()
total_cost_variance = df_chart['Cost Variance'].sum()

# ======================
# Dashboard main panel
col = st.columns((1.5, 4, 2.5), gap='medium')

with col[0]:
    st.subheader("Summary Statistics")
    
    st.metric(label="Total Budget Revenue", value=f"${df_chart['Budget Revenue'].sum():,.0f}")
    st.metric(
        label="Total Actual Revenue",
        value=f"${df_chart['Actual Revenue'].sum():,.0f}",
        delta=f"{revenue_variance:,.0f}",
        )

    st.metric(label="Total Budget Cost", value=f"${df_chart['Budget Cost'].sum():,.0f}")
    st.metric(label="Total Actual Cost", value=f"${df_chart['Actual Cost'].sum():,.0f}", delta=f"{cost_variance:,.0f}", delta_color="inverse" if cost_variance < 0 else "normal")

    st.metric(label="Total Cost Variance", value=f"${df_chart['Cost Variance'].sum():,.0f}")
    st.metric(label="Total Gross Margin (%)", value=f"{df_chart['Gross Margin (%)'].mean():.1f}%")

with col[1]:
    st.header("Key Financial Metrics")    
    with st.container(height=600, gap='medium',):
        with st.container(horizontal=True):
            # Chart selector
            selected_chart = st.selectbox(
                "Display Chart:",
                ["All Charts", "Budget vs Actual Revenue", "Budget vs Actual Cost", 
                "Gross Margin Percentage", "Monthly Cost Variance"],
                index=0
            )

        # Show charts based on selection
        if selected_chart == "All Charts" or selected_chart == "Budget vs Actual Revenue":
            st.subheader("Budget vs Actual Revenue")
            st.line_chart(df_chart.set_index("Month")[["Budget Revenue", "Actual Revenue"]], use_container_width=True)

        if selected_chart == "All Charts" or selected_chart == "Budget vs Actual Cost":
            st.subheader("Budget vs Actual Cost")
            st.line_chart(df_chart.set_index("Month")[["Budget Cost", "Actual Cost"]], use_container_width=True)

        if selected_chart == "All Charts" or selected_chart == "Gross Margin Percentage":
            st.subheader("Gross Margin Percentage")
            st.line_chart(df_chart.set_index("Month")[["Budget Gross Margin (%)", "Actual Gross Margin (%)"]], use_container_width=True)

        if selected_chart == "All Charts" or selected_chart == "Monthly Cost Variance":
            st.subheader("Monthly Cost Variance")
            st.bar_chart(df_chart.set_index("Month")[["Cost Variance"]], use_container_width=True)

with col[2]:
    st.header("")
    with st.expander("Detailed Data View", expanded=True):
        st.write("This section provides a detailed view of the financial data.")
        st.dataframe(df_chart, use_container_width=True, hide_index=True)