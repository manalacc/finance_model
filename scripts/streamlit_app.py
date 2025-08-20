import streamlit as st
import pandas as pd
from data_analysis import DataAnalysis

df = pd.read_excel("../data/Budget_vs_Actuals_Template.xlsx", sheet_name="Actuals & Variance")
analysis = DataAnalysis(df)

# Define proper month order
month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

with st.sidebar:
    st.header("📊 Chart Navigation")
    
    # Chart selector
    selected_chart = st.selectbox(
        "Display Chart:",
        ["All Charts", "Budget vs Actual Revenue", "Budget vs Actual Cost", 
         "Gross Margin Percentage", "Monthly Cost Variance"],
        index=0
    )
    
    st.markdown("---")  # Separator line
    
    st.header("Summary Statistics")
    df_chart = df.copy()
    df_chart['Month'] = pd.Categorical(df_chart['Month'], categories=month_order, ordered=True)
    df_chart['Cost Variance'] = df_chart['Actual Cost'] - df_chart['Budget Cost']
    df_chart['Gross Margin (%)'] = (df_chart['Actual Revenue'] - df_chart['Actual Cost']) / df_chart['Actual Revenue'] * 100
    
    st.write("Total Budget Revenue:", f"${df_chart['Budget Revenue'].sum():,.0f}")
    st.write("Total Actual Revenue:", f"${df_chart['Actual Revenue'].sum():,.0f}")
    st.write("Total Budget Cost:", f"${df_chart['Budget Cost'].sum():,.0f}")
    st.write("Total Actual Cost:", f"${df_chart['Actual Cost'].sum():,.0f}")
    st.write("Total Gross Margin (%):", f"{df_chart['Gross Margin (%)'].mean():.1f}%")

# Main content area with conditional chart display
with st.container():
    st.header("Key Financial Metrics")

    df_chart = df.copy()
    df_chart['Month'] = pd.Categorical(df_chart['Month'], categories=month_order, ordered=True)
    df_chart['Cost Variance'] = df_chart['Actual Cost'] - df_chart['Budget Cost']
    df_chart['Gross Margin (%)'] = (df_chart['Actual Revenue'] - df_chart['Actual Cost']) / df_chart['Actual Revenue'] * 100
    df_chart['Monthly Cost Variance'] = df_chart['Cost Variance'].diff()

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
        st.bar_chart(df_chart.set_index("Month")[["Monthly Cost Variance"]], use_container_width=True)
