import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook

def main():
    wb = Workbook()
    ws = wb.active
    if ws is not None:
        ws.title = "Data Analysis"
    else:
        raise ValueError("No active worksheet found in the workbook.")

    # Perform data analysis
    df = pd.read_excel("../data/Budget_vs_Actuals_Template.xlsx", sheet_name="Actuals & Variance")

    # top 3 months where actual revenue was below budget
    below_budget_rev = df[df["Actual Revenue"] < df["Budget Revenue"]]
    top_3_below_rev = below_budget_rev.nsmallest(3, "Actual Revenue")

    # top 3 cost overruns
    cost_overruns = df[df["Actual Cost"] > df["Budget Cost"]]
    top_3_cost_overruns = cost_overruns.nlargest(3, "Actual Cost")

    # months with the largest gross margin drops
    df["Gross Margin Drop"] = df["Budget Gross Margin (%)"] - df["Actual Gross Margin (%)"]
    largest_gross_margin_drops = df.nlargest(3, "Gross Margin Drop")

    # print(top_3_below_rev)
    # print(top_3_cost_overruns)
    # print(largest_gross_margin_drops)

    # line chart: budget vs actual revenue
    plt.figure(figsize=(12, 7))
    budget_line = plt.plot(df["Month"], df["Budget Revenue"], label="Budget Revenue", marker="o", markersize=8, linewidth=2)
    actual_line = plt.plot(df["Month"], df["Actual Revenue"], label="Actual Revenue", marker="s", markersize=8, linewidth=2)
    
    # Add labels to points
    for i, row in df.iterrows():
        # Actual Revenue labels
        actual_label = f"${row['Actual Revenue']:,.0f}"
        plt.annotate(actual_label, 
                    (row["Month"], row["Actual Revenue"]), 
                    textcoords="offset points", 
                    xytext=(0,-15), 
                    ha='center',
                    fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", facecolor="lightcoral", alpha=0.7))
    
    plt.title("Budget vs Actual Revenue")
    plt.xlabel("Month")
    plt.ylabel("Revenue ($)")
    plt.xticks(rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig("../charts/budget_vs_actual_revenue.png", dpi=300, bbox_inches='tight')
    plt.close()

    # line chart: gross margin %
    plt.figure(figsize=(12, 7))
    plt.plot(df["Month"], df["Budget Gross Margin (%)"], label="Budget Gross Margin (%)", marker="o", markersize=8)
    plt.plot(df["Month"], df["Actual Gross Margin (%)"], label="Actual Gross Margin (%)", marker="s", markersize=8)
    
    # Add labels to points BEFORE saving the chart
    for i, row in df.iterrows():
        # Label for Actual Gross Margin
        actual_label = f"{row['Actual Gross Margin (%)']:.1f}%"
        plt.annotate(actual_label, 
                    (row["Month"], row["Actual Gross Margin (%)"]), 
                    textcoords="offset points", 
                    xytext=(0,-15), 
                    ha='center',
                    fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", facecolor="lightcoral", alpha=0.7))
    
    plt.title("Gross Margin %: Budget vs Actual")
    plt.xlabel("Month")
    plt.ylabel("Gross Margin (%)")
    plt.xticks(rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig("../charts/gross_margin_percentage.png", dpi=300, bbox_inches='tight')
    plt.close()

    # bar chart: monthly cost variance
    plt.figure(figsize=(12, 7))
    cost_variance = df["Actual Cost"] - df["Budget Cost"]
    plt.bar(df["Month"], cost_variance, color="lightcoral")
    plt.axhline(0, color="black", linewidth=0.8, linestyle="--")
    plt.title("Monthly Cost Variance")
    plt.xlabel("Month")
    plt.ylabel("Cost Variance ($)")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("../charts/monthly_cost_variance.png", dpi=300, bbox_inches='tight')
    plt.close()

    # pie chart: cost breakdown
    plt.figure(figsize=(8, 8))
    cost_breakdown = df[["Month", "Actual Cost", "Budget Cost"]].set_index("Month")
    cost_breakdown.plot.pie(y="Actual Cost", autopct="%1.1f%%", startangle=90, legend=False, ax=plt.gca())
    plt.title("Cost Breakdown: Actual vs Budget")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig("../charts/cost_breakdown.png", dpi=300, bbox_inches='tight')
    plt.close()

    return 0

if __name__ == "__main__":
    main()
