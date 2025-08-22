import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook, load_workbook

class DataAnalysis:
    def __init__(self, df):
        self.df = df

    def get_key_totals(self):
        return {
            "total_budget_rev": self.df['Budget Revenue'].sum(),
            "total_actual_rev": self.df['Actual Revenue'].sum(),
            "total_budget_exp": self.df['Budget Cost'].sum(),
            "total_actual_exp": self.df['Actual Cost'].sum(),
            "total_profit": self.df['Actual Revenue'].sum() - self.df['Actual Cost'].sum()
        }

    def get_variance_analysis(self):
        return {
            "best_month": self.df.loc[self.df['Variance ($)'].idxmax(), 'Month'],
            "best_val": self.df['Variance ($)'].max(),
            "worst_month": self.df.loc[self.df['Variance ($)'].idxmin(), 'Month'],
            "worst_val": self.df['Variance ($)'].min()
        }
    
    def get_top_below_budget_revenue(self):
        below_budget_rev = self.df[self.df["Actual Revenue"] < self.df["Budget Revenue"]]
        return below_budget_rev.nsmallest(3, "Actual Revenue")
    
    def get_top_cost_overruns(self):
        cost_overruns = self.df[self.df["Actual Cost"] > self.df["Budget Cost"]]
        return cost_overruns.nlargest(3, "Actual Cost")

    def get_largest_gross_margin_drops(self):
        self.df["Gross Margin Drop"] = self.df["Budget Gross Margin (%)"] - self.df["Actual Gross Margin (%)"]
        return self.df.nlargest(3, "Gross Margin Drop")

def main():
    wb = Workbook()
    ws = wb.active
    if ws is not None:
        ws.title = "Data Analysis"
    else:
        raise ValueError("No active worksheet found in the workbook.")

    df = pd.read_excel("../data/Budget_vs_Actuals_Template.xlsx", sheet_name="Actuals & Variance")
    analysis = DataAnalysis(df)

    key_totals = analysis.get_key_totals()
    variance_analysis = analysis.get_variance_analysis()

    # Key totals
    total_budget_rev = key_totals['total_budget_rev']
    total_actual_rev = key_totals['total_actual_rev']
    total_budget_exp = key_totals['total_budget_exp']
    total_actual_exp = key_totals['total_actual_exp']
    total_profit = total_actual_rev - total_actual_exp

    # Variance extremes
    best_month = variance_analysis['best_month']
    best_val = variance_analysis['best_val']
    worst_month = variance_analysis['worst_month']
    worst_val = variance_analysis['worst_val']

    # top 3 months where actual revenue was below budget
    # top_3_below_rev = analysis.get_top_below_budget_revenue()

    # top 3 cost overruns
    # top_3_cost_overruns = analysis.get_top_cost_overruns()

    # months with the largest gross margin drops
    # largest_gross_margin_drops = analysis.get_largest_gross_margin_drops()

    # make_chart(df, title="Budget vs Actual Revenue", x_label="Month", y_label="Revenue ($)", metrics=["Budget Revenue", "Actual Revenue"], chart_type='line')

    # line chart: gross margin %
    # make_chart(df, title="Gross Margin Percentage", x_label="Month", y_label="Gross Margin (%)", metrics=["Budget Gross Margin (%)", "Actual Gross Margin (%)"], chart_type='line', label_format="{:.1f}%")
    
    # bar chart: monthly cost variance
    """
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

    """
    
    summary = f"""
    For the period, total revenue was ${total_actual_rev:,.0f} 
    vs a budget of ${total_budget_rev:,.0f}, a variance of {((total_actual_rev-total_budget_rev)/total_budget_rev)*100:.1f}%.
    Total expenses were ${total_actual_exp:,.0f} vs a budget of ${total_budget_exp:,.0f}.
    Net profit reached ${total_profit:,.0f}.
    The best revenue month was {best_month} (+${best_val:,.0f} variance) 
    and the worst was {worst_month} ({worst_val:,.0f} variance).
    """
    # print(summary)
    return 0

def make_chart(df, title, x_label, y_label, metrics, chart_type='line', label_format="{:.0f}"):
    plt.figure(figsize=(12, 7))
    plt.plot(df[x_label], df[metrics[0]], label=metrics[0], marker="o", markersize=8, linewidth=2)
    plt.plot(df[x_label], df[metrics[1]], label=metrics[1], marker="s", markersize=8, linewidth=2)

    # Add labels to points
    for i, row in df.iterrows():
        # Actual labels
        actual_label = f"${row[metrics[1]]:,.0f}"
        plt.annotate(actual_label, 
                    (row[x_label], row[metrics[1]]), 
                    textcoords="offset points", 
                    xytext=(0,-15), 
                    ha='center',
                    fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", facecolor="lightcoral", alpha=0.7))
    
    plt.title(title)
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    plt.xticks(rotation=45)
    plt.legend()
    plt.tight_layout()
    # save chart title so that there are no underscores or capitals
    title_clean = title.replace(" ", "_").lower()
    plt.savefig(f"../charts/{title_clean}.png", dpi=300, bbox_inches='tight')
    plt.close()

if __name__ == "__main__":
    main()
