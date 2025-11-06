# retention_trend.py
# -------------------------------------------------------
# Calculates and visualizes dropout, churn, and retention
# trends across 12 months using cohort and rolling analysis.
# -------------------------------------------------------

import matplotlib
matplotlib.use("TkAgg")  # ✅ ensures windows stay interactive

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

plt.ion()  # ✅ interactive mode so multiple charts stay open


def run_retention_trend():
    FILE_PATH = "Documents/Mar24_Mar25_Cleansed.xlsx"
    SHEET = "Main"

    if not os.path.exists(FILE_PATH):
        print(f"❌ File not found: {FILE_PATH}")
        return

    # ----------------------------------------------------
    # 1. Load and Prepare Data
    # ----------------------------------------------------
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date", "Attendee ID"])
    df = df.drop_duplicates(subset=["Attendee ID", "Date"])  # one record per day per attendee

    # Month-based fields
    df["Month"] = df["Date"].dt.to_period("M")
    df["JoinMonth"] = df.groupby("Attendee ID")["Date"].transform("min").dt.to_period("M")
    df["MonthOffset"] = (df["Month"] - df["JoinMonth"]).apply(lambda x: x.n)

    # ----------------------------------------------------
    # 2. Cohort Retention Table
    # ----------------------------------------------------
    cohort = (
        df.groupby(["JoinMonth", "MonthOffset"])["Attendee ID"]
        .nunique()
        .reset_index()
    )

    cohort_pivot = cohort.pivot(index="JoinMonth", columns="MonthOffset", values="Attendee ID")
    cohort_size = cohort_pivot[0]
    retention = cohort_pivot.divide(cohort_size, axis=0) * 100

    # ----------------------------------------------------
    # 3. Cohort Heatmap (Counts)
    # ----------------------------------------------------
    plt.figure(figsize=(12, 6))
    sns.heatmap(
        cohort_pivot,
        annot=True,
        fmt=".0f",
        cmap="YlGnBu",
        cbar_kws={'label': 'Active Participants'},
    )
    plt.title("Cohort Retention Heatmap (Active Participants by Month Since Joining)", fontsize=12, pad=15)
    plt.xlabel("Months Since Joining")
    plt.ylabel("Cohort (Join Month)")
    plt.tight_layout()
    plt.show(block=False)

    # ----------------------------------------------------
    # 4. Churn Rate Trend (Simplified View)
    # ----------------------------------------------------
    overall = (
        df.groupby("Month")["Attendee ID"]
        .nunique()
        .reset_index(name="ActiveParticipants")
        .sort_values("Month")
    )

    # Calculate churn rate (month-over-month)
    overall["PrevMonth"] = overall["ActiveParticipants"].shift(1)
    overall["ChurnRate"] = (1 - (overall["ActiveParticipants"] / overall["PrevMonth"])) * 100
    overall["ChurnRate"] = overall["ChurnRate"].fillna(0).clip(lower=0)

    overall["MonthLabel"] = overall["Month"].dt.strftime("%b %Y")

    # Plot churn only
    plt.figure(figsize=(10, 5))
    plt.plot(
        overall["MonthLabel"],
        overall["ChurnRate"],
        color="red",
        marker="s",
        linestyle="--",
        linewidth=2,
        label="Churn Rate (%)",
    )
    plt.title("Monthly Churn Rate Trend", fontsize=12)
    plt.xlabel("Month")
    plt.ylabel("Churn Rate (%)")
    plt.legend(loc="upper right")
    plt.grid(True, linestyle="--", alpha=0.6)
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.show(block=False)

    # ----------------------------------------------------
    # 5. Rolling Retention of New Joiners
    # ----------------------------------------------------
    month_groups = df.groupby("Month")["Attendee ID"].apply(set).sort_index()
    months = list(month_groups.index)
    results = []

    for i in range(1, len(months)):
        prev_all = month_groups[months[i - 1]]
        earlier = set().union(*month_groups.iloc[: i - 1]) if i > 1 else set()
        prev_joiners_only = prev_all - earlier

        curr_set = month_groups[months[i]]
        retained = prev_joiners_only & curr_set
        prev_total = len(prev_joiners_only)

        retention_rate = (len(retained) / prev_total) * 100 if prev_total > 0 else 0.0
        dropout_rate = 100 - retention_rate

        results.append({
            "Month": str(months[i]),
            "Retention%": round(retention_rate, 2),
            "Dropout%": round(dropout_rate, 2),
            "PrevMonthNewJoiners": prev_total,
            "RetainedFromPrevJoiners": len(retained),
        })

    trend = pd.DataFrame(results)

    plt.figure(figsize=(10, 5))
    plt.plot(trend["Month"], trend["Retention%"], color="green", marker="o", label="Retention %")
    plt.plot(trend["Month"], trend["Dropout%"], color="red", marker="o", label="Dropout %")
    plt.title("Retention of Previous Month's New Joiners (Month-to-Month Continuity)")
    plt.xlabel("Month")
    plt.ylabel("Percentage Retained from Previous Month’s Joiners")
    plt.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.ioff()  # turn off interactive mode
    plt.show()  # wait until user closes all plots

    print("\nRolling Retention Summary:")
    print(trend.to_string(index=False))


# --------------------------------------------------------
# Allow standalone run or main.py trigger
# --------------------------------------------------------
if __name__ == "__main__":
    run_retention_trend()
