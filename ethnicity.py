# descriptive_summary.py
# -------------------------------------------------------
# Generates key descriptive statistics for FFP dataset
# Treats multiple sessions per day as one attendance
# Adds % Male, % Female, Average Attendances per Month,
# and participation upper/lower limits
# -------------------------------------------------------

import pandas as pd
import numpy as np

FILE_PATH = "Documents/Mar24_Mar25_Cleansed.xlsx"
SHEET = "Main"

def generate_descriptive_summary():
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])

    # --- Prepare Age if missing ---
    if "RajiNewColumn-Age" not in df.columns and "DOB" in df.columns:
        df["DOB"] = pd.to_datetime(df["DOB"], errors="coerce")
        df["RajiNewColumn-Age"] = (pd.Timestamp("today") - df["DOB"]).dt.days // 365

    # Treat multiple sessions per day as one attendance
    df_unique = df.drop_duplicates(subset=["Attendee ID", "Date"]).copy()

    # --- Metrics ---
    total_participants = df_unique["Attendee ID"].nunique()
    avg_age = round(df_unique["RajiNewColumn-Age"].mean(), 1) if "RajiNewColumn-Age" in df_unique.columns else np.nan

    # Gender %
    if "Gender" in df_unique.columns:
        gender_clean = df_unique["Gender"].str.lower().str.strip()
        male_count = gender_clean.eq("male").sum()
        female_count = gender_clean.eq("female").sum()
        total_gender = male_count + female_count
        male_pct = round((male_count / total_gender) * 100, 1) if total_gender > 0 else np.nan
        female_pct = round((female_count / total_gender) * 100, 1) if total_gender > 0 else np.nan
    else:
        male_pct = female_pct = np.nan

    # Unique attendances per attendee
    attendances_per_attendee = df_unique.groupby("Attendee ID").size()
    mean_attendances = round(attendances_per_attendee.mean(), 1)
    max_attendances = int(attendances_per_attendee.max())
    min_attendances = int(attendances_per_attendee.min())

    # Months active per attendee
    attendance = df_unique.groupby("Attendee ID")["Date"].agg(["min", "max"])
    attendance["MonthsActive"] = (attendance["max"] - attendance["min"]) / pd.Timedelta(days=30)
    median_months_active = round(attendance["MonthsActive"].median(), 1)

    # Average attendances per month (normalized engagement)
    attendance["Attendances"] = attendances_per_attendee
    attendance["AttendancesPerMonth"] = attendance["Attendances"] / attendance["MonthsActive"].replace(0, np.nan)
    avg_attendances_per_month = round(attendance["AttendancesPerMonth"].mean(), 1)

    # Top Activity and Constituency
    top_activity = df["Activity type"].mode()[0] if "Activity type" in df.columns else "N/A"
    top_constituency = df["Constituency"].mode()[0] if "Constituency" in df.columns else "N/A"

    # Average IMD rank
    if "IMD rank" in df.columns:
        avg_imd = int(round(df["IMD rank"].mean(), 0))
    else:
        avg_imd = "N/A"

    # Dropout % (3 and 6 months)
    attendance.reset_index(inplace=True)
    latest_date = df["Date"].max()
    cutoff_3 = latest_date - pd.DateOffset(months=3)
    cutoff_6 = latest_date - pd.DateOffset(months=6)
    valid_3 = attendance[attendance["min"] <= cutoff_3]
    valid_6 = attendance[attendance["min"] <= cutoff_6]
    drop3 = len(valid_3[valid_3["MonthsActive"] <= 3])
    drop6 = len(valid_6[valid_6["MonthsActive"] <= 6])
    pct3 = round(drop3 / len(valid_3) * 100, 1) if len(valid_3) > 0 else 0
    pct6 = round(drop6 / len(valid_6) * 100, 1) if len(valid_6) > 0 else 0

    # --- Output summary ---
    summary = pd.DataFrame({
        "Metric": [
            "Total Participants",
            "Avg. Age",
            "% Male",
            "% Female",
            "Mean Attendances per Attendee (Total)",
            "Average Attendances per Month",
            "Median Months Active",
            "Upper Limit (Max Attendances)",
            "Lower Limit (Min Attendances)",
            "Top Activity",
            "Top Constituency",
            "Avg. IMD Rank",
            "Dropout 3-Month",
            "Dropout 6-Month"
        ],
        "Value": [
            f"{total_participants:,}",
            f"{avg_age} years",
            f"{male_pct}%",
            f"{female_pct}%",
            mean_attendances,
            avg_attendances_per_month,
            median_months_active,
            max_attendances,
            min_attendances,
            top_activity,
            top_constituency,
            f"{avg_imd} (high deprivation)" if isinstance(avg_imd, int) else "N/A",
            f"{pct3}%",
            f"{pct6}%"
        ]
    })

    print("\n📋 Descriptive Statistics Summary:\n")
    print(summary.to_string(index=False))
    return summary


if __name__ == "__main__":
    generate_descriptive_summary()
