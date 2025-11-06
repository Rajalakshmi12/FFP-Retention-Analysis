# trend_summary.py
# -------------------------------------------------------
# Generate a clean trend summary with dropout characteristics
# for Fight For Peace dataset (Gender, Programme, Age, Location)
# -------------------------------------------------------

import pandas as pd
import os


def generate_trend_summary():
    FILE_PATH = "Documents/Mar24_Mar25_Cleansed.xlsx"
    SHEET = "Main"

    if not os.path.exists(FILE_PATH):
        return f"❌ Excel file not found:\n{FILE_PATH}"

    # ---------------------------
    # Load and prepare data
    # ---------------------------
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Month"] = df["Date"].dt.to_period("M")

    # Clean common categorical columns
    for col in ["Gender", "Activity type", "Constituency"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df = df[~df[col].isin(["", "nan", "None", "Not provided", "Not Provided"])]

    # ---------------------------
    # Define earliest and latest months
    # ---------------------------
    earliest_month = df["Month"].min()
    latest_month = df["Month"].max()
    df_range = df[(df["Month"] >= earliest_month) & (df["Month"] <= latest_month)].copy()

    # ---------------------------
    # Helper functions
    # ---------------------------
    def calculate_change(latest, earliest):
        if earliest > 0:
            return ((latest - earliest) / earliest) * 100
        elif latest > 0 and earliest == 0:
            return 100.0
        elif latest == 0 and earliest > 0:
            return -100.0
        else:
            return 0.0

    def make_summary_line(cat, change, label):
        if abs(change) < 1:
            return f"{cat} {label} remained stable."
        elif change > 0:
            return f"{cat} {label} increased by {abs(change):.1f}%."
        else:
            return f"{cat} {label} decreased by {abs(change):.1f}%."

    summary_text = "📈 Trend Summary (Full Data Period)\n" + "-" * 50 + "\n\n"

    # ---------------------------
    # Gender Trend (clear wording)
    # ---------------------------
    gender_trend = df_range.groupby(["Month", "Gender"]).size().unstack(fill_value=0)
    if len(gender_trend) >= 2:
        gender_change = {
            cat: calculate_change(gender_trend.iloc[-1][cat], gender_trend.iloc[0][cat])
            for cat in gender_trend.columns
        }
        male_change = gender_change.get("Male", 0)
        female_change = gender_change.get("Female", 0)

        male_text = (
            f"Male participation increased by {male_change:.1f}%"
            if male_change > 0
            else f"Male participation decreased by {abs(male_change):.1f}%"
        )
        female_text = (
            f"Female participation increased by {female_change:.1f}%"
            if female_change > 0
            else f"Female participation decreased by {abs(female_change):.1f}%"
        )

        summary_text += f"{female_text}, while {male_text}.\n\n"

    # ---------------------------
    # Programme Trend (Top 2 ↑ & ↓)
    # ---------------------------
    prog_trend = df_range.groupby(["Month", "Activity type"]).size().unstack(fill_value=0)
    if len(prog_trend) >= 2:
        prog_change = {
            cat: calculate_change(prog_trend.iloc[-1][cat], prog_trend.iloc[0][cat])
            for cat in prog_trend.columns
        }
        sorted_prog = sorted(prog_change.items(), key=lambda x: x[1], reverse=True)
        top2_up = sorted_prog[:2]
        top2_down = sorted(sorted_prog, key=lambda x: x[1])[:2]

        summary_text += "Programme Trends:\n"
        for cat, val in top2_up:
            summary_text += f"• {make_summary_line(cat, val, 'sessions')}\n"
        for cat, val in top2_down:
            summary_text += f"• {make_summary_line(cat, val, 'sessions')}\n"
        summary_text += "\n"

    # ---------------------------
    # Age Bucket Trends (Retention-based, not % change)
    # ---------------------------
    if "RajiNewColumn-Age" in df_range.columns:
        df_age = df_range.dropna(subset=["RajiNewColumn-Age"]).copy()
        df_age["RajiNewColumn-Age"] = pd.to_numeric(df_age["RajiNewColumn-Age"], errors="coerce")
        df_age = df_age.dropna(subset=["RajiNewColumn-Age"])

        if not df_age.empty:
            # Assign age bucket
            df_age["Age Bucket"] = pd.cut(
                df_age["RajiNewColumn-Age"],
                bins=[0, 12, 17, 22, 30, 40, 100],
                labels=["0–12", "13–17", "18–22", "23–30", "31–40", "40+"],
            )

            # One row per attendee with age bucket + first/last month
            att_age = (
                df_age.sort_values("Date")
                .groupby("Attendee ID")
                .agg({
                    "Age Bucket": "first",
                    "Month": ["min", "max"]
                })
            )
            att_age.columns = ["AgeBucket", "FirstMonth", "LastMonth"]
            att_age = att_age.dropna(subset=["AgeBucket"])

            # Retention by age bucket: still active in final month
            latest = latest_month
            age_lines = []
            for bucket in att_age["AgeBucket"].cat.categories:
                subset = att_age[att_age["AgeBucket"] == bucket]
                total = len(subset)
                if total == 0:
                    continue
                retained = (subset["LastMonth"] == latest).sum()
                retention_rate = (retained / total) * 100
                age_lines.append(
                    f"• For age group {bucket}, about {retention_rate:.1f}% of participants who ever attended were still active in the final month."
                )

            if age_lines:
                summary_text += "Age Group Trends (retention-based):\n"
                summary_text += "\n".join(age_lines) + "\n\n"

    # ---------------------------
    # Location Trend (Top 2 ↑ & ↓)
    # ---------------------------
    loc_trend = df_range.groupby(["Month", "Constituency"]).size().unstack(fill_value=0)
    if len(loc_trend) >= 2:
        loc_change = {
            cat: calculate_change(loc_trend.iloc[-1][cat], loc_trend.iloc[0][cat])
            for cat in loc_trend.columns
        }
        sorted_loc = sorted(loc_change.items(), key=lambda x: x[1], reverse=True)
        top2_up = sorted_loc[:2]
        top2_down = sorted(sorted_loc, key=lambda x: x[1])[:2]

        summary_text += "Location Highlights:\n"
        for cat, val in top2_up:
            summary_text += f"• {make_summary_line(cat, val, 'participation')}\n"
        for cat, val in top2_down:
            summary_text += f"• {make_summary_line(cat, val, 'participation')}\n"
        summary_text += "\n"

    # ---------------------------
    # Dropout Summary (based on “stopped before dataset end”)
    # ---------------------------
    summary_text += "Dropout Characteristics:\n"

    join_month = df.groupby("Attendee ID")["Month"].min()
    last_active = df.groupby("Attendee ID")["Month"].max()

    # Dropouts: last activity before final month AND joined earlier than final month - 1
    dropout_ids = last_active[
        (last_active < latest_month) & (join_month < (latest_month - 1))
    ].index

    df_dropouts = df[df["Attendee ID"].isin(dropout_ids)].copy()

    for col in ["Gender", "Constituency", "RajiNewColumn-Age"]:
        if col in df_dropouts.columns:
            df_dropouts[col] = df_dropouts[col].astype(str).str.strip()
            df_dropouts = df_dropouts[
                ~df_dropouts[col].isin(["", "nan", "None", "Not provided", "Not Provided"])
            ]

    top_gender, top_location, top_age = "N/A", "N/A", "N/A"

    if not df_dropouts.empty:
        if "Gender" in df_dropouts.columns and not df_dropouts["Gender"].empty:
            top_gender = df_dropouts["Gender"].mode(dropna=True).iloc[0]
        if "Constituency" in df_dropouts.columns and not df_dropouts["Constituency"].empty:
            top_location = df_dropouts["Constituency"].mode(dropna=True).iloc[0]
        if "RajiNewColumn-Age" in df_dropouts.columns:
            df_dropouts["RajiNewColumn-Age"] = pd.to_numeric(df_dropouts["RajiNewColumn-Age"], errors="coerce")
            df_dropouts = df_dropouts.dropna(subset=["RajiNewColumn-Age"])
            if not df_dropouts.empty:
                df_dropouts["Age Bucket"] = pd.cut(
                    df_dropouts["RajiNewColumn-Age"],
                    bins=[0, 12, 17, 22, 30, 40, 100],
                    labels=["0–12", "13–17", "18–22", "23–30", "31–40", "40+"],
                )
                if not df_dropouts["Age Bucket"].empty:
                    top_age = df_dropouts["Age Bucket"].mode(dropna=True).iloc[0]

    summary_text += f"• Most dropouts were aged {top_age}\n"
    summary_text += f"• Majority of dropouts were {top_gender}\n"
    summary_text += f"• Dropouts were most frequent from {top_location}\n"

    return summary_text.strip()


# Standalone run
if __name__ == "__main__":
    print(generate_trend_summary())
