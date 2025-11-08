# trend_summary.py
# -------------------------------------------------------
# Compares participation between first 30 days and last 30 days
# of the latest 60-day window in the dataset
# + Includes IMD & Ward-based dropout summary
# + Adds Weekday Attendance Weightage Summary
# + Adds corrected 3 & 6 month dropout duration summary
# + Adds Dropout Characteristics (Gender, Age, Constituency)
# -------------------------------------------------------

import pandas as pd
import os
from datetime import timedelta


def generate_trend_summary():
    FILE_PATH = "Documents/Mar24_Mar25_Cleansed.xlsx"
    SHEET = "Main"

    if not os.path.exists(FILE_PATH):
        return f"❌ File not found:\n{FILE_PATH}"

    # ---------------------------
    # Load and clean
    # ---------------------------
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])

    for col in ["Gender", "Activity type", "RajiNewColumn-Range", "Constituency", "Ward"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df = df[~df[col].isin(["", "nan", "None", "Not provided", "Not Provided"])]

    latest_date = df["Date"].max().normalize()
    start_date = latest_date - timedelta(days=60)
    mid_date = start_date + timedelta(days=30)

    df_window = df[(df["Date"] >= start_date) & (df["Date"] <= latest_date)]
    df_first30 = df_window[df_window["Date"] < mid_date]
    df_last30 = df_window[df_window["Date"] >= mid_date]

    summary_text = (
        f"📆 Participation Comparison (Last 60 Days Window)\n"
        f"Period: {start_date.date()} → {latest_date.date()}\n"
        + "-" * 65 + "\n\n"
    )

    # ---------------------------
    # Helper Function
    # ---------------------------
    def compare_categories(df_first, df_last, col, label,
                           show_top=False, show_all=False,
                           top_n=2, wording="Participation in {}"):
        prev_counts = df_first.groupby(col)["Attendee ID"].nunique()
        last_counts = df_last.groupby(col)["Attendee ID"].nunique()

        cats = sorted(set(prev_counts.index) | set(last_counts.index))
        changes = []
        for c in cats:
            prev_val = prev_counts.get(c, 0)
            last_val = last_counts.get(c, 0)
            if prev_val == 0 and last_val == 0:
                continue
            pct_change = (
                ((last_val - prev_val) / prev_val * 100)
                if prev_val > 0
                else (100.0 if last_val > 0 else 0.0)
            )
            abs_change = abs(last_val - prev_val)
            changes.append((c, prev_val, last_val, abs_change, pct_change))

        if not changes:
            return f"No valid data for {label}.\n"

        lines = [f"{label} Highlights (First 30 days → Last 30 days):"]

        if show_top == "absolute":
            top_shift = sorted(changes, key=lambda x: x[3], reverse=True)[:top_n]
            for c, prev_val, last_val, abs_change, pct_change in top_shift:
                direction = "increased" if last_val > prev_val else "decreased"
                lines.append(
                    f"• {c} participation {direction} by {abs_change} "
                    f"({abs(pct_change):.1f}% change, {prev_val} → {last_val})"
                )

        elif show_all:
            for c, prev_val, last_val, abs_change, pct_change in sorted(changes, key=lambda x: x[0]):
                if abs(pct_change) < 1:
                    lines.append(f"• {wording.format(c)} remained steady ({prev_val} → {last_val})")
                elif pct_change > 0:
                    lines.append(f"• {wording.format(c)} increased by {abs(pct_change):.1f}% ({prev_val} → {last_val})")
                else:
                    lines.append(f"• {wording.format(c)} decreased by {abs(pct_change):.1f}% ({prev_val} → {last_val})")

        else:
            for c, prev_val, last_val, abs_change, pct_change in changes:
                if abs(pct_change) < 1:
                    status = "remained stable"
                elif pct_change > 0:
                    status = f"increased by {abs(pct_change):.1f}%"
                else:
                    status = f"decreased by {abs(pct_change):.1f}%"
                lines.append(f"• {c} participation {status} ({prev_val} → {last_val})")

        return "\n".join(lines) + "\n"

    # ---------------------------
    # Gender Trend
    # ---------------------------
    if "Gender" in df.columns:
        summary_text += compare_categories(
            df_first30, df_last30, "Gender", "Gender"
        ) + "\n"

    # ---------------------------
    # Activity Trend
    # ---------------------------
    if "Activity type" in df.columns:
        summary_text += compare_categories(
            df_first30, df_last30, "Activity type", "Activity",
            show_all=True, wording="Participation in {} Activity"
        ) + "\n"

    # ---------------------------
    # Age Bucket Trend
    # ---------------------------
    if "RajiNewColumn-Range" in df.columns:
        summary_text += compare_categories(
            df_first30, df_last30, "RajiNewColumn-Range", "Age Bucket",
            show_all=True, wording="Participation among {} age group"
        ) + "\n"

    # ---------------------------
    # Constituency Trend (Top 5)
    # ---------------------------
    if "Constituency" in df.columns:
        summary_text += compare_categories(
            df_first30, df_last30, "Constituency", "Constituency",
            show_top="absolute", top_n=5
        ) + "\n"

    # ---------------------------
    # IMD Dropout Trend + Characteristics
    # ---------------------------
    if {"IMD rank", "Ward", "Attendee ID"}.issubset(df.columns):
        first_ids = set(df_first30["Attendee ID"].unique())
        last_ids = set(df_last30["Attendee ID"].unique())
        dropout_ids = first_ids - last_ids
        dropout_df = df_first30[df_first30["Attendee ID"].isin(dropout_ids)]

        if not dropout_df.empty:
            imd_ward_counts = (
                dropout_df.groupby(["IMD rank", "Ward"])["Attendee ID"]
                .nunique()
                .reset_index(name="Dropouts")
                .sort_values(by="Dropouts", ascending=False)
                .head(3)
            )

            summary_text += "IMD Dropout Highlights (Top 3 by dropout count):\n"
            for _, row in imd_ward_counts.iterrows():
                summary_text += (
                    f"• Ward: {row['Ward']} | IMD rank: {row['IMD rank']} "
                    f"→ {int(row['Dropouts'])} dropouts\n"
                )

            # --- Dropout characteristics ---
            gender_top = dropout_df["Gender"].mode()[0] if "Gender" in dropout_df.columns and not dropout_df["Gender"].dropna().empty else "N/A"
            age_top = dropout_df["RajiNewColumn-Range"].mode()[0] if "RajiNewColumn-Range" in dropout_df.columns and not dropout_df["RajiNewColumn-Range"].dropna().empty else "N/A"
            const_top = dropout_df["Constituency"].mode()[0] if "Constituency" in dropout_df.columns and not dropout_df["Constituency"].dropna().empty else "N/A"

            summary_text += (
                "\n📊 Dropout Characteristics Summary:\n"
                f"• Most common gender among dropouts: {gender_top}\n"
                f"• Most common age range among dropouts: {age_top}\n"
                f"• Top constituency with highest dropouts: {const_top}\n\n"
            )
        else:
            summary_text += "No IMD dropout data found in the current window.\n\n"

    # ---------------------------
    # Weekday Attendance Weightage
    # ---------------------------
    if {"Date", "Attendee ID"}.issubset(df.columns):
        df["Weekday"] = df["Date"].dt.day_name()
        unique_attendance = df.groupby(["Date", "Attendee ID"]).size().reset_index(name="Sessions")
        unique_attendance = unique_attendance.drop_duplicates(subset=["Date", "Attendee ID"])
        weekday_counts = unique_attendance["Date"].dt.day_name().value_counts(normalize=True) * 100
        weekday_counts = weekday_counts.reindex(
            ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        ).fillna(0)

        summary_text += "📅 Attendance Weightage by Weekday:\n"
        for day, pct in weekday_counts.items():
            summary_text += f"• {day}: {pct:.1f}% of total unique attendances\n"

    # ---------------------------
    # Drop-off Duration Summary (corrected 3 & 6 months)
    # ---------------------------
    if {"Date", "Attendee ID"}.issubset(df.columns):
        attendance = df.groupby("Attendee ID")["Date"].agg(["min", "max"]).reset_index()
        attendance.rename(columns={"min": "FirstSession", "max": "LastSession"}, inplace=True)
        attendance["MonthsActive"] = (attendance["LastSession"] - attendance["FirstSession"]) / pd.Timedelta(days=30)

        latest_date = df["Date"].max()
        cutoff_3 = latest_date - pd.DateOffset(months=3)
        cutoff_6 = latest_date - pd.DateOffset(months=6)

        valid_3 = attendance[attendance["FirstSession"] <= cutoff_3]
        valid_6 = attendance[attendance["FirstSession"] <= cutoff_6]

        total3 = len(valid_3)
        total6 = len(valid_6)
        drop3 = len(valid_3[valid_3["MonthsActive"] <= 3])
        drop6 = len(valid_6[valid_6["MonthsActive"] <= 6])

        pct3 = round(drop3 / total3 * 100, 2) if total3 > 0 else 0
        pct6 = round(drop6 / total6 * 100, 2) if total6 > 0 else 0

        summary_text += (
            "\n⏳ Drop-off Duration Summary (adjusted for recent joiners):\n"
            f"• {pct3}% of eligible members did not attend any session after 3 months from their joining date.\n"
            f"• {pct6}% of eligible members did not attend any session after 6 months from their joining date.\n"
        )

    return summary_text.strip()


# Standalone run
if __name__ == "__main__":
    print(generate_trend_summary())
