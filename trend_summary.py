# trend_summary.py
# -------------------------------------------------------
# Compares participation between first 30 days and last 30 days
# of the latest 60-day window in the dataset
# + Includes IMD & Ward-based dropout summary
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
    # Programme Trend
    # ---------------------------
    if "Activity type" in df.columns:
        summary_text += compare_categories(
            df_first30, df_last30, "Activity type", "Programme",
            show_all=True, wording="Participation in {} programme"
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
    # IMD Dropout Trend (with Ward)
    # ---------------------------
    if {"IMD rank", "Ward", "Attendee ID"}.issubset(df.columns):
        first_ids = set(df_first30["Attendee ID"].unique())
        last_ids = set(df_last30["Attendee ID"].unique())
        dropout_ids = first_ids - last_ids

        dropout_df = df_first30[df_first30["Attendee ID"].isin(dropout_ids)]
        if not dropout_df.empty:
            # Group by both IMD rank and Ward
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
            summary_text += "\n"
        else:
            summary_text += "No IMD dropout data found in the current window.\n\n"

    return summary_text.strip()


# Standalone run
if __name__ == "__main__":
    print(generate_trend_summary())
