import pandas as pd
import tkinter as tk
from tkinter import ttk
from datetime import timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
from matplotlib.gridspec import GridSpec

# -----------------------------
# 1. Load Excel
# -----------------------------
df = pd.read_excel("Documents/Mar24_Mar25_Cleansed.xlsx", sheet_name="Main")
required = {"Attendee ID", "Activity ID", "Date"}
if not required.issubset(df.columns):
    raise ValueError(f"Excel must contain: {required}")

df["Date"] = pd.to_datetime(df["Date"])

# -----------------------------
# 2. Filter last 2 months
# -----------------------------
max_date = df["Date"].max()
cutoff = max_date - timedelta(days=60)
df = df[df["Date"] >= cutoff]

# -----------------------------
# 3. Deduplicate & weekly grouping (Mon–Fri)
# -----------------------------
df = df.drop_duplicates(subset=["Attendee ID", "Date"])
df = df[df["Date"].dt.dayofweek < 5]
df["Week"] = df["Date"] - pd.to_timedelta(df["Date"].dt.dayofweek, unit="d")

weekly = (
    df.groupby(["Attendee ID", "Week"])["Date"]
    .nunique()
    .reset_index(name="Days_Attended")
)
weekly["Days_Attended"] = weekly["Days_Attended"].clip(upper=5)

# -----------------------------
# 4. Sort participants
# -----------------------------
totals = weekly.groupby("Attendee ID")["Days_Attended"].sum().reset_index(name="Total")
sorted_ids = totals.sort_values(by="Total", ascending=False)["Attendee ID"].tolist()
total_count = len(sorted_ids)

# Pivot table for heatmap
pivot_df = (
    weekly.pivot_table(index="Attendee ID", columns="Week", values="Days_Attended", fill_value=0)
    .reindex(sorted_ids)
)
pivot_df = pivot_df.reindex(sorted(pivot_df.columns), axis=1)

# -----------------------------
# 5. Tkinter UI
# -----------------------------
root = tk.Tk()
root.title("Dynamic Weekly Attendance Heatmap")
root.geometry("1250x780")

frame = ttk.Frame(root)
frame.pack(fill="both", expand=True, padx=10, pady=10)

# Title
ttk.Label(frame, text="Weekly Attendance Heatmap (Last 2 Months, Mon–Fri)", font=("Arial", 15)).pack(pady=10)

# Control Frame (for group options + dropdown)
control_frame = ttk.Frame(frame)
control_frame.pack(pady=5)

# Variable for number of groups
num_groups_var = tk.IntVar(value=20)

# Variable for range dropdown
group_var = tk.StringVar()

# Create frame for group buttons
ttk.Label(control_frame, text="Divide participants into:", font=("Arial", 11)).grid(row=0, column=0, padx=5)

def update_groups(*args):
    """Recalculate groups and update dropdown dynamically."""
    num_groups = num_groups_var.get()
    group_size = max(total_count // num_groups, 1)

    groups = [(i * group_size + 1, min((i + 1) * group_size, total_count)) for i in range(num_groups)]
    group_labels = [f"{a}-{b}" for a, b in groups]

    dropdown["values"] = group_labels
    group_var.set(group_labels[0])
    plot_group()  # redraw with new grouping

# Radio buttons for group selection
group_options = [5, 10, 20, 30, 40, 50]
for i, val in enumerate(group_options):
    ttk.Radiobutton(control_frame, text=f"{val} groups", variable=num_groups_var, value=val, command=update_groups).grid(row=0, column=i + 1, padx=4)

# Dropdown for participant range
ttk.Label(control_frame, text="Select range:", font=("Arial", 11)).grid(row=0, column=len(group_options) + 2, padx=(15, 5))
dropdown = ttk.Combobox(control_frame, textvariable=group_var, state="readonly", width=18, font=("Arial", 11))
dropdown.grid(row=0, column=len(group_options) + 3, padx=5)

# -----------------------------
# 6. Matplotlib Figure + Colorbar (GridSpec)
# -----------------------------
fig = plt.Figure(figsize=(12, 6), dpi=100)
gs = GridSpec(1, 2, width_ratios=[20, 1], wspace=0.05, figure=fig)
ax = fig.add_subplot(gs[0, 0])
cbar_ax = fig.add_subplot(gs[0, 1])

canvas = FigureCanvasTkAgg(fig, master=frame)
canvas.get_tk_widget().pack(fill="both", expand=True, pady=10)

# -----------------------------
# 7. Plot function
# -----------------------------
def plot_group(event=None):
    ax.clear()
    selection = group_var.get()
    if not selection:
        return

    start, end = map(int, selection.split("-"))
    subset_ids = sorted_ids[start - 1:end]
    sub = pivot_df.loc[subset_ids]

    sns.heatmap(
        sub,
        cmap="YlOrRd",
        linewidths=0.3,
        linecolor="gray",
        vmin=0,
        vmax=5,
        cbar_ax=cbar_ax,
        cbar=True,
        ax=ax
    )

    ax.set_title(f"Participants {start}–{end} (of {total_count})", fontsize=13)
    ax.set_xlabel("Week Starting (Monday)")
    ax.set_ylabel("Attendee ID")

    # Clean formatted X-axis
    xtick_labels = [
        f"Week {d.isocalendar().week} – {d.strftime('%b')} – {d.year}"
        for d in sub.columns
    ]
    ax.set_xticklabels(xtick_labels, rotation=45, ha="right")

    fig.subplots_adjust(bottom=0.25, right=0.93, top=0.9, wspace=0.05)
    canvas.draw()

# -----------------------------
# 8. Initialize Default Group & Plot
# -----------------------------
update_groups()  # creates dropdown and first plot
dropdown.bind("<<ComboboxSelected>>", plot_group)

root.mainloop()
