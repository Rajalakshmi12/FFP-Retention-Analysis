import tkinter as tk
from tkinter import ttk
import subprocess
import sys
import os

# Function to run a child Python file
def run_script(script_name):
    # Get absolute path
    script_path = os.path.join(os.path.dirname(__file__), script_name)

    if os.path.exists(script_path):
        root.destroy()  # close main launcher before opening next window
        subprocess.run([sys.executable, script_path])
    else:
        tk.messagebox.showerror("Error", f"Script not found:\n{script_name}")

# -----------------------------
# MAIN DASHBOARD WINDOW
# -----------------------------
root = tk.Tk()
root.title("Retention & Heatmap Dashboard")
root.geometry("500x300")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(expand=True, fill="both")

title = ttk.Label(frame, text="Select an Option", font=("Arial", 18))
title.pack(pady=30)

# Buttons for both apps
btn1 = ttk.Button(
    frame,
    text="🧮 Retention Percentage Calculation",
    command=lambda: run_script("pyqt5UI-chart-with-map.py"),
    width=40
)
btn1.pack(pady=15)

btn2 = ttk.Button(
    frame,
    text="🌡️ Heat Map Generation",
    command=lambda: run_script("Engagement_Analysis.py"),
    width=40
)
btn2.pack(pady=15)

# Optional: Exit button
ttk.Button(frame, text="❌ Exit", command=root.destroy).pack(pady=20)

root.mainloop()
