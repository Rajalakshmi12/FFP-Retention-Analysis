# main.py
# -------------------------------------------------------
# FFP Dashboard Launcher (Retention / Heatmap / Summary)
# -------------------------------------------------------

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import sys
import os
sys.path.append(os.path.dirname(__file__))  # ensures current folder is visible
from trend_summary import generate_trend_summary   # 👈 import summary logic


# ----------------------------------------------------
# Run other scripts (Retention, Heatmap)
# ----------------------------------------------------
def run_script(script_name):
    script_path = os.path.join(os.path.dirname(__file__), script_name)
    if os.path.exists(script_path):
        #root.destroy()
        subprocess.run([sys.executable, script_path])
    else:
        messagebox.showerror("Error", f"Script not found:\n{script_name}")


# ----------------------------------------------------
# Generate Trend Summary in same UI
# ----------------------------------------------------
def show_summary():
    try:
        summary_text = generate_trend_summary()
        summary_box.config(state="normal")
        summary_box.delete(1.0, tk.END)
        summary_box.insert(tk.END, summary_text)
        summary_box.config(state="disabled")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# ----------------------------------------------------
# MAIN DASHBOARD UI
# ----------------------------------------------------
root = tk.Tk()
root.title("Retention & Heatmap Dashboard")
root.geometry("600x550")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(expand=True, fill="both")

title = ttk.Label(frame, text="FFP Analysis", font=("Arial", 18))
title.pack(pady=20)

btn1 = ttk.Button(frame, text="🧮  Retention Percentage & Geo-spatial mapping",
                  command=lambda: run_script("pyqt5UI-chart-with-map.py"), width=50)
btn1.pack(pady=10)

btn2 = ttk.Button(frame, text="🌡️ Participants Engagement Map",
                  command=lambda: run_script("Engagement_Analysis.py"), width=50)
btn2.pack(pady=10)

btn3 = ttk.Button(frame, text="📉  Dropout Trend Analysis",
                  command=lambda: subprocess.Popen([sys.executable, "retention-trend.py"]),
                  width=50)
btn3.pack(pady=10)

btn4 = ttk.Button(frame, text="🧾  Generate Summary",
                  command=show_summary, width=50)
btn4.pack(pady=10)

summary_box = scrolledtext.ScrolledText(frame, wrap=tk.WORD, width=70, height=12, font=("Consolas", 10))
summary_box.pack(pady=10)
summary_box.insert(tk.END, "Click '🧾 Generate Summary' to view the latest 12-month trends here...")
summary_box.config(state="disabled")

ttk.Button(frame, text="❌ Exit", command=root.destroy).pack(pady=10)

root.mainloop()
