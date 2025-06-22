import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt

# Get the

# Dummy function for chart display (replace with actual logic)
def show_chart(label):
    plt.figure()
    plt.title(f"{label} - Weekly Retention")
    plt.plot([1, 2, 3, 4], [4, 5, 6, 7])  # Replace with actual data
    plt.xlabel("Week")
    plt.ylabel("Users")
    plt.grid(True)
    plt.show()

# GUI setup
root = tk.Tk()
root.title("Retention Analysis")
root.geometry("300x250")

# Define 4 buttons
buttons = {
    "Once a Week": lambda: show_chart("Once"),
    "Twice a Week": lambda: show_chart("Twice"),
    "Thrice a Week": lambda: show_chart("Thrice"),
    "Seven Days a Week": lambda: show_chart("Seven")
}

for label, command in buttons.items():
    ttk.Button(root, text=label, command=command).pack(pady=10)

root.mainloop()
