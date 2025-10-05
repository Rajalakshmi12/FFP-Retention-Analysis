import sys
import os
import pandas as pd
from PyQt5.QtCore import QSize, Qt, QProcess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
)
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt


def make_excel_icon(size=32):
    """Draw an Excel-like icon (two-tone green + white X + grid lines)."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)

    # Left half - dark green, right half - lighter green
    painter.fillRect(0, 0, size // 2, size, QColor("#107C41"))
    painter.fillRect(size // 2, 0, size // 2, size, QColor("#185C37"))

    # Grid lines
    painter.setPen(QColor(200, 255, 200))
    step = size // 4
    for i in range(1, 4):
        painter.drawLine(i * step, 0, i * step, size)
        painter.drawLine(0, i * step, size, i * step)

    # White bold X
    painter.setPen(Qt.white)
    font = QFont("Arial", int(size * 0.6), QFont.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "X")

    painter.end()
    return QIcon(pixmap)


class RetentionChart(FigureCanvas):
    def __init__(self, df):
        self.fig, self.ax = plt.subplots(figsize=(10, 5))
        super().__init__(self.fig)
        self.plot_chart(df)

    def plot_chart(self, df):
        self.ax.clear()
        self.ax.set_title("All Activity Types - Participants by Gender")
        data = df.groupby(["Activity type", "Gender"]).size().unstack(fill_value=0)
        activity_types = list(data.index)
        x = range(len(activity_types))
        bottom = [0] * len(activity_types)
        for gender in data.columns:
            counts = data[gender].tolist()
            self.ax.bar(x, counts, label=gender, bottom=bottom)
            bottom = [sum(x) for x in zip(bottom, counts)]
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(activity_types, rotation=15)
        self.ax.set_ylabel("Number of Participants")
        self.ax.legend()
        self.draw()


class RetentionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_sessions = 4
        self.session_buttons = []
        self.geo_process = None  # to track external script

        self.setWindowTitle("FFP Retention & Engagement Dashboard")
        self.layout = QVBoxLayout(self)

        # Load data
        df = pd.read_excel("Mar24_Mar25.xlsx", sheet_name="UI-Sheet")
        df.rename(columns=lambda x: x.strip(), inplace=True)
        df["Date"] = pd.to_datetime(df["Date"])
        df["Week"] = df["Date"].dt.isocalendar().week
        self.df = df

        # --- Compact button row ---
        button_row_widget = QWidget()
        button_row_layout = QHBoxLayout(button_row_widget)
        button_row_layout.setContentsMargins(0, 0, 0, 0)
        button_row_layout.setSpacing(5)

        excel_icon = make_excel_icon(28)

        for i in range(1, 11):
            vbox = QVBoxLayout()
            vbox.setSpacing(2)

            # Number button
            num_btn = QPushButton(str(i))
            num_btn.setFixedSize(QSize(70, 35))
            num_btn.clicked.connect(lambda _, count=i, b=num_btn: self.set_and_run(count, b))
            self.session_buttons.append(num_btn)
            vbox.addWidget(num_btn, alignment=Qt.AlignHCenter)

            # Excel button
            excel_btn = QPushButton()
            excel_btn.setIcon(excel_icon)
            excel_btn.setIconSize(QSize(20, 20))
            excel_btn.setFixedSize(QSize(26, 26))
            excel_btn.setStyleSheet("background: transparent; border: none;")
            excel_btn.clicked.connect(lambda _, count=i: self.generate_excel(count))
            vbox.addWidget(excel_btn, alignment=Qt.AlignHCenter)

            button_row_layout.addLayout(vbox)

        self.layout.addWidget(QLabel("Select session count for retention calculation (monthly):"))
        self.layout.addWidget(button_row_widget)

        # Result label
        self.result_label = QLabel("")
        self.result_label.setStyleSheet("font-size: 16pt; font-weight: bold; color: #003366;")
        self.result_label.setWordWrap(True)
        self.result_label.setMinimumHeight(60)
        self.layout.addWidget(self.result_label)

        # Weekly activity buttons
        for freq in ["Once", "Twice", "Five+"]:
            btn = QPushButton(f"{freq} per Week - Last Quarter")
            btn.clicked.connect(lambda _, cat=freq: self.weekly_activity(cat))
            self.layout.addWidget(btn)

        # Chart
        chart = RetentionChart(self.df)
        self.layout.addWidget(chart)

        # --- New bottom button to run external script ---
        self.run_geo_btn = QPushButton("Run Ward Geospatial Mapping")
        self.run_geo_btn.setStyleSheet("""
            QPushButton {
                background-color: #185C37; 
                color: white; 
                font-weight: bold; 
                padding: 8px;
            }
            QPushButton:disabled {
                background-color: #555555;   /* dark grey */
                color: #dddddd;
            }
        """)
        self.run_geo_btn.clicked.connect(self.run_geospatial)
        self.layout.addWidget(self.run_geo_btn, alignment=Qt.AlignCenter)

        # Default = 4
        self.set_and_run(4, self.session_buttons[3])

    def set_and_run(self, val, btn):
        for b in self.session_buttons:
            b.setStyleSheet("")
        btn.setStyleSheet("background-color: green; color: white;")
        self.selected_sessions = val
        self.calc_retention(val)

    def calc_retention(self, dynamic_sessions):
        latest_date = self.df["Date"].max()
        cutoff = latest_date - pd.DateOffset(months=2)

        eligible = self.df.groupby("Attendee ID")["Date"].min()
        eligible_ids = eligible[eligible < cutoff].index
        df_eligible = self.df[self.df["Attendee ID"].isin(eligible_ids)].copy()
        df_eligible["Month"] = df_eligible["Date"].dt.to_period("M")

        month_counts = df_eligible.groupby(["Attendee ID", "Month"]).size().reset_index(name="count")
        latest = month_counts["Month"].max()
        prev = latest - 1

        active_last_two = set(month_counts[month_counts["Month"].isin([prev, latest])]["Attendee ID"])

        retained = set(
            month_counts[(month_counts["Month"] == latest) & (month_counts["count"] >= dynamic_sessions)]["Attendee ID"]
        ) & set(
            month_counts[(month_counts["Month"] == prev) & (month_counts["count"] >= dynamic_sessions)]["Attendee ID"]
        )

        pct = round(len(retained) / len(active_last_two) * 100, 2) if len(active_last_two) > 0 else 0
        self.result_label.setText(
            f"{len(retained)} out of {len(active_last_two)} participants in the last two months "
            f"had ≥{dynamic_sessions} sessions ({pct}%)."
        )

        self.last_retained = retained
        self.last_df_eligible = df_eligible

    def generate_excel(self, dynamic_sessions):
        self.calc_retention(dynamic_sessions)

        retained = getattr(self, "last_retained", set())
        df_eligible = getattr(self, "last_df_eligible", self.df)

        folder = "Retention Excels"
        os.makedirs(folder, exist_ok=True)
        export_file = os.path.join(folder, f"retention_report_{dynamic_sessions}.xlsx")

        if not os.path.exists(export_file):
            df_retained = df_eligible[df_eligible["Attendee ID"].isin(retained)]
            grouped = df_retained.groupby("Attendee ID")["Activity type"].apply(
                lambda x: ", ".join(sorted(set(x)))
            ).reset_index()
            grouped.to_excel(export_file, index=False)

        if sys.platform == "win32":
            os.startfile(export_file)
        elif sys.platform == "darwin":
            os.system(f"open '{export_file}'")
        else:
            os.system(f"xdg-open '{export_file}'")

    def run_geospatial(self):
        """Run external ward-geospatial-mapping.py file with button disable/enable."""
        self.run_geo_btn.setEnabled(False)
        self.run_geo_btn.setText("Opening Ward Geospatial Mapping...")

        self.geo_process = QProcess(self)
        self.geo_process.finished.connect(self.enable_geo_button)

        # Run the external Python file
        self.geo_process.start(sys.executable, ["ward-geospatial-mapping.py"])

    def enable_geo_button(self):
        self.run_geo_btn.setEnabled(True)
        self.run_geo_btn.setText("Run Ward Geospatial Mapping")

    def weekly_activity(self, category):
        latest_date = self.df["Date"].max()
        cutoff = latest_date - pd.DateOffset(months=3)
        eligible = self.df.groupby("Attendee ID")["Date"].min()
        eligible_ids = eligible[eligible < cutoff].index
        df_eligible = self.df[self.df["Attendee ID"].isin(eligible_ids)].copy()

        df_eligible["Week"] = df_eligible["Date"].dt.isocalendar().week
        weekly_counts = df_eligible.groupby(["Attendee ID", "Week"]).size().reset_index(name="Weekly Count")
        avg_week = weekly_counts.groupby("Attendee ID")["Weekly Count"].mean().reset_index()

        label_map = {"Once": (0, 1.5), "Twice": (1.5, 2.5), "Five+": (2.5, 7)}
        lo, hi = label_map[category]

        selected_ids = avg_week[(avg_week["Weekly Count"] > lo) & (avg_week["Weekly Count"] <= hi)]["Attendee ID"]
        df_selected = df_eligible[df_eligible["Attendee ID"].isin(selected_ids)]

        top_activities = df_selected["Activity type"].value_counts().nlargest(2).index.tolist()
        act_text = " and ".join(top_activities) if top_activities else "various activities"
        count = df_selected["Attendee ID"].nunique()

        label = QLabel(f"{count} participants were active {category.lower()} per week. Mostly engaged in {act_text}.")
        self.layout.addWidget(label)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app_window = RetentionApp()
    app_window.showMaximized()
    sys.exit(app.exec_())
