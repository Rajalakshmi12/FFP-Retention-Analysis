import sys
import os
import pandas as pd
from PyQt5.QtCore import QSize, Qt, QProcess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QScrollArea
)
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt


def make_excel_icon(size=32):
    """Draw an Excel-like icon (two-tone green + white X + grid lines)."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.fillRect(0, 0, size // 2, size, QColor("#107C41"))
    painter.fillRect(size // 2, 0, size // 2, size, QColor("#185C37"))
    painter.setPen(QColor(200, 255, 200))
    step = size // 4
    for i in range(1, 4):
        painter.drawLine(i * step, 0, i * step, size)
        painter.drawLine(0, i * step, size, i * step)
    painter.setPen(Qt.white)
    font = QFont("Arial", int(size * 0.6), QFont.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "X")
    painter.end()
    return QIcon(pixmap)


class RetentionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_sessions = 4
        self.session_buttons = []
        self.geo_process = None

        self.setWindowTitle("FFP Retention & Engagement Dashboard")

        # ---------- Scrollable main layout ----------
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        container = QWidget()
        self.layout = QVBoxLayout(container)
        scroll.setWidget(container)
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(scroll)

        # ---------- Load data ----------
        df = pd.read_excel("Documents/Mar24_Mar25_Cleansed.xlsx", sheet_name="Main")
        df.rename(columns=lambda x: x.strip(), inplace=True)
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Week"] = df["Date"].dt.isocalendar().week
        self.df = df

        # ---------- Buttons Row ----------
        button_row_widget = QWidget()
        button_row_layout = QHBoxLayout(button_row_widget)
        button_row_layout.setContentsMargins(0, 0, 0, 0)
        button_row_layout.setSpacing(5)
        excel_icon = make_excel_icon(28)

        for i in range(1, 11):
            vbox = QVBoxLayout()
            vbox.setSpacing(2)
            num_btn = QPushButton(str(i))
            num_btn.setFixedSize(QSize(70, 35))
            num_btn.clicked.connect(lambda _, count=i, b=num_btn: self.set_and_run(count, b))
            self.session_buttons.append(num_btn)
            vbox.addWidget(num_btn, alignment=Qt.AlignHCenter)
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

        # ---------- Retention summary ----------
        self.result_label = QLabel("")
        self.result_label.setStyleSheet("font-size: 16pt; font-weight: bold; color: #003366;")
        self.result_label.setWordWrap(True)
        self.result_label.setMinimumHeight(60)
        self.layout.addWidget(self.result_label)

        # ---------- Charts Section ----------
        self.add_charts(self.layout, self.df)

        # ---------- Run Geospatial Mapping Button ----------
        self.run_geo_btn = QPushButton("Run Ward Geospatial Mapping")
        self.run_geo_btn.setStyleSheet("""
            QPushButton {
                background-color: #185C37;
                color: white;
                font-weight: bold;
                padding: 8px;
            }
            QPushButton:disabled {
                background-color: #555555;
                color: #dddddd;
            }
        """)
        self.run_geo_btn.clicked.connect(self.run_geospatial)
        self.layout.addWidget(self.run_geo_btn, alignment=Qt.AlignCenter)

        self.set_and_run(4, self.session_buttons[3])

    # ---------- Chart Function ----------
    def add_charts(self, layout, df):
        # Helper function to rotate and align x-axis labels safely
        def rotate_labels(ax):
            for label in ax.get_xticklabels():
                label.set_rotation(45)
                label.set_ha('right')

        # 1. Gender by Programme
        fig1, ax1 = plt.subplots(figsize=(10, 5))
        data1 = df.groupby(["Activity type", "Gender"]).size().unstack(fill_value=0)
        data1.plot(kind="bar", stacked=True, ax=ax1)
        ax1.set_title("Gender Distribution by Programme (Activity Type)")
        ax1.set_xlabel("Programme (Activity Type)")
        ax1.set_ylabel("Number of Participants")
        rotate_labels(ax1)
        fig1.tight_layout()
        fig1.subplots_adjust(bottom=0.25)
        layout.addWidget(FigureCanvas(fig1))

        # 2. Age by Programme
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        df_age = df.dropna(subset=["RajiNewColumn-Age"]).copy()
        df_age["Age Range"] = pd.cut(
            df_age["RajiNewColumn-Age"],
            bins=[0, 12, 17, 22, 30, 40, 100],
            labels=["0–12", "13–17", "18–22", "23–30", "31–40", "40+"],
        )
        data2 = df_age.groupby(["Activity type", "Age Range"]).size().unstack(fill_value=0)
        data2.plot(kind="bar", stacked=True, ax=ax2)
        ax2.set_title("Age Distribution by Programme")
        ax2.set_xlabel("Programme (Activity Type)")
        ax2.set_ylabel("Number of Participants")
        rotate_labels(ax2)
        fig2.tight_layout()
        fig2.subplots_adjust(bottom=0.25)
        layout.addWidget(FigureCanvas(fig2))

        # 3. Programme Participation Volume
        fig3, ax3 = plt.subplots(figsize=(10, 5))
        df["Activity type"].value_counts().plot(kind="bar", ax=ax3, color="orange")
        ax3.set_title("Programme Participation Volume")
        ax3.set_xlabel("Programme (Activity Type)")
        ax3.set_ylabel("Count")
        rotate_labels(ax3)
        fig3.tight_layout()
        fig3.subplots_adjust(bottom=0.25)
        layout.addWidget(FigureCanvas(fig3))

        # 4. Constituency by Programme
        fig4, ax4 = plt.subplots(figsize=(10, 5))
        data4 = df.groupby(["Activity type", "Constituency"]).size().unstack(fill_value=0)
        top_cols = data4.sum().sort_values(ascending=False).head(8).index
        data4[top_cols].plot(kind="bar", stacked=True, ax=ax4)
        ax4.set_title("Top Constituencies by Programme")
        ax4.set_xlabel("Programme (Activity Type)")
        ax4.set_ylabel("Number of Participants")
        rotate_labels(ax4)
        fig4.tight_layout()
        fig4.subplots_adjust(bottom=0.25)
        layout.addWidget(FigureCanvas(fig4))

    # ---------- Retention & Utilities ----------
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
        df_retained = df_eligible[df_eligible["Attendee ID"].isin(retained)]
        grouped = df_retained.groupby("Attendee ID")["Activity type"].apply(lambda x: ", ".join(sorted(set(x)))).reset_index()
        grouped.to_excel(export_file, index=False)
        if sys.platform == "win32":
            os.startfile(export_file)
        elif sys.platform == "darwin":
            os.system(f"open '{export_file}'")
        else:
            os.system(f"xdg-open '{export_file}'")

    def run_geospatial(self):
        self.run_geo_btn.setEnabled(False)
        self.run_geo_btn.setText("Opening Ward Geospatial Mapping...")
        self.geo_process = QProcess(self)
        self.geo_process.finished.connect(self.enable_geo_button)
        self.geo_process.start(sys.executable, ["ward-geospatial-mapping.py"])

    def enable_geo_button(self):
        self.run_geo_btn.setEnabled(True)
        self.run_geo_btn.setText("Run Ward Geospatial Mapping")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app_window = RetentionApp()
    app_window.showMaximized()
    sys.exit(app.exec_())
