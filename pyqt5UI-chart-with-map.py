import sys
import os
import pandas as pd
from PyQt5.QtCore import Qt, QProcess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QScrollArea, QFrame, QSizePolicy
)
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QBrush

# Matplotlib (embedded charts)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt


def make_excel_icon(size=32):
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    p = QPainter(pixmap)
    p.fillRect(0, 0, size // 2, size, QBrush(QColor("#107C41")))
    p.fillRect(size // 2, 0, size // 2, size, QBrush(QColor("#185C37")))

    p.setPen(Qt.white)
    p.setFont(QFont("Arial", int(size * 0.6), QFont.Bold))
    p.drawText(pixmap.rect(), Qt.AlignCenter, "X")
    p.end()

    return QIcon(pixmap)


class RetentionApp(QWidget):
    def __init__(self):
        super().__init__()

        self.geo_process = None
        self.session_buttons = []

        self.setWindowTitle("FFP Retention & Engagement Dashboard")

        # ---------------- SCROLL AREA ----------------
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

        container = QWidget()
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.layout = QVBoxLayout(container)
        self.layout.setContentsMargins(20, 10, 20, 10)
        self.layout.setSpacing(10)

        scroll.setWidget(container)

        main = QVBoxLayout(self)
        main.addWidget(scroll)

        # ---------------- LOAD DATA ----------------
        df = pd.read_excel("Documents/FFP_cleansed.xlsx", sheet_name="Main")
        df["Date"] = pd.to_datetime(df["Date"])
        self.df = df

        # ---------------- TITLE ----------------
        title = QLabel("Select session count for retention calculation (monthly):")
        title.setStyleSheet("font-size: 11pt; margin-bottom: 4px;")
        self.layout.addWidget(title)

        # ---------------- BUTTON ROW ----------------
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setSpacing(5)
        row_layout.setContentsMargins(0, 0, 0, 0)

        excel_icon = make_excel_icon(22)

        for i in range(1, 11):
            col = QVBoxLayout()
            col.setSpacing(0)
            col.setContentsMargins(0, 0, 0, 0)

            btn = QPushButton(str(i))
            btn.setFixedSize(48, 26)
            btn.setAutoDefault(False)
            btn.setDefault(False)
            btn.clicked.connect(lambda _, v=i, b=btn: self.set_and_run(v, b))
            self.session_buttons.append(btn)
            col.addWidget(btn, alignment=Qt.AlignCenter)

            ex = QPushButton()
            ex.setIcon(excel_icon)
            ex.setFixedSize(22, 22)
            ex.setStyleSheet("background:none; border:none;")
            ex.setAutoDefault(False)
            ex.setDefault(False)
            ex.clicked.connect(lambda _, v=i: self.generate_excel(v))
            col.addWidget(ex, alignment=Qt.AlignCenter)

            row_layout.addLayout(col)

        self.layout.addWidget(row)

        # ---------------- RETENTION RESULT ----------------
        self.result_label = QLabel("")
        self.result_label.setStyleSheet("font-size: 13pt; font-weight:bold; color:#003366;")
        self.layout.addWidget(self.result_label)

        # ---------------- GEO BUTTON ----------------
        self.geo_btn = QPushButton("Run Ward Geospatial Mapping")
        self.geo_btn.setFixedHeight(38)
        self.geo_btn.setStyleSheet("""
            QPushButton {
                background-color:#185C37; 
                color:white; 
                font-size:12pt;
                border-radius:4px;
            }
        """)
        self.geo_btn.setAutoDefault(False)
        self.geo_btn.setDefault(False)
        self.geo_btn.clicked.connect(self.launch_geospatial)
        self.layout.addWidget(self.geo_btn, alignment=Qt.AlignCenter)

        # ---------------- GAP SECTION ----------------
        self.add_gap_section()

        self.layout.addStretch(0)

        # Default selection
        self.set_and_run(4, self.session_buttons[3])

    # ---------------------------------------------------------
    # RETENTION LOGIC (6 months, split 3+3)
    # ---------------------------------------------------------
    def set_and_run(self, val, btn):
        for b in self.session_buttons:
            b.setStyleSheet("")
        btn.setStyleSheet("background-color:green; color:white;")
        self.calc_retention(val)

    def calc_retention(self, sessions):
        """
        Corrected:
        - Baseline = participants active in FIRST 3 months of the last-6-month window (any attendance).
        - Retained = from baseline, participants who hit >=sessions in at least 1 month in LAST 3 months.
        """
        df = self.df.copy()
        latest = df["Date"].max().normalize()

        start_6m = (latest - pd.DateOffset(months=6)).normalize()
        mid_3m = (start_6m + pd.DateOffset(months=3)).normalize()

        # 6-month window
        dfw = df[(df["Date"] >= start_6m) & (df["Date"] <= latest)].copy()
        if dfw.empty:
            self.result_label.setText("No participation found in the last 6 months window.")
            return

        dfw["Month"] = dfw["Date"].dt.to_period("M")
        dfw["Half"] = dfw["Month"].dt.to_timestamp().apply(lambda d: "FIRST3" if d < mid_3m else "LAST3")

        # Baseline: active at least once in FIRST3
        baseline_ids = set(dfw[dfw["Half"] == "FIRST3"]["Attendee ID"].unique())
        if not baseline_ids:
            self.result_label.setText("No participants found in the first 3 months of the 6-month window.")
            return

        # Monthly session counts for baseline participants
        m = (
            dfw[dfw["Attendee ID"].isin(baseline_ids)]
            .groupby(["Attendee ID", "Month"])
            .size()
            .reset_index(name="count")
        )
        m["Half"] = m["Month"].dt.to_timestamp().apply(lambda d: "FIRST3" if d < mid_3m else "LAST3")

        # Retained: baseline who hit threshold in >=1 month in LAST3
        retained_ids = set(m[(m["Half"] == "LAST3") & (m["count"] >= sessions)]["Attendee ID"].unique())

        pct = round(100 * len(retained_ids) / len(baseline_ids), 2) if baseline_ids else 0
        self.result_label.setText(
            f"{len(retained_ids)} of {len(baseline_ids)} participants retained in the last 3 months out of 6 months ({pct}%)."
        )

    # ---------------------------------------------------------
    # EXCEL EXPORT FOR RETENTION (does not close app)
    # ---------------------------------------------------------
    def generate_excel(self, sessions):
        df = self.df.copy()
        latest = df["Date"].max().normalize()

        start_6m = (latest - pd.DateOffset(months=6)).normalize()
        mid_3m = (start_6m + pd.DateOffset(months=3)).normalize()

        dfw = df[(df["Date"] >= start_6m) & (df["Date"] <= latest)].copy()
        if dfw.empty:
            return

        dfw["Month"] = dfw["Date"].dt.to_period("M")
        dfw["Half"] = dfw["Month"].dt.to_timestamp().apply(lambda d: "FIRST3" if d < mid_3m else "LAST3")

        baseline_ids = set(dfw[dfw["Half"] == "FIRST3"]["Attendee ID"].unique())
        if not baseline_ids:
            return

        m = (
            dfw[dfw["Attendee ID"].isin(baseline_ids)]
            .groupby(["Attendee ID", "Month"])
            .size()
            .reset_index(name="count")
        )
        m["Half"] = m["Month"].dt.to_timestamp().apply(lambda d: "FIRST3" if d < mid_3m else "LAST3")

        retained_ids = sorted(list(set(m[(m["Half"] == "LAST3") & (m["count"] >= sessions)]["Attendee ID"].unique())))

        # Export table
        out = pd.DataFrame({"Attendee ID": retained_ids})

        # Add demographics if present
        for c in ["Gender", "Ward", "Constituency"]:
            if c in df.columns:
                demo = df.dropna(subset=[c]).drop_duplicates("Attendee ID")[["Attendee ID", c]]
                out = out.merge(demo, on="Attendee ID", how="left")

        # Add first/last session inside the 6-month window
        first_last = dfw[dfw["Attendee ID"].isin(retained_ids)].groupby("Attendee ID")["Date"].agg(
            FirstSessionInWindow="min",
            LastSessionInWindow="max"
        ).reset_index()
        out = out.merge(first_last, on="Attendee ID", how="left")

        out["WindowStart"] = start_6m.date()
        out["WindowMid"] = mid_3m.date()
        out["WindowEnd"] = latest.date()
        out["ThresholdSessionsPerMonth"] = sessions

        folder = "Retention Excels"
        os.makedirs(folder, exist_ok=True)
        filename = f"retained_last6m_split3m_threshold_{sessions}.xlsx"
        path = os.path.join(folder, filename)
        out.to_excel(path, index=False)

        # open file without affecting app
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f"open '{path}'")
            else:
                os.system(f"xdg-open '{path}'")
        except Exception:
            pass

    # ---------------------------------------------------------
    # GAP SECTION (WITH VISUALIZATION)
    # ---------------------------------------------------------
    def add_gap_section(self):
        frame = QFrame()
        lay = QVBoxLayout(frame)
        lay.setSpacing(8)
        lay.setContentsMargins(10, 5, 10, 5)

        title = QLabel("Attendance Gap Analysis")
        title.setStyleSheet("font-size: 13pt; font-weight: bold; color:#660000;")
        lay.addWidget(title)

        df = self.df.sort_values(["Attendee ID", "Date"])
        df["Prev"] = df.groupby("Attendee ID")["Date"].shift(1)
        df["Gap"] = (df["Date"] - df["Prev"]).dt.days

        gap3 = df[df["Gap"] > 90].groupby("Attendee ID")["Gap"].max().sort_values(ascending=False)
        gap6 = df[df["Gap"] > 180].groupby("Attendee ID")["Gap"].max().sort_values(ascending=False)

        summary = QLabel(f">3-month gaps: {len(gap3)}    |    >6-month gaps: {len(gap6)}")
        summary.setStyleSheet("font-size: 11pt; font-weight:bold;")
        lay.addWidget(summary)

        excel_icon = make_excel_icon(22)

        export3 = QPushButton(" Export >3-Month Gap List")
        export3.setIcon(excel_icon)
        export3.setAutoDefault(False)
        export3.setDefault(False)
        export3.clicked.connect(lambda: self.export_gap_excel(3))
        lay.addWidget(export3)

        export6 = QPushButton(" Export >6-Month Gap List")
        export6.setIcon(excel_icon)
        export6.setAutoDefault(False)
        export6.setDefault(False)
        export6.clicked.connect(lambda: self.export_gap_excel(6))
        lay.addWidget(export6)

        if len(gap3) > 0:
            chart3 = self.plot_gap_chart(gap3, "Gap Chart (>3 Months)")
            lay.addWidget(chart3)

        if len(gap6) > 0:
            chart6 = self.plot_gap_chart(gap6, "Gap Chart (>6 Months)")
            lay.addWidget(chart6)

        self.layout.addWidget(frame)

    def plot_gap_chart(self, series, title):
        fig, ax = plt.subplots(figsize=(8, max(4, len(series) * 0.35)))

        attendees = series.index.astype(str)
        gaps_months = (series.values / 30).round(1)

        bars = ax.barh(attendees, gaps_months, color="#C04000")

        for bar, value in zip(bars, gaps_months):
            ax.text(
                bar.get_width() * 0.5,
                bar.get_y() + bar.get_height() / 2,
                f"{value} mo",
                ha="center",
                va="center",
                color="white",
                fontsize=9,
                fontweight="bold"
            )

        ax.margins(y=0)
        ax.set_xlabel("Gap Length (Months)")
        ax.set_ylabel("Attendee ID")
        ax.set_title(title, pad=10, fontsize=12, fontweight="bold")
        ax.invert_yaxis()

        fig.subplots_adjust(top=0.90)
        ax2 = ax.secondary_xaxis('top')
        ax2.set_xlabel("Gap Length (Months)")

        plt.tight_layout(pad=0.2)
        return FigureCanvas(fig)

    def export_gap_excel(self, months):
        df = self.df.sort_values(["Attendee ID", "Date"]).copy()
        df["Prev"] = df.groupby("Attendee ID")["Date"].shift(1)
        df["GapDays"] = (df["Date"] - df["Prev"]).dt.days

        threshold = 90 if months == 3 else 183
        filename = "gap_over_3_months.xlsx" if months == 3 else "gap_over_6_months.xlsx"

        gap_df = df[df["GapDays"] > threshold].copy()

        cols = ["Attendee ID", "Gender", "Ward", "Prev", "Date", "GapDays"]
        export_cols = [c for c in cols if c in gap_df.columns]

        final_df = gap_df[export_cols].copy()
        final_df.rename(columns={
            "Prev": "Gap Start Date",
            "Date": "Gap End Date",
            "GapDays": "Gap (Days)"
        }, inplace=True)

        final_df.sort_values(by="Gap (Days)", ascending=False, inplace=True)

        folder = "Gap Analysis Excels"
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, filename)

        final_df.to_excel(path, index=False)

        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            os.system(f"open '{path}'")
        else:
            os.system(f"xdg-open '{path}'")

    # ---------------------------------------------------------
    # GEOSPATIAL BUTTON
    # ---------------------------------------------------------
    def launch_geospatial(self):
        self.geo_btn.setEnabled(False)
        self.geo_btn.setText("Opening Ward Geospatial Mapping...")

        if self.geo_process is None:
            self.geo_process = QProcess(self)

        self.geo_process.finished.connect(self._geo_complete)
        self.geo_process.start(sys.executable, ["ward-geospatial-mapping.py"])

    def _geo_complete(self):
        self.geo_btn.setEnabled(True)
        self.geo_btn.setText("Run Ward Geospatial Mapping")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = RetentionApp()
    win.showMaximized()
    sys.exit(app.exec_())
