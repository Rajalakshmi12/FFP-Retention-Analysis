import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton
import matplotlib.pyplot as plt

# Dummy function for chart display (replace with actual logic)
def show_chart(label):
    plt.figure()
    plt.title(f"{label} - Weekly Retention")
    plt.plot([1, 2, 3, 4], [4, 5, 6, 7])  # Replace with actual data
    plt.xlabel("Week")
    plt.ylabel("Users")
    plt.grid(True)
    plt.show()

class RetentionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Retention Analysis")
        self.setGeometry(100, 100, 300, 250)

        layout = QVBoxLayout()

        # Define buttons
        buttons = {
            "Once a Week": lambda: show_chart("Once"),
            "Twice a Week": lambda: show_chart("Twice"),
            "Thrice a Week": lambda: show_chart("Thrice"),
            "Seven Days a Week": lambda: show_chart("Seven")
        }

        for label, action in buttons.items():
            btn = QPushButton(label)
            btn.clicked.connect(action)
            layout.addWidget(btn)

        self.setLayout(layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RetentionApp()
    window.show()
    sys.exit(app.exec_())
