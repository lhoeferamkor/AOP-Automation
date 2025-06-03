import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QTableWidget, QTableWidgetItem, QSplitter, QHeaderView
)
from PyQt5.QtGui import QColor, QBrush, QPalette, QFont
from PyQt5.QtCore import Qt

class DualTableWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dual Tables with Splitter")
        self.setGeometry(100, 100, 800, 500)

        # Main widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- GroupBox ---
        group_box = QGroupBox("Product Categorization")
        group_box_layout = QVBoxLayout(group_box) # Layout for inside the GroupBox
        group_box.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid gray;
                border-radius: 5px;
                margin-top: 0.5em; /* Space for the title */
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center; /* Position of the title */
                padding: 0 3px;
                background-color: #f0f0f0; /* Light gray for title background */
                border-radius: 3px;
            }
        """)

        # --- Splitter ---
        splitter = QSplitter(Qt.Horizontal) # Horizontal splitter

        # --- "Remove" Table ---
        self.remove_table = QTableWidget()
        self.remove_table.setColumnCount(2)
        self.remove_table.setHorizontalHeaderLabels(["Product", "PL"])
        self.setup_table_appearance(self.remove_table, "Remove", QColor(255, 200, 200)) # Light Red

        # --- "Keep" Table ---
        self.keep_table = QTableWidget()
        self.keep_table.setColumnCount(2)
        self.keep_table.setHorizontalHeaderLabels(["Product", "PL"])
        self.setup_table_appearance(self.keep_table, "Keep", QColor(200, 255, 200)) # Light Green


        # Add tables to the splitter
        splitter.addWidget(self.remove_table)
        splitter.addWidget(self.keep_table)
        splitter.setSizes([400, 400]) # Initial equal sizing

        # Add splitter to the GroupBox layout
        group_box_layout.addWidget(splitter)
        main_layout.addWidget(group_box)

        # Populate with some dummy data
        self.populate_dummy_data(self.remove_table, 5)
        self.populate_dummy_data(self.keep_table, 5)


    def setup_table_appearance(self, table: QTableWidget, title: str, header_bg_color: QColor):
        """Helper function to set up common appearance for tables."""
        # Style the table itself
        table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid #c0c0c0;
                border-radius: 5px;
                gridline-color: #dcdcdc;
                background-color: #ffffff;
                selection-background-color: #a8d8ff;
                selection-color: #000000;
            }}
            QTableWidget::item {{
                padding: 5px;
                border-bottom: 1px solid #e8e8e8;
                border-right: 1px solid #e8e8e8;
            }}
            QTableWidget::item:focus {{
                 border: 1px solid #5cacee;
            }}
            QHeaderView::section:horizontal {{
                background-color: {header_bg_color.name()}; /* Dynamic color */
                color: #111111;
                padding: 6px;
                border-top-left-radius: 0px;
                border-top-right-radius: 0px;
                border-bottom: 2px solid #b0b8c0;
                border-right: 1px solid #c0c8d0;
                font-weight: bold;
            }}
            QHeaderView::section:horizontal:last {{
                border-right: 1px solid #b0b8c0;
            }}
            QHeaderView::section:vertical {{
                background-color: #f0f2f4;
                padding: 5px;
                border-right: 1px solid #c0c8d0;
                border-bottom: 1px solid #d0d8e0;
            }}
        """)

        # Optional: Make the overall table title (Not a standard QTableWidget feature)
        # We can achieve a similar effect by using the GroupBox title or adding a QLabel above
        # For individual header styling (Remove/Keep), we color the QHeaderView::section

        # Make columns stretch to fill available space
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch) # Stretch all columns
        # header.setSectionResizeMode(0, QHeaderView.Stretch) # Stretch specific column
        # header.setSectionResizeMode(1, QHeaderView.Interactive) # Let user resize PL

        # Set vertical header (row numbers) to be hidden or styled if needed
        # table.verticalHeader().setVisible(False)

    def populate_dummy_data(self, table: QTableWidget, num_rows: int):
        table.setRowCount(num_rows)
        for r in range(num_rows):
            product_item = QTableWidgetItem(f"Product {r+1:03d}")
            pl_item = QTableWidgetItem(f"PL{random.randint(100,999)}")
            table.setItem(r, 0, product_item)
            table.setItem(r, 1, pl_item)


# Dummy random import for populate_dummy_data
import random

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = DualTableWindow()
    main_win.show()
    sys.exit(app.exec_())