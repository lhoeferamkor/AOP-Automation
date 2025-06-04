import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QDateEdit, QGroupBox, QFormLayout,
    QCheckBox, QProgressBar, QSpacerItem, QSizePolicy, 
    QToolButton, QSpinBox, QFileDialog, QTabWidget, QTableWidget,
    QSplitter, QHeaderView, QTableWidgetItem
)
from PyQt5.QtCore import QDate, Qt, QSize, QDir
from PyQt5.QtGui import QFont, QTextCursor, QIcon, QColor

import datetime
import time
import random
import subprocess
import os 

import SAP_File_Automation as file_reader
import remove_specified_rows as trimmer

class SearchApp(QWidget):
    def __init__(self):
        super().__init__()
        self.function_progress_bars = {}
        self.function_checkboxes = {} # Store checkboxes for easier access
        self.icon_arrow_right = "►" # Placeholder for QIcon(QDir.currentPath() + "/icons/arrow_right.png")
        self.icon_arrow_down = "▼" # Placeholder for QIcon(QDir.currentPath() + "/icons/arrow_down.png")
        
        self.setWindowIcon(QIcon("Amkor_icon.png"))

        self.initUI()
        self.apply_styles()

    def initUI(self):
        self.setWindowTitle('AOP Automation')
        self.setGeometry(250, 150, 650, 730) # Adjusted height slightly

        # --- Add QTabWidget at the top ---
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.North)

        # --- Main Tab ---
        main_tab_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(12)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(12)

        # --- Input Group (Search Criteria) ---
        self.input_group = QGroupBox("Search Criteria")
        form_layout = QFormLayout()
        form_layout.setContentsMargins(10, 25, 10, 10)
        form_layout.setSpacing(8)
        
        # --- Select Input File --- 
        self.download_in_label = QLabel("Read File: ")
        download_in_layout = QHBoxLayout() # For text box and button
        download_in_layout.setSpacing(5)
        self.download_in_input = QLineEdit()
        self.download_in_input.setText("AOP Automation Scripts/input_data/ZANALYSIS_PATTERN.xls") 
        self.download_in_input.setReadOnly(True) # Make it read-only, changed by button
        download_in_layout.addWidget(self.download_in_input, 1)
        self.browse_button_in = QPushButton("Select") # Browse button
        self.browse_button_in.setFixedWidth(80)
        self.browse_button_in.clicked.connect(self.browse_download_file_in)
        download_in_layout.addWidget(self.browse_button_in)
        form_layout.addRow(self.download_in_label, download_in_layout)

        # --- Select Output File ---
        self.download_out_label = QLabel("Write File: ")
        download_out_layout = QHBoxLayout() # For text box and button
        download_out_layout.setSpacing(5)
        self.download_out_input = QLineEdit()
        self.download_out_input.setText("AOP Automation Scripts/output_data") 
        self.download_out_input.setReadOnly(True) # Make it read-only, changed by button
        download_out_layout.addWidget(self.download_out_input, 1)
        self.browse_button_out = QPushButton("Select") # Browse button
        self.browse_button_out.setFixedWidth(80)
        self.browse_button_out.text
        self.browse_button_out.clicked.connect(self.browse_download_file_out)
        download_out_layout.addWidget(self.browse_button_out)
        form_layout.addRow(self.download_out_label, download_out_layout)
        self.input_group.setLayout(form_layout)
        main_layout.addWidget(self.input_group)


        # --- Function Selection Group (Available Tasks) ---
        self.functions_group = QGroupBox("Available Tasks")
        functions_main_h_layout = QHBoxLayout() # Horizontal layout: Checkboxes | Spacer | ProgressBars
        functions_main_h_layout.setContentsMargins(10, 25, 10, 10)
        functions_main_h_layout.setSpacing(15) # Space between checkbox column and progress bar column

        # Column for Checkboxes
        checkbox_v_layout = QVBoxLayout()
        checkbox_v_layout.setSpacing(8) # Spacing between checkboxes

        # Column for Progress Bars
        progress_v_layout = QVBoxLayout()
        progress_v_layout.setSpacing(8) # Spacing between progress bars

        task_definitions = [
            ("download", "Download SAP (Not Yet Functional)"),
            ("convert", "Convert File"),
            ("highlight", "Higlight Rows"),
            ("remove", "Remove Rows"),
            ("configure", "Configure for Report")
        ]

        for key, display_text in task_definitions:
            # Checkbox
            checkbox = QCheckBox(display_text)
            checkbox.setChecked(True)
            self.function_checkboxes[key] = checkbox
            checkbox_v_layout.addWidget(checkbox)

            # Progress Bar
            progress_bar = QProgressBar()
            progress_bar.setTextVisible(True)
            progress_bar.setAlignment(Qt.AlignCenter)
            progress_bar.setRange(0, 100)
            progress_bar.setValue(0)
            # progress_bar.setFixedWidth(200) # Option 1: Fixed width for all
            self.function_progress_bars[key] = progress_bar
            progress_v_layout.addWidget(progress_bar)

        checkbox_v_layout.addStretch(1) # Pushes checkboxes up if space
        progress_v_layout.addStretch(1) # Pushes progress bars up

        functions_main_h_layout.addLayout(checkbox_v_layout, 1) # Checkboxes take some space
        # functions_main_h_layout.addSpacerItem(QSpacerItem(20, 10, QSizePolicy.Fixed, QSizePolicy.Minimum)) # Fixed spacer
        functions_main_h_layout.addLayout(progress_v_layout, 2) # Progress bars take more space

        self.functions_group.setLayout(functions_main_h_layout)
        main_layout.addWidget(self.functions_group)

        # --- Execution Controls Group ---
        self.execution_controls_group = QGroupBox("Execution Controls")
        exec_controls_outer_h_layout = QHBoxLayout() # Main Horizontal layout for this group
        exec_controls_outer_h_layout.setContentsMargins(10, 25, 10, 10)
        exec_controls_outer_h_layout.setSpacing(10)

        # Vertical layout for checkboxes
        checkboxes_v_layout_exec = QVBoxLayout()
        checkboxes_v_layout_exec.setSpacing(5)
        self.cb_build_files = QCheckBox("Build Multiple Files")
        self.cb_build_files.setChecked(True) # Default to checked
        checkboxes_v_layout_exec.addWidget(self.cb_build_files)
        self.open_on_finish = QCheckBox("Open on Finish")
        self.open_on_finish.setChecked(True) # Default to checked
        checkboxes_v_layout_exec.addWidget(self.open_on_finish)
        checkboxes_v_layout_exec.addStretch(1) # Push checkboxes up

        exec_controls_outer_h_layout.addLayout(checkboxes_v_layout_exec)
        exec_controls_outer_h_layout.addSpacerItem(QSpacerItem(20, 10, QSizePolicy.Expanding, QSizePolicy.Minimum)) # Spacer

        self.run_button = QPushButton('Run Tasks') # Shortened text
        self.run_button.clicked.connect(self.on_run_clicked)
        self.run_button.setFixedHeight(35)
        self.run_button.setMinimumWidth(200) # Ensure button has decent width
        exec_controls_outer_h_layout.addWidget(self.run_button, 0, Qt.AlignVCenter) # Align button vertically centered

        self.execution_controls_group.setLayout(exec_controls_outer_h_layout)
        main_layout.addWidget(self.execution_controls_group)


        # --- Results Area (Output Log) ---
        self.results_group = QGroupBox("Output Log")
        results_layout = QVBoxLayout()
        results_layout.setContentsMargins(10, 25, 10, 10)
        self.results_output = QTextEdit(readOnly=True, placeholderText="Logs and results will appear here...")
        results_layout.addWidget(self.results_output)
        self.results_group.setLayout(results_layout)
        main_layout.addWidget(self.results_group, 1)

        main_tab_widget.setLayout(main_layout)
        self.tabs.addTab(main_tab_widget, "Main")

        # --- Advanced Settings Tab ---
        self.advanced_tab_widget = QWidget()
        self.advanced_layout = QVBoxLayout()

        # --- Advanced Variables - Remove and Keep Criteria 
        self.advanced_variables_group = QGroupBox("Design Rules")
        functions_av_group = QVBoxLayout()  # Use QVBoxLayout for full vertical expansion

        # --- Splitter ---
        splitter = QSplitter(Qt.Horizontal) # Horizontal splitter

        # --- "Remove" Table ---
        remove_vbox = QVBoxLayout()
        self.remove_table_button = QPushButton("Remove Criteria")
        self.remove_table_button.setStyleSheet("""
            font-size: 13pt;
            font-weight: bold;
            color: "#e0e7ef";
            background: "#5c7da8";
            border-radius: 6px;
            padding: 6px 0 6px 0;
            margin-bottom: 4px;
            """)
        self.remove_table_button.clicked.connect(self.empty_rtable_row)

        self.remove_table = QTableWidget()
        self.remove_table.setColumnCount(2)
        self.remove_table.setHorizontalHeaderLabels(["Product", "PL"])
        self.setup_table_appearance(self.remove_table, "Remove", QColor(255, 200, 200)) # Light Red
        self.remove_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.remove_table.cellChanged.connect(self.update_remove_table)
        remove_vbox.addWidget(self.remove_table_button)
        remove_vbox.addWidget(self.remove_table)
        remove_widget = QWidget()
        remove_widget.setLayout(remove_vbox)

        # --- "Keep" Table ---
        keep_vbox = QVBoxLayout()
        self.keep_table_button = QPushButton("Keep Criteria")
        self.keep_table_button.setStyleSheet("""
            font-size: 13pt;
            font-weight: bold;
            color: "#e0e7ef";
            background: "#5c7da8";
            border-radius: 6px;
            padding: 6px 0 6px 0;
            margin-bottom: 4px;
            """)
        self.keep_table_button.clicked.connect(self.empty_ktable_row)

        self.keep_table = QTableWidget()
        self.keep_table.setColumnCount(2)
        self.keep_table.setHorizontalHeaderLabels(["Product", "PL"])
        self.setup_table_appearance(self.keep_table, "Keep", QColor(200, 255, 200)) # Light Green
        self.keep_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.keep_table.cellChanged.connect(self.update_keep_table)
        keep_vbox.addWidget(self.keep_table_button)
        keep_vbox.addWidget(self.keep_table)
        keep_widget = QWidget()
        keep_widget.setLayout(keep_vbox)
        

        # --- Add values to Tables ---
        for name in trimmer.red_keywords_group:
            self.add_table_row(self.remove_table, name, 0)
        
        for name in trimmer.green_keywords_complex:
            self.add_table_row(self.keep_table, name, 1)

        for name in trimmer.green_keywords_group:
            self.add_table_row(self.keep_table, name, 0)

        # Add tables to the splitter
        splitter.addWidget(remove_widget)
        splitter.addWidget(keep_widget)
        splitter.setSizes([1, 1])  # Make both tables expand equally

        functions_av_group.addWidget(splitter)

        self.advanced_variables_group.setLayout(functions_av_group)
        self.advanced_variables_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.advanced_layout.addWidget(self.advanced_variables_group, stretch=1)
        self.advanced_tab_widget.setLayout(self.advanced_layout)
        self.tabs.addTab(self.advanced_tab_widget, "Advanced Settings")
        # --- End of Advanced Variables - Remove and Keep Criteria 

        # Set the main layout for the window to just the tabs
        window_layout = QVBoxLayout()
        window_layout.addWidget(self.tabs)
        self.setLayout(window_layout)
        self.show()

    def add_table_row(self, table : QTableWidget, text : str, pos : int):
        row_idx = table.rowCount()
        table.insertRow(row_idx)
        text_item = QTableWidgetItem(text)
        dash_item = QTableWidgetItem(" - ")
        text_item.setTextAlignment(Qt.AlignCenter)
        dash_item.setTextAlignment(Qt.AlignCenter)
        table.setItem(row_idx, pos, text_item)
        table.setItem(row_idx, abs(pos-1), dash_item)

    def empty_rtable_row(self):
        row_idx = self.remove_table.rowCount()
        self.remove_table.insertRow(row_idx)
        # Optionally, add empty QTableWidgetItems to make the cells editable
        self.remove_table.setItem(row_idx, 0, QTableWidgetItem(""))
        self.remove_table.setItem(row_idx, 1, QTableWidgetItem(""))

    def empty_ktable_row(self):
        row_idx = self.keep_table.rowCount()
        self.keep_table.insertRow(row_idx)
        # Optionally, add empty QTableWidgetItems to make the cells editable
        self.keep_table.setItem(row_idx, 0, QTableWidgetItem(""))
        self.keep_table.setItem(row_idx, 1, QTableWidgetItem(""))

    def update_remove_table(self, row, column):
        self.remove_table.item(row, column).setTextAlignment(Qt.AlignCenter)
        other_col = abs(column - 1)
        if not self.remove_table.item(row, other_col) or not self.remove_table.item(row, other_col).text():
            self.remove_table.setItem(row, other_col, QTableWidgetItem(" - "))

    def update_keep_table(self, row, column):
        self.keep_table.item(row, column).setTextAlignment(Qt.AlignCenter)
        other_col = abs(column - 1)
        if not self.keep_table.item(row, other_col) or not self.keep_table.item(row, other_col).text():
            self.keep_table.setItem(row, other_col, QTableWidgetItem(" - "))
    
    def save_tables(self):
        #Save both tables and associated values 
        pass

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
        table.verticalHeader().setVisible(False)
    
    def browse_download_file_in(self):
        options = QFileDialog.Options()
        file = QFileDialog.getOpenFileName(self, "Select a File", "", "All Files (*);;Excel Files (*.xls)", options=options)
        if file: # If a directory was selected
            self.download_in_input.setText(file[0])
            QApplication.processEvents()

    def browse_download_file_out(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Download Directory", self.download_out_input.text())
        if directory: # If a directory was selected
            self.download_out_input.setText(directory)

    def update_progress(self, function_key, percentage):
        if function_key in self.function_progress_bars:
            self.function_progress_bars[function_key].setValue(percentage)
            QApplication.processEvents()

    def on_run_clicked(self):
        is_build_files = self.cb_build_files.isChecked()
        is_headless = self.open_on_finish.isChecked()

        for pb in self.function_progress_bars.values(): pb.setValue(0)

        selected_tasks = []
        for key, checkbox in self.function_checkboxes.items():
            if checkbox.isChecked():
                selected_tasks.append((key, checkbox.text())) # (key, display_name)

        if not selected_tasks:
            self.results_output.append("No tasks selected to run.")
            return

        self.run_button.setEnabled(False)
        try:
            temp_path = None
            for task_key, task_name_display in selected_tasks:
                progress_callback_for_task = lambda p, tk=task_key: self.update_progress(tk, p)
                if task_key == 'convert':
                    self.results_output.insertHtml(f'<b> Loading File {self.download_in_input.text()} ... </b>')
                    temp_path = os.path.join(self.download_out_input.text(), "temporary_file.xlsx")
                    QApplication.processEvents()
                    file_reader.convert_mhtml_to_excel(self.download_in_input.text(), temp_path)
                    self.results_output.insertHtml('<b><font color = "green"> DONE </font></b>')
                    self.results_output.append("")
                    QApplication.processEvents()
                    

                if task_key == 'highlight' or task_key == 'remove':
                    if task_key == 'highlight': self.results_output.insertHtml(f'<b> Highlighting Rows ... </b>')
                    if task_key == 'remove': self.results_output.insertHtml(f'<b> Removing Rows ... </b>')
                    QApplication.processEvents()  
                    if temp_path:
                        trimmer.apply_conditional_formatting(temp_path, os.path.join(self.download_out_input.text(), 'formatted_report.xlsx'), task=task_key)
                        self.results_output.insertHtml('<b><font color = "green"> DONE </font></b>')
                        self.results_output.append("")
                        QApplication.processEvents()
                    elif not temp_path:
                        try:
                            self.results_output.insertHtml('<b><font color = "red"> FAILED </font></b>')
                            self.results_output.append("")
                            self.results_output.insertHtml('<b><font color = "blue"> Couldnt download using conventional methods. Switching to direct download... </font></b>')
                            QApplication.processEvents()
                            trimmer.apply_conditional_formatting(self.download_in_input.text(), os.path.join(self.download_out_input.text(), 'formatted_report.xlsx'), sheet_name = task_key)
                        except Exception as e:
                            self.results_output.insertHtml('<b><font color = "red"> ERROR! Problem applying conditional Formatting </font></b>')
                            QApplication.processEvents()          
                    else:
                        self.results_output.insertHtml('<b><font color = "red"> FAILED </font></b>')
                        self.results_output.append("")
                        self.results_output.insertHtml('<b><font color = "red"> ERROR! No File Path From File convert. Could not find file destination or lookup </font></b>')
                        QApplication.processEvents()

                    if task_key == 'remove':
                        self.results_output.append("")
                        self.results_output.insertHtml(f'<b> Deleting File {self.download_in_input.text()} ... </b>')
                        QApplication.processEvents()
                        if os.path.exists(temp_path):
                            os.remove(temp_path)
                            self.results_output.insertHtml('<b><font color = "red"> DONE </font></b>')
                            QApplication.processEvents()
                self.function_progress_bars[task_key].setValue(100)

        finally:
            self.run_button.setEnabled(True)

    def apply_styles(self):
        MAIN_WINDOW_BACKGROUND = "#e0e7ef" # Slightly bluish gray
        GROUP_BOX_CONTENT_BACKGROUND = "#f8faff" # Very light blue
        BORDER_COLOR = "#5c7da8" # Softer blue border
        
        TITLE_COLOR = "#2c3e50" # Darker, less saturated blue for title
        PROGRESS_BAR_CHUNK_COLOR = "#3498db" # A nice blue
        BUTTON_BG_COLOR = "#3498db" # A friendly blue for button
        BUTTON_HOVER_COLOR = "#2980b9"

        self.setStyleSheet(f"""
            QWidget {{
                font-family: "Segoe UI", Arial, sans-serif; /* Common modern font */
                font-size: 9pt;
                font-weight: 500; 
            }}
            SearchApp {{
                background-color: {MAIN_WINDOW_BACKGROUND};
            }}
            
            QGroupBox {{
                background-color: {GROUP_BOX_CONTENT_BACKGROUND};
                border: 2px solid {BORDER_COLOR};
                border-radius: 12px;
                margin-top: 14px; /* Adjust for title height */
                /* Font for content INSIDE QGroupBoxes, if different from QWidget default */
                font-weight: 600;
                font-size: 11pt;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 1px 6px; /* Adjusted padding */
                left: 10px;
                background-color: {TITLE_COLOR};
                font-family: "Segoe UI", Arial, sans-serif;
                color: {MAIN_WINDOW_BACKGROUND};
                border-radius: 3px;
            }}
            QLineEdit, QDateEdit, QTextEdit {{
                background-color: white;
                border: 1px solid #bdc3c7; /* Lighter gray border */
                border-radius: 4px;
                padding: 6px;
                min-height: 22px;
            }}
            QDateEdit {{ padding-right: 2px; }}

            QPushButton {{
                background-color: {BUTTON_BG_COLOR};
                color: white;
                border: none;
                padding: 7px 15px; /* Adjusted padding */
                border-radius: 4px;
                font-weight: bold;
                min-height: 24px; /* Consistent height */
            }}
            QPushButton:hover {{
                background-color: {BUTTON_HOVER_COLOR};
            }}
            QPushButton:disabled {{
                background-color: #dbe0e3; /* Lighter disabled color */
                color: #7f8c8d;
            }}
            QCheckBox {{
                spacing: 6px;
                padding: 4px 0;
            }}
            QProgressBar {{
                border: 1px solid #b0bec5; /* Softer border for progress bar */
                border-radius: 4px;
                text-align: center;
                background-color: #eceff1; /* Light gray background */
                min-height: 22px; /* Match input fields */
                font-weight: bold; /* Make percentage text bold */
                color: #263238; /* Darker text for percentage */
            }}
            QProgressBar::chunk {{
                background-color: {PROGRESS_BAR_CHUNK_COLOR};
                border-radius: 3px;
                margin: 1px;
            }}

            QTableWidget {{
                border: 2px solid #000000;
                gridline-color: #000000;
                background-color: #eceff1;

                selection-background-color: #000000;
                selection-color: #000000;
            }}

            QTableWidget::item {{
                padding: 5px;                    
                border: 1px #000000;   
                border-right: #000000;
                border-bottom: #000000;          
                                                 
            }}
            QTableWidget::item:focus {{
                /* Optional: style for the cell currently being edited */
                 border: 1px solid #5cacee;
            }}


        """)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SearchApp()
    sys.exit(app.exec_())

