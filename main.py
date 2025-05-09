from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox, QPushButton,
    QDialog, QDateEdit, QTextEdit, QMessageBox, QTableWidget, QTableWidgetItem, QMenu, QLineEdit, QComboBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator
from datetime import datetime
import sys
import json
import pandas as pd
import os
import calendar


# Constants for file paths
SALARY_FILE = 'salary_data.json'
OVERTIME_FILE = 'overtime_data.json'

class OvertimeTrackerApp(QWidget):
    def __init__(self):
        super().__init__()

        # Load initial data
        self.salary = self.load_salary()
        self.overtime_entries = self.load_overtime_entries()
        self.overtime_entries.sort(key=lambda e: datetime.strptime(e['date'], '%d-%m-%Y'))

        # Initialize UI
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Overtime Tracker')
        layout = QVBoxLayout()

        self.setFixedWidth(340)

        # Salary Input and Real-Time Rate Calculations
        salary_layout = QHBoxLayout()
        self.salary_input = QLineEdit(str(self.salary))
        self.salary_input.setValidator(QIntValidator())
        self.salary_input.textChanged.connect(self.update_rates)
        salary_layout.addWidget(QLabel('Salary:'))
        salary_layout.addWidget(self.salary_input)

        # Overtime Rate Dropdown
        self.overtime_rate_combo = QComboBox()  # Initialize QComboBox
        self.overtime_rate_combo.addItems(['x1', 'x1.5', 'x2', 'x3'])
        self.overtime_rate_combo.setCurrentIndex(0)  # Default to x1
        self.overtime_rate_combo.currentIndexChanged.connect(self.update_rates)  # Update rates when changed
        salary_layout.addWidget(QLabel('Overtime Rate:'))
        salary_layout.addWidget(self.overtime_rate_combo)

        layout.addLayout(salary_layout)

        # Add some space above the rates layout
        layout.addSpacing(5)

        # Daily and Hourly Rate Display - Place them in a horizontal layout
        rates_layout = QHBoxLayout()
        self.daily_rate_label = QLabel('Daily Rate: 0.00')
        self.hourly_rate_label = QLabel('Hourly Rate: 0.00')

        # Add "|" separator
        separator_label = QLabel(' | ')

        # Center-align the labels
        self.daily_rate_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hourly_rate_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        separator_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Add labels to the layout
        rates_layout.addWidget(self.daily_rate_label)
        rates_layout.addWidget(separator_label)
        rates_layout.addWidget(self.hourly_rate_label)

        # Center the entire layout
        rates_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addLayout(rates_layout)

        # Add some space below the rates layout
        layout.addSpacing(5)

        # Add Entry Button
        add_entry_button = QPushButton('Add Entry')
        add_entry_button.clicked.connect(self.show_add_entry_dialog)
        layout.addWidget(add_entry_button)

        # Reset Entries Button
        reset_button = QPushButton('Reset Entries')
        reset_button.clicked.connect(self.reset_entries)
        layout.addWidget(reset_button)

        # Generate Report Button
        generate_report_button = QPushButton('Generate Report')
        generate_report_button.clicked.connect(self.generate_report)
        layout.addWidget(generate_report_button)

        # Show Overtime Entries Table Button
        self.show_entries_button = QPushButton('▼ Show Overtime Entries')
        self.show_entries_button.clicked.connect(self.toggle_overtime_table)
        layout.addWidget(self.show_entries_button)

        # Overtime Entries Table
        self.overtime_table = QTableWidget()
        self.overtime_table.setColumnCount(3)
        self.overtime_table.setHorizontalHeaderLabels(['Hours', 'Date', 'Task'])
        self.overtime_table.setVisible(False)
        self.overtime_table.cellChanged.connect(self.handle_cell_changed)  # Connect cellChanged signal
        self.overtime_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.overtime_table.customContextMenuRequested.connect(self.show_context_menu)
        layout.addWidget(self.overtime_table)

        # Info Label for displaying totals
        self.info_label = QLabel('Total Hours = 0  |  Days: 0.0 |  Amount = 0.00')
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Center the label
        layout.addWidget(self.info_label)

        self.update_overtime_table()
        self.setLayout(layout)
        self.update_rates()

    def load_salary(self):
        if os.path.exists(SALARY_FILE):
            with open(SALARY_FILE, 'r') as file:
                return json.load(file).get('salary', 0)
        return 0

    def save_salary(self, salary):
        with open(SALARY_FILE, 'w') as file:
            json.dump({'salary': salary}, file)

    def load_overtime_entries(self):
        if os.path.exists(OVERTIME_FILE):
            with open(OVERTIME_FILE, 'r') as file:
                entries = json.load(file)
                # Ensure all dates are converted to the desired format
                for entry in entries:
                    try:
                        # Attempt to parse dates in both formats
                        entry['date'] = pd.to_datetime(entry['date'], dayfirst=True).strftime('%d-%m-%Y')
                    except ValueError:
                        # Skip or handle errors in date parsing
                        pass
                return entries
        return []

    def save_overtime_entries(self):
        with open(OVERTIME_FILE, 'w') as file:
            json.dump(self.overtime_entries, file)

    def update_rates(self):
        salary = float(self.salary_input.text() or 0)
        self.save_salary(salary)

        # Get the current year and month
        now = datetime.now()
        current_year = now.year
        current_month = now.month

        # Get the number of days in the current month
        days_in_month = calendar.monthrange(current_year, current_month)[1]

        # Calculate daily and hourly rates
        daily_rate = salary / days_in_month
        hourly_rate = daily_rate / 8

        # Update labels
        self.daily_rate_label.setText(f'Daily Rate: {daily_rate:.2f}')
        self.hourly_rate_label.setText(f'Hourly Rate: {hourly_rate:.2f}')

        # Update the info label
        self.update_info_label()

    def get_overtime_multiplier(self):
        """ Get the selected overtime multiplier from the dropdown """
        multiplier_str = self.overtime_rate_combo.currentText().replace('x', '')
        return float(multiplier_str)

    def update_info_label(self):
        # Calculate total hours, days, and amount
        total_hours = sum(entry['hours'] for entry in self.overtime_entries)
        salary = float(self.salary_input.text() or 0)
        now = datetime.now()
        days_in_month = calendar.monthrange(now.year, now.month)[1]
        daily_rate = salary / days_in_month
        hourly_rate = daily_rate / 8
        overtime_multiplier = self.get_overtime_multiplier()
        total_days = total_hours / 8
        total_amount = total_hours * hourly_rate * overtime_multiplier

        # Update the info label
        self.info_label.setText(f'Hours = {total_hours}  |  Days: {total_days:.2f}  |  Amount = {total_amount:.2f}')

    def show_add_entry_dialog(self):
        dialog = AddEntryDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.overtime_entries.append(dialog.get_entry())
            self.overtime_entries.sort(key=lambda e: datetime.strptime(e['date'], '%d-%m-%Y'))
            self.save_overtime_entries()
            self.update_overtime_table()

    def toggle_overtime_table(self):
        if self.overtime_table.isVisible():
            # Hide the table and reset its height to zero
            self.overtime_table.setVisible(False)
            self.overtime_table.setFixedHeight(0)  # Set height to 0 to shrink the layout
            self.show_entries_button.setText('▼ Show Overtime Entries')
        else:
            # Show the table and set a fixed height to expand
            self.overtime_table.setVisible(True)
            self.overtime_table.setFixedHeight(200)  # Adjust the height as needed
            self.show_entries_button.setText('▲ Hide Overtime Entries')
        self.adjustSize()  # Adjust the size of the window

    def handle_cell_changed(self, row, column):
        """Persist edits made directly in the table."""
        try:
            item_text = self.overtime_table.item(row, column).text().strip()

            if column == 0:  # Hours
                new_hours = float(item_text)
                if new_hours < 0:
                    raise ValueError
                self.overtime_entries[row]['hours'] = new_hours

            elif column == 1:  # Date
                # validate DD-MM-YYYY
                datetime.strptime(item_text, '%d-%m-%Y')
                self.overtime_entries[row]['date'] = item_text
                self.overtime_entries.sort(key=lambda e: datetime.strptime(e['date'], '%d-%m-%Y'))

            elif column == 2:  # Task
                if not item_text:
                    raise ValueError
                self.overtime_entries[row]['task'] = item_text

            # save and refresh
            self.save_overtime_entries()
            self.update_info_label()
            self.update_overtime_table()  # ← add this line

        except (ValueError, AttributeError):
            QMessageBox.warning(self, 'Invalid Input',
                                'Please enter:\n• a positive number for Hours\n'
                                '• a valid date DD-MM-YYYY for Date\n'
                                '• non-empty text for Task')
            self.update_overtime_table()  # revert to previous values

    def show_context_menu(self, position):
        """Show a context menu when right-clicking on a table row."""
        index = self.overtime_table.indexAt(position)
        if index.isValid():  # Ensure a valid row is clicked
            menu = QMenu(self)
            delete_action = menu.addAction("Delete Entry")
            action = menu.exec(self.overtime_table.viewport().mapToGlobal(position))
            if action == delete_action:
                self.delete_entry()

    # Add this method to the OvertimeTrackerApp class
    def delete_entry(self):
        """Delete the selected entry from the table and update all variables."""
        current_row = self.overtime_table.currentRow()
        if current_row >= 0:  # Ensure a row is selected
            # Confirm deletion with the user
            reply = QMessageBox.question(
                self,
                'Delete Entry',
                'Are you sure you want to delete this entry?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                # Remove the entry from the in-memory list
                del self.overtime_entries[current_row]

                # Save updated data to the file
                self.save_overtime_entries()

                # Update the table and other related variables
                self.update_overtime_table()

    # Replace the existing update_overtime_table() method with this updated version
    def update_overtime_table(self):
        """Update the table widget and refresh other related information."""
        self.overtime_entries.sort(key=lambda e: datetime.strptime(e['date'], '%d-%m-%Y'))  # ← add this
        self.overtime_table.blockSignals(True)  # Block signals to prevent recursive updates
        self.overtime_table.setRowCount(len(self.overtime_entries))
        for row, entry in enumerate(self.overtime_entries):
            # Center-align the "Hours" column
            hours_item = QTableWidgetItem(str(entry['hours']))
            hours_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Center-align text
            self.overtime_table.setItem(row, 0, hours_item)

            # Center-align the "Date" column
            date_item = QTableWidgetItem(entry['date'])
            date_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Center-align text
            self.overtime_table.setItem(row, 1, date_item)

            # Keep the "Task" column left-aligned (default)
            task_item = QTableWidgetItem(entry['task'])
            self.overtime_table.setItem(row, 2, task_item)

        self.overtime_table.blockSignals(False)  # Re-enable signals after updates

        # Update info label whenever the table is updated
        self.update_info_label()

    def generate_report(self):
        try:
            salary = float(self.salary_input.text() or 0)
            now = datetime.now()
            days_in_month = calendar.monthrange(now.year, now.month)[1]
            hourly_rate = (salary / days_in_month) / 8

            # Create a DataFrame with the overtime entries
            report_data = pd.DataFrame(self.overtime_entries)

            # If there are no entries, show a message
            if report_data.empty:
                QMessageBox.information(self, 'No Data', 'There are no overtime entries to generate a report.')
                return

            # Convert the 'date' column to day-month-year format
            report_data['date'] = pd.to_datetime(report_data['date'], dayfirst=True).dt.strftime('%d-%m-%Y')
            report_data.sort_values('date', inplace=True, ignore_index=True)

            # Compute the total hours, days, and amount
            total_hours = report_data['hours'].sum()
            total_days = total_hours / 8
            total_amount = total_hours * hourly_rate

            # Define the filename with the desired day-month-year format
            # Save to specified directory "D:\This PC\Work\nvs\Overtime"
            # file_name = f"D:\\This PC\\Work\\nvs\\Overtime\\overtime-report-{datetime.now().strftime('%d-%m-%Y')}.xlsx"
            report_dir = r"D:\Folder\Work\nvs\Overtime"
            os.makedirs(report_dir, exist_ok=True)  # create folder if it doesn’t exist
            file_name = os.path.join(
                report_dir,
                f"overtime-report-{datetime.now().strftime('%d-%m-%Y')}.xlsx"
            )

            # Use pandas to create a basic Excel file first
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                # Save the DataFrame to Excel without the header
                report_data[['date', 'hours', 'task']].to_excel(writer, index=False, startrow=1, header=False)

                # Access the workbook and sheet
                workbook = writer.book
                sheet = writer.sheets['Sheet1']

                # Set header titles manually
                headers = ['Date', 'Hours', 'Task']
                for col_num, header in enumerate(headers, 1):
                    cell = sheet.cell(row=1, column=col_num)
                    cell.value = header
                    # Set header style
                    cell.font = cell.font.copy(bold=True)
                    cell.alignment = cell.alignment.copy(horizontal='center')

                # Center align all cells and auto adjust column widths
                for column in sheet.columns:
                    max_length = max(len(str(cell.value)) for cell in column) + 2  # Add padding
                    if column[0].column_letter == 'A':  # Make the first column ("Date") slightly wider
                        max_length += 5  # Increase width more for the first column
                    sheet.column_dimensions[column[0].column_letter].width = max_length
                    for cell in column:
                        cell.alignment = cell.alignment.copy(horizontal='center')

                # Write the summary information below the data
                summary_row = len(report_data) + 3
                sheet.cell(row=summary_row, column=1).value = "Total Hours"
                sheet.cell(row=summary_row, column=2).value = total_hours
                sheet.cell(row=summary_row, column=3).value = ""

                sheet.cell(row=summary_row + 1, column=1).value = "In Days"
                sheet.cell(row=summary_row + 1, column=2).value = total_days
                sheet.cell(row=summary_row + 1, column=3).value = ""

                sheet.cell(row=summary_row + 2, column=1).value = "Total Amount"
                sheet.cell(row=summary_row + 2, column=2).value = total_amount
                sheet.cell(row=summary_row + 2, column=3).value = ""

                # Bold and center align all summary cells
                for row in range(summary_row, summary_row + 3):
                    for col in range(1, 4):
                        cell = sheet.cell(row=row, column=col)
                        cell.alignment = cell.alignment.copy(horizontal='center')
                        cell.font = cell.font.copy(bold=True)

            # Inform the user that the report has been saved
            QMessageBox.information(self, 'Report Generated', f'Report saved as {file_name}')

            # Automatically open the generated Excel file
            os.startfile(file_name)

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'An error occurred while generating the report: {str(e)}')

    def get_data_file_path(self, filename):
        """ Helper method to get the correct file path for data files, compatible with PyInstaller """
        try:
            if getattr(sys, 'frozen', False):  # Check if running as a PyInstaller executable
                base_path = os.path.dirname(sys.executable)  # Directory of the executable
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))  # Directory of the script
            return os.path.join(base_path, filename)

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to get file path: {str(e)}')
            return filename  # Fallback in case of an error

    def reset_entries(self):
        try:
            # Show a confirmation dialog
            reply = QMessageBox.question(
                self,
                'Reset Data',
                'This will reset all data including salary and overtime entries. Do you want to proceed?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                # Get paths to data files using the helper function
                salary_file_path = self.get_data_file_path(SALARY_FILE)
                overtime_file_path = self.get_data_file_path(OVERTIME_FILE)

                # Clear the contents of the salary file
                if os.path.exists(salary_file_path):
                    with open(salary_file_path, 'w') as file:
                        json.dump({'salary': 0}, file)
                # Clear the contents of the overtime file
                if os.path.exists(overtime_file_path):
                    with open(overtime_file_path, 'w') as file:
                        json.dump([], file)

                # Reset in-memory data
                self.salary = 0
                self.salary_input.setText('0')
                self.overtime_entries = []

                # Reset overtime rate to x1
                self.overtime_rate_combo.setCurrentIndex(0)  # Set to 'x1'

                # Clear data in the overtime table
                self.overtime_table.clearContents()
                self.overtime_table.setRowCount(0)

                # Reset info label to default values
                self.info_label.setText('Total Hours = 0  |  Days: 0.0 |  Amount = 0.00')

                # Update the UI for rates
                self.update_rates()

                # Inform the user that data has been reset
                QMessageBox.information(self, 'Data Reset', 'All data has been reset successfully.')

        except Exception as e:
            # Handle any unexpected errors and show the error message
            QMessageBox.critical(self, 'Error', f'An error occurred while resetting data: {str(e)}')
            print(f"Error during reset: {e}")  # Debugging statement


class AddEntryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Add Overtime Entry')
        self.setFixedSize(300, 250)

        # Use QSpinBox for Overtime Hours input with up/down arrows
        self.hours_input = QSpinBox()
        self.hours_input.setRange(0, 100)  # Set range as needed
        self.hours_input.setStyleSheet(
            "QSpinBox::up-button { subcontrol-position: top right; width: 16px; }"
            "QSpinBox::down-button { subcontrol-position: bottom right; width: 16px; }"
        )  # Style to place arrows on top of each other
        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDate(datetime.now())
        self.date_input.setDisplayFormat('dd-MM-yyyy')
        self.task_input = QTextEdit()

        # Layouts for aligned input fields
        layout = QVBoxLayout()

        # Overtime Hours Field with Label
        hours_layout = QHBoxLayout()
        hours_layout.addWidget(QLabel('Overtime Hours:'))
        hours_layout.addWidget(self.hours_input)
        layout.addLayout(hours_layout)

        # Date Field with Label
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel('Date:'))
        date_layout.addWidget(self.date_input)
        layout.addLayout(date_layout)

        # Task Description Field
        layout.addWidget(QLabel('Task Description:'))
        layout.addWidget(self.task_input)

        # Add and Cancel Buttons
        button_layout = QHBoxLayout()
        add_button = QPushButton('Add')
        add_button.clicked.connect(self.add_entry)
        cancel_button = QPushButton('Cancel')
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(add_button)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def add_entry(self):
        hours = self.hours_input.value()  # Use .value() with QSpinBox
        task = self.task_input.toPlainText().strip()

        if not hours or not task:
            QMessageBox.warning(self, 'Invalid Input', 'Please fill all fields correctly.')
            return

        self.accept()

    def get_entry(self):
        return {
            'hours': float(self.hours_input.value()),  # Use .value() with QSpinBox
            'date': self.date_input.date().toString('dd-MM-yyyy'),
            'task': self.task_input.toPlainText()
        }

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OvertimeTrackerApp()
    window.show()
    sys.exit(app.exec())
