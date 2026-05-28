import sys
import os
import json
import shutil
import sqlite3
from datetime import datetime
from openpyxl import Workbook

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import matplotlib.pyplot as plt

from PySide6.QtCore import QDate
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QListWidget, QLineEdit, QMessageBox, QComboBox, QTableWidget,
    QTableWidgetItem, QFileDialog, QDateEdit, QCheckBox, QTabWidget
)


DB_NAME = "payroll.db"
SETTINGS_FILE = "settings.json"


DEFAULT_SETTINGS = {
    "theme": "dark",
    "view": "tabs"
}


class PayrollApp(QWidget):
    def __init__(self):
        super().__init__()

        self.settings = self.load_settings()
        self.editing_record_id = None

        self.setWindowTitle("Payroll Manager")
        self.resize(1700, 950)

        self.conn = sqlite3.connect(DB_NAME)
        self.cursor = self.conn.cursor()
        self.create_tables()

        self.apply_theme()

        main_layout = QVBoxLayout()

        if self.settings["view"] == "classic":
            main_layout.addLayout(self.build_classic_view())
        else:
            main_layout.addWidget(self.build_tabs_view())

        self.setLayout(main_layout)

        self.load_employees()
        self.load_history()

    def load_settings(self):
        if not os.path.exists(SETTINGS_FILE):
            return DEFAULT_SETTINGS.copy()

        try:
            with open(SETTINGS_FILE, "r") as file:
                return json.load(file)
        except Exception:
            return DEFAULT_SETTINGS.copy()

    def save_settings(self):
        with open(SETTINGS_FILE, "w") as file:
            json.dump(self.settings, file, indent=4)

    def apply_theme(self):
        if self.settings["theme"] == "light":
            self.setStyleSheet("""
                QWidget { background-color: #f5f5f5; color: #111; font-size: 14px; }
                QLineEdit, QComboBox, QDateEdit {
                    background-color: #ffffff; color: #111; border: 1px solid #bbb;
                    padding: 6px; border-radius: 6px;
                }
                QPushButton {
                    background-color: #2d6cdf; color: white; padding: 8px;
                    border-radius: 8px; font-weight: bold;
                }
                QPushButton:hover { background-color: #3f7df0; }
                QTableWidget { background-color: #ffffff; color: #111; gridline-color: #ccc; }
                QHeaderView::section {
                    background-color: #e0e0e0; color: #111; padding: 6px; border: 1px solid #ccc;
                }
                QListWidget { background-color: #ffffff; border: 1px solid #bbb; }
            """)
        else:
            self.setStyleSheet("""
                QWidget { background-color: #121212; color: #eeeeee; font-size: 14px; }
                QLineEdit, QComboBox, QDateEdit {
                    background-color: #1f1f1f; color: #ffffff; border: 1px solid #444;
                    padding: 6px; border-radius: 6px;
                }
                QPushButton {
                    background-color: #2d6cdf; color: white; padding: 8px;
                    border-radius: 8px; font-weight: bold;
                }
                QPushButton:hover { background-color: #3f7df0; }
                QTableWidget { background-color: #1a1a1a; color: #ffffff; gridline-color: #444; }
                QHeaderView::section {
                    background-color: #2a2a2a; color: white; padding: 6px; border: 1px solid #444;
                }
                QListWidget { background-color: #1f1f1f; border: 1px solid #444; }
                QTabBar::tab { background: #222; color: white; padding: 10px; border-radius: 6px; }
                QTabBar::tab:selected { background: #2d6cdf; }
            """)

    def build_tabs_view(self):
        tabs = QTabWidget()

        self.employees_tab = QWidget()
        self.attendance_tab = QWidget()
        self.payroll_tab = QWidget()
        self.reports_tab = QWidget()
        self.settings_tab = QWidget()

        tabs.addTab(self.employees_tab, "Employees")
        tabs.addTab(self.attendance_tab, "Attendance")
        tabs.addTab(self.payroll_tab, "Payroll")
        tabs.addTab(self.reports_tab, "Reports")
        tabs.addTab(self.settings_tab, "Settings")

        self.build_employees_ui(self.employees_tab)
        self.build_attendance_ui(self.attendance_tab)
        self.build_payroll_ui(self.payroll_tab)
        self.build_reports_ui(self.reports_tab)
        self.build_settings_ui(self.settings_tab)

        return tabs

    def build_classic_view(self):
        main = QHBoxLayout()

        left = QVBoxLayout()
        middle = QVBoxLayout()
        right = QVBoxLayout()

        employee_box = QWidget()
        attendance_box = QWidget()
        payroll_box = QWidget()
        reports_box = QWidget()
        settings_box = QWidget()

        self.build_employees_ui(employee_box)
        self.build_attendance_ui(attendance_box)
        self.build_payroll_ui(payroll_box)
        self.build_reports_ui(reports_box)
        self.build_settings_ui(settings_box)

        left.addWidget(employee_box)
        left.addWidget(attendance_box)

        middle.addWidget(payroll_box)
        middle.addWidget(settings_box)

        right.addWidget(reports_box)

        main.addLayout(left, 1)
        main.addLayout(middle, 1)
        main.addLayout(right, 2)

        return main

    def build_employees_ui(self, parent):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Διαχείριση Υπαλλήλων"))

        self.employee_input = QLineEdit()
        self.employee_input.setPlaceholderText("Όνομα υπαλλήλου")
        layout.addWidget(self.employee_input)

        add_btn = QPushButton("Προσθήκη Υπαλλήλου")
        add_btn.clicked.connect(self.add_employee)
        layout.addWidget(add_btn)

        self.employee_list = QListWidget()
        layout.addWidget(self.employee_list)

        remove_btn = QPushButton("Αφαίρεση Υπαλλήλου")
        remove_btn.clicked.connect(self.remove_employee)
        layout.addWidget(remove_btn)

        parent.setLayout(layout)

    def build_attendance_ui(self, parent):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Καταχώρηση Παρουσίας"))

        self.employee_select = QComboBox()
        layout.addWidget(self.employee_select)

        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDate(QDate.currentDate())
        layout.addWidget(self.date_input)

        self.check_in_input = QLineEdit()
        self.check_in_input.setPlaceholderText("Προσέλευση π.χ. 9 ή 09:00")
        layout.addWidget(self.check_in_input)

        self.check_out_input = QLineEdit()
        self.check_out_input.setPlaceholderText("Αποχώρηση π.χ. 18 ή 18:30")
        layout.addWidget(self.check_out_input)

        self.break_input = QLineEdit()
        self.break_input.setPlaceholderText("Διάλειμμα σε λεπτά")
        layout.addWidget(self.break_input)

        self.paid_hours_input = QLineEdit()
        self.paid_hours_input.setPlaceholderText("Πληρωμένες ώρες")
        layout.addWidget(self.paid_hours_input)

        self.unpaid_overtime_amount_input = QLineEdit()
        self.unpaid_overtime_amount_input.setPlaceholderText("Ποσό απλήρωτων υπερωριών")
        layout.addWidget(self.unpaid_overtime_amount_input)

        self.unpaid_overtime_paid_checkbox = QCheckBox("Οι υπερωρίες πληρώθηκαν")
        layout.addWidget(self.unpaid_overtime_paid_checkbox)

        layout.addWidget(QLabel("Ημερομηνία πληρωμής υπερωριών"))

        self.payment_date_input = QDateEdit()
        self.payment_date_input.setCalendarPopup(True)
        self.payment_date_input.setDate(QDate.currentDate())
        layout.addWidget(self.payment_date_input)

        self.save_btn = QPushButton("Αποθήκευση Ωρών")
        self.save_btn.clicked.connect(self.save_attendance)
        layout.addWidget(self.save_btn)

        cancel_btn = QPushButton("Ακύρωση Επεξεργασίας")
        cancel_btn.clicked.connect(self.cancel_edit)
        layout.addWidget(cancel_btn)

        self.result_label = QLabel("")
        layout.addWidget(self.result_label)

        parent.setLayout(layout)

    def build_payroll_ui(self, parent):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Φίλτρα Payroll"))

        self.employee_search_input = QLineEdit()
        self.employee_search_input.setPlaceholderText("Search υπαλλήλου")
        self.employee_search_input.textChanged.connect(self.load_history)
        layout.addWidget(self.employee_search_input)

        self.filter_employee_select = QComboBox()
        layout.addWidget(self.filter_employee_select)

        self.overtime_payment_filter = QComboBox()
        self.overtime_payment_filter.addItem("Όλες οι υπερωρίες", "all")
        self.overtime_payment_filter.addItem("Μόνο πληρωμένες υπερωρίες", "paid")
        self.overtime_payment_filter.addItem("Μόνο απλήρωτες υπερωρίες", "unpaid")
        self.overtime_payment_filter.currentIndexChanged.connect(self.load_history)
        layout.addWidget(self.overtime_payment_filter)

        layout.addWidget(QLabel("Από ημερομηνία"))
        self.from_date_input = QDateEdit()
        self.from_date_input.setCalendarPopup(True)
        self.from_date_input.setDate(QDate.currentDate().addMonths(-1))
        self.from_date_input.dateChanged.connect(self.load_history)
        layout.addWidget(self.from_date_input)

        layout.addWidget(QLabel("Έως ημερομηνία"))
        self.to_date_input = QDateEdit()
        self.to_date_input.setCalendarPopup(True)
        self.to_date_input.setDate(QDate.currentDate())
        self.to_date_input.dateChanged.connect(self.load_history)
        layout.addWidget(self.to_date_input)

        clear_btn = QPushButton("Καθαρισμός Φίλτρων")
        clear_btn.clicked.connect(self.clear_filters)
        layout.addWidget(clear_btn)

        self.summary_label = QLabel("")
        layout.addWidget(self.summary_label)

        self.employee_summary_table = QTableWidget()
        self.employee_summary_table.setColumnCount(6)
        self.employee_summary_table.setHorizontalHeaderLabels([
            "Υπάλληλος", "Σύνολο Ωρών", "Πληρωμένες",
            "Απλήρωτες", "Υπερωρίες", "Οφειλόμενο Ποσό"
        ])
        layout.addWidget(self.employee_summary_table)

        parent.setLayout(layout)

    def build_reports_ui(self, parent):
        layout = QVBoxLayout()

        self.history_table = QTableWidget()
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.setSelectionMode(QTableWidget.SingleSelection)
        self.history_table.setColumnCount(12)
        self.history_table.setHorizontalHeaderLabels([
            "ID", "Υπάλληλος", "Ημερομηνία", "Προσέλευση", "Αποχώρηση",
            "Σύνολο", "Πληρωμένες", "Απλήρωτες", "Υπερωρίες",
            "Ποσό Υπερωριών", "Πληρώθηκε;", "Ημ/νία Πληρωμής"
        ])
        layout.addWidget(self.history_table)

        buttons = QHBoxLayout()

        for text, action in [
            ("Επεξεργασία", self.load_selected_record_for_edit),
            ("Διαγραφή", self.delete_selected_record),
            ("Σήμανση Πληρωμένη", self.mark_selected_as_paid),
            ("Σήμανση Απλήρωτη", self.mark_selected_as_unpaid),
            ("Export Excel", self.export_to_excel),
            ("Export PDF", self.export_to_pdf),
            ("Backup Database", self.backup_database),
            ("Charts", self.create_charts),
        ]:
            btn = QPushButton(text)
            btn.clicked.connect(action)
            buttons.addWidget(btn)

        layout.addLayout(buttons)
        parent.setLayout(layout)

    def build_settings_ui(self, parent):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Ρυθμίσεις Εμφάνισης"))

        layout.addWidget(QLabel("Theme"))

        self.theme_select = QComboBox()
        self.theme_select.addItem("Dark Mode", "dark")
        self.theme_select.addItem("Light Mode", "light")
        self.theme_select.setCurrentIndex(0 if self.settings["theme"] == "dark" else 1)
        layout.addWidget(self.theme_select)

        layout.addWidget(QLabel("View"))

        self.view_select = QComboBox()
        self.view_select.addItem("Modern Tabs View", "tabs")
        self.view_select.addItem("Classic Single Window View", "classic")
        self.view_select.setCurrentIndex(0 if self.settings["view"] == "tabs" else 1)
        layout.addWidget(self.view_select)

        save_btn = QPushButton("Αποθήκευση Ρυθμίσεων")
        save_btn.clicked.connect(self.save_ui_settings)
        layout.addWidget(save_btn)

        layout.addWidget(QLabel("Σημείωση: οι αλλαγές σε theme/view εφαρμόζονται αφού κλείσεις και ξανανοίξεις την εφαρμογή."))

        parent.setLayout(layout)

    def save_ui_settings(self):
        self.settings["theme"] = self.theme_select.currentData()
        self.settings["view"] = self.view_select.currentData()
        self.save_settings()

        QMessageBox.information(
            self,
            "Αποθηκεύτηκε",
            "Οι ρυθμίσεις αποθηκεύτηκαν. Κλείσε και ξανάνοιξε την εφαρμογή."
        )

    def create_tables(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL
            )
        """)

        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER,
                date TEXT,
                check_in TEXT,
                check_out TEXT,
                break_minutes INTEGER,
                total_hours REAL,
                paid_hours REAL,
                unpaid_hours REAL,
                overtime_hours REAL,
                unpaid_overtime_amount REAL DEFAULT 0,
                unpaid_overtime_paid INTEGER DEFAULT 0,
                overtime_payment_date TEXT
            )
        """)

        self.add_column_if_missing("attendance", "unpaid_overtime_amount", "REAL DEFAULT 0")
        self.add_column_if_missing("attendance", "unpaid_overtime_paid", "INTEGER DEFAULT 0")
        self.add_column_if_missing("attendance", "overtime_payment_date", "TEXT")

        self.conn.commit()

    def add_column_if_missing(self, table_name, column_name, column_type):
        self.cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [column[1] for column in self.cursor.fetchall()]

        if column_name not in columns:
            self.cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")

    def normalize_time(self, value):
        value = value.strip()
        if ":" in value:
            return value
        if value.isdigit():
            return f"{int(value):02d}:00"
        raise ValueError

    def load_employees(self):
        current_filter_id = self.filter_employee_select.currentData() if self.filter_employee_select.count() > 0 else None

        self.employee_list.clear()
        self.employee_select.clear()

        self.filter_employee_select.blockSignals(True)
        self.filter_employee_select.clear()
        self.filter_employee_select.addItem("Όλοι οι υπάλληλοι", None)

        self.cursor.execute("SELECT id, name FROM employees ORDER BY name")
        employees = self.cursor.fetchall()

        selected_index = 0

        for employee_id, name in employees:
            self.employee_list.addItem(name)
            self.employee_select.addItem(name, employee_id)
            self.filter_employee_select.addItem(name, employee_id)

            if current_filter_id == employee_id:
                selected_index = self.filter_employee_select.count() - 1

        self.filter_employee_select.setCurrentIndex(selected_index)
        self.filter_employee_select.blockSignals(False)
        self.filter_employee_select.currentIndexChanged.connect(self.load_history)

    def add_employee(self):
        name = self.employee_input.text().strip()

        if name == "":
            QMessageBox.warning(self, "Σφάλμα", "Δώσε όνομα υπαλλήλου")
            return

        self.cursor.execute("INSERT INTO employees (name) VALUES (?)", (name,))
        self.conn.commit()

        self.employee_input.clear()
        self.load_employees()
        self.load_history()

    def remove_employee(self):
        selected_item = self.employee_list.currentItem()

        if selected_item is None:
            QMessageBox.warning(self, "Σφάλμα", "Επίλεξε υπάλληλο")
            return

        name = selected_item.text()

        confirm = QMessageBox.question(self, "Επιβεβαίωση", f"Να αφαιρεθεί ο υπάλληλος {name};")

        if confirm != QMessageBox.Yes:
            return

        self.cursor.execute("DELETE FROM employees WHERE name = ?", (name,))
        self.conn.commit()

        self.load_employees()
        self.load_history()

    def calculate_attendance_values(self):
        check_in_text = self.check_in_input.text().strip()
        check_out_text = self.check_out_input.text().strip()
        break_text = self.break_input.text().strip()
        paid_text = self.paid_hours_input.text().strip()
        amount_text = self.unpaid_overtime_amount_input.text().strip()

        if check_in_text == "" or check_out_text == "":
            raise ValueError

        check_in = self.normalize_time(check_in_text)
        check_out = self.normalize_time(check_out_text)

        break_minutes = int(break_text) if break_text else 0
        paid_hours = float(paid_text) if paid_text else 0
        unpaid_overtime_amount = float(amount_text) if amount_text else 0
        unpaid_overtime_paid = 1 if self.unpaid_overtime_paid_checkbox.isChecked() else 0

        overtime_payment_date = None
        if unpaid_overtime_paid == 1:
            overtime_payment_date = self.payment_date_input.date().toString("yyyy-MM-dd")

        start = datetime.strptime(check_in, "%H:%M")
        end = datetime.strptime(check_out, "%H:%M")

        actual_hours = ((end - start).total_seconds() / 60 - break_minutes) / 60

        if actual_hours < 0:
            raise ValueError

        total_hours = max(actual_hours, 8)
        overtime_hours = max(0, total_hours - 8)
        unpaid_hours = max(0, total_hours - paid_hours)

        return {
            "check_in": check_in,
            "check_out": check_out,
            "break_minutes": break_minutes,
            "paid_hours": paid_hours,
            "unpaid_hours": unpaid_hours,
            "total_hours": total_hours,
            "overtime_hours": overtime_hours,
            "unpaid_overtime_amount": unpaid_overtime_amount,
            "unpaid_overtime_paid": unpaid_overtime_paid,
            "overtime_payment_date": overtime_payment_date
        }

    def save_attendance(self):
        employee_id = self.employee_select.currentData()
        date = self.date_input.date().toString("yyyy-MM-dd")

        if employee_id is None:
            QMessageBox.warning(self, "Σφάλμα", "Δεν υπάρχει υπάλληλος")
            return

        try:
            values = self.calculate_attendance_values()
        except ValueError:
            QMessageBox.warning(self, "Σφάλμα", "Έλεγξε ώρες και αριθμούς")
            return

        if self.editing_record_id is None:
            self.cursor.execute("""
                INSERT INTO attendance (
                    employee_id, date, check_in, check_out, break_minutes,
                    total_hours, paid_hours, unpaid_hours, overtime_hours,
                    unpaid_overtime_amount, unpaid_overtime_paid,
                    overtime_payment_date
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                employee_id, date, values["check_in"], values["check_out"],
                values["break_minutes"], values["total_hours"], values["paid_hours"],
                values["unpaid_hours"], values["overtime_hours"],
                values["unpaid_overtime_amount"], values["unpaid_overtime_paid"],
                values["overtime_payment_date"]
            ))
            message = "Αποθηκεύτηκε"
        else:
            self.cursor.execute("""
                UPDATE attendance
                SET employee_id = ?, date = ?, check_in = ?, check_out = ?,
                    break_minutes = ?, total_hours = ?, paid_hours = ?,
                    unpaid_hours = ?, overtime_hours = ?,
                    unpaid_overtime_amount = ?, unpaid_overtime_paid = ?,
                    overtime_payment_date = ?
                WHERE id = ?
            """, (
                employee_id, date, values["check_in"], values["check_out"],
                values["break_minutes"], values["total_hours"], values["paid_hours"],
                values["unpaid_hours"], values["overtime_hours"],
                values["unpaid_overtime_amount"], values["unpaid_overtime_paid"],
                values["overtime_payment_date"], self.editing_record_id
            ))
            message = "Ενημερώθηκε"

        self.conn.commit()
        self.result_label.setText(f"{message}: {values['total_hours']:.2f} ώρες")
        self.clear_attendance_form()
        self.load_history()

    def clear_attendance_form(self):
        self.editing_record_id = None
        self.save_btn.setText("Αποθήκευση Ωρών")
        self.check_in_input.clear()
        self.check_out_input.clear()
        self.break_input.clear()
        self.paid_hours_input.clear()
        self.unpaid_overtime_amount_input.clear()
        self.unpaid_overtime_paid_checkbox.setChecked(False)
        self.payment_date_input.setDate(QDate.currentDate())

    def cancel_edit(self):
        self.clear_attendance_form()
        self.result_label.setText("Η επεξεργασία ακυρώθηκε")

    def clear_filters(self):
        self.employee_search_input.clear()
        self.overtime_payment_filter.setCurrentIndex(0)
        self.filter_employee_select.setCurrentIndex(0)
        self.from_date_input.setDate(QDate.currentDate().addMonths(-1))
        self.to_date_input.setDate(QDate.currentDate())
        self.load_history()

    def get_selected_record_id(self):
        selected_row = self.history_table.currentRow()

        if selected_row < 0:
            QMessageBox.warning(self, "Σφάλμα", "Επίλεξε μία καταχώρηση")
            return None

        item = self.history_table.item(selected_row, 0)

        if item is None:
            return None

        return int(item.text())

    def load_selected_record_for_edit(self):
        record_id = self.get_selected_record_id()
        if record_id is None:
            return

        self.cursor.execute("""
            SELECT id, employee_id, date, check_in, check_out, break_minutes,
                   paid_hours, unpaid_overtime_amount, unpaid_overtime_paid,
                   overtime_payment_date
            FROM attendance
            WHERE id = ?
        """, (record_id,))

        record = self.cursor.fetchone()

        if record is None:
            return

        (
            attendance_id, employee_id, date, check_in, check_out,
            break_minutes, paid_hours, unpaid_overtime_amount,
            unpaid_overtime_paid, overtime_payment_date
        ) = record

        self.editing_record_id = attendance_id

        index = self.employee_select.findData(employee_id)
        if index >= 0:
            self.employee_select.setCurrentIndex(index)

        parsed_date = QDate.fromString(date, "yyyy-MM-dd")
        if parsed_date.isValid():
            self.date_input.setDate(parsed_date)

        self.check_in_input.setText(check_in)
        self.check_out_input.setText(check_out)
        self.break_input.setText(str(break_minutes or 0))
        self.paid_hours_input.setText(str(paid_hours or 0))
        self.unpaid_overtime_amount_input.setText(str(unpaid_overtime_amount or 0))
        self.unpaid_overtime_paid_checkbox.setChecked(int(unpaid_overtime_paid or 0) == 1)

        parsed_payment_date = QDate.fromString(str(overtime_payment_date), "yyyy-MM-dd")
        if parsed_payment_date.isValid():
            self.payment_date_input.setDate(parsed_payment_date)

        self.save_btn.setText("Ενημέρωση Καταχώρησης")
        self.result_label.setText(f"Επεξεργασία ID {attendance_id}")

    def delete_selected_record(self):
        record_id = self.get_selected_record_id()
        if record_id is None:
            return

        confirm = QMessageBox.question(self, "Επιβεβαίωση", f"Να διαγραφεί το ID {record_id};")
        if confirm != QMessageBox.Yes:
            return

        self.cursor.execute("DELETE FROM attendance WHERE id = ?", (record_id,))
        self.conn.commit()
        self.load_history()

    def mark_selected_as_paid(self):
        record_id = self.get_selected_record_id()
        if record_id is None:
            return

        payment_date = QDate.currentDate().toString("yyyy-MM-dd")

        self.cursor.execute("""
            UPDATE attendance
            SET unpaid_overtime_paid = 1,
                overtime_payment_date = ?
            WHERE id = ?
        """, (payment_date, record_id))

        self.conn.commit()
        self.load_history()

    def mark_selected_as_unpaid(self):
        record_id = self.get_selected_record_id()
        if record_id is None:
            return

        self.cursor.execute("""
            UPDATE attendance
            SET unpaid_overtime_paid = 0,
                overtime_payment_date = NULL
            WHERE id = ?
        """, (record_id,))

        self.conn.commit()
        self.load_history()

    def get_filtered_records(self):
        employee_id = self.filter_employee_select.currentData()
        search = self.employee_search_input.text().strip()
        payment_filter = self.overtime_payment_filter.currentData()

        from_date = self.from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.to_date_input.date().toString("yyyy-MM-dd")

        query = """
            SELECT attendance.id, employees.name, attendance.date,
                   attendance.check_in, attendance.check_out,
                   attendance.total_hours, attendance.paid_hours,
                   attendance.unpaid_hours, attendance.overtime_hours,
                   attendance.unpaid_overtime_amount,
                   attendance.unpaid_overtime_paid,
                   attendance.overtime_payment_date
            FROM attendance
            JOIN employees ON attendance.employee_id = employees.id
            WHERE attendance.date >= ?
            AND attendance.date <= ?
        """

        params = [from_date, to_date]

        if employee_id is not None:
            query += " AND attendance.employee_id = ?"
            params.append(employee_id)

        if search:
            query += " AND employees.name LIKE ?"
            params.append(f"%{search}%")

        if payment_filter == "paid":
            query += " AND attendance.overtime_hours > 0 AND attendance.unpaid_overtime_paid = 1"

        if payment_filter == "unpaid":
            query += " AND attendance.overtime_hours > 0 AND attendance.unpaid_overtime_paid = 0"

        query += " ORDER BY attendance.date DESC, attendance.id DESC"

        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def load_history(self):
        records = self.get_filtered_records()

        self.history_table.clearContents()
        self.history_table.setRowCount(len(records))

        totals = {}
        total_hours_sum = 0
        paid_hours_sum = 0
        unpaid_hours_sum = 0
        overtime_hours_sum = 0
        overtime_amount_sum = 0
        unpaid_balance_sum = 0

        for row, record in enumerate(records):
            employee = record[1]
            total_hours = float(record[5] or 0)
            paid_hours = float(record[6] or 0)
            unpaid_hours = float(record[7] or 0)
            overtime_hours = float(record[8] or 0)
            amount = float(record[9] or 0)
            paid_status = int(record[10] or 0)

            total_hours_sum += total_hours
            paid_hours_sum += paid_hours
            unpaid_hours_sum += unpaid_hours
            overtime_hours_sum += overtime_hours
            overtime_amount_sum += amount

            if paid_status == 0:
                unpaid_balance_sum += amount

            if employee not in totals:
                totals[employee] = [0, 0, 0, 0, 0]

            totals[employee][0] += total_hours
            totals[employee][1] += paid_hours
            totals[employee][2] += unpaid_hours
            totals[employee][3] += overtime_hours

            if paid_status == 0:
                totals[employee][4] += amount

            display = list(record)
            display[10] = "Ναι" if paid_status == 1 else "Όχι"
            display[11] = display[11] if display[11] else "-"

            for col, value in enumerate(display):
                item = QTableWidgetItem(str(value))

                if col == 10:
                    item.setBackground(QColor("#1f7a1f") if paid_status == 1 else QColor("#8a1f1f"))

                if col == 9 and paid_status == 0 and amount > 0:
                    item.setBackground(QColor("#8a1f1f"))

                self.history_table.setItem(row, col, item)

        self.history_table.resizeColumnsToContents()

        self.summary_label.setText(
            f"Σύνολα Περιόδου\n\n"
            f"Σύνολο Ωρών: {total_hours_sum:.2f}\n"
            f"Πληρωμένες: {paid_hours_sum:.2f}\n"
            f"Απλήρωτες: {unpaid_hours_sum:.2f}\n"
            f"Υπερωρίες: {overtime_hours_sum:.2f}\n\n"
            f"Ποσά Υπερωριών: {overtime_amount_sum:.2f} €\n"
            f"Οφειλόμενα: {unpaid_balance_sum:.2f} €"
        )

        self.employee_summary_table.clearContents()
        self.employee_summary_table.setRowCount(len(totals))

        for row, (employee, values) in enumerate(totals.items()):
            row_values = [
                employee,
                f"{values[0]:.2f}",
                f"{values[1]:.2f}",
                f"{values[2]:.2f}",
                f"{values[3]:.2f}",
                f"{values[4]:.2f} €"
            ]

            for col, value in enumerate(row_values):
                item = QTableWidgetItem(value)
                if col == 5 and values[4] > 0:
                    item.setBackground(QColor("#8a1f1f"))
                self.employee_summary_table.setItem(row, col, item)

        self.employee_summary_table.resizeColumnsToContents()

    def export_to_excel(self):
        records = self.get_filtered_records()

        if not records:
            QMessageBox.warning(self, "Σφάλμα", "Δεν υπάρχουν δεδομένα")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Αποθήκευση Excel", "payroll_export.xlsx", "Excel Files (*.xlsx)"
        )

        if not file_path:
            return

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Payroll"

        headers = [
            "ID", "Υπάλληλος", "Ημερομηνία", "Προσέλευση", "Αποχώρηση",
            "Σύνολο", "Πληρωμένες", "Απλήρωτες", "Υπερωρίες",
            "Ποσό Υπερωριών", "Πληρώθηκε;", "Ημ/νία Πληρωμής"
        ]

        sheet.append(headers)

        for record in records:
            row = list(record)
            row[10] = "Ναι" if int(record[10] or 0) == 1 else "Όχι"
            row[11] = row[11] if row[11] else "-"
            sheet.append(row)

        workbook.save(file_path)
        QMessageBox.information(self, "Ολοκληρώθηκε", "Το Excel αποθηκεύτηκε")

    def export_to_pdf(self):
        records = self.get_filtered_records()

        if not records:
            QMessageBox.warning(self, "Σφάλμα", "Δεν υπάρχουν δεδομένα")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Αποθήκευση PDF", "payroll_report.pdf", "PDF Files (*.pdf)"
        )

        if not file_path:
            return

        pdf = canvas.Canvas(file_path, pagesize=A4)
        width, height = A4

        y = height - 40
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawString(40, y, "Payroll Report")

        y -= 30
        pdf.setFont("Helvetica", 8)

        for record in records:
            line = (
                f"ID:{record[0]} | {record[1]} | {record[2]} | "
                f"{record[3]}-{record[4]} | Hours:{record[5]} | "
                f"OT:{record[8]} | Amount:{record[9]} | "
                f"Paid:{'Yes' if int(record[10] or 0) == 1 else 'No'}"
            )

            pdf.drawString(40, y, line)
            y -= 14

            if y < 40:
                pdf.showPage()
                y = height - 40
                pdf.setFont("Helvetica", 8)

        pdf.save()
        QMessageBox.information(self, "Ολοκληρώθηκε", "Το PDF αποθηκεύτηκε")

    def backup_database(self):
        if not os.path.exists(DB_NAME):
            QMessageBox.warning(self, "Σφάλμα", "Δεν βρέθηκε βάση δεδομένων")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Αποθήκευση Backup",
            f"payroll_backup_{timestamp}.db",
            "Database Files (*.db)"
        )

        if not file_path:
            return

        shutil.copy(DB_NAME, file_path)
        QMessageBox.information(self, "Backup", "Το backup ολοκληρώθηκε")

    def create_charts(self):
        records = self.get_filtered_records()

        if not records:
            QMessageBox.warning(self, "Σφάλμα", "Δεν υπάρχουν δεδομένα")
            return

        totals = {}

        for record in records:
            employee = record[1]
            amount = float(record[9] or 0)
            paid_status = int(record[10] or 0)

            if employee not in totals:
                totals[employee] = 0

            if paid_status == 0:
                totals[employee] += amount

        names = list(totals.keys())
        amounts = list(totals.values())

        plt.figure(figsize=(10, 6))
        plt.bar(names, amounts)
        plt.title("Οφειλόμενα ποσά υπερωριών ανά υπάλληλο")
        plt.xlabel("Υπάλληλος")
        plt.ylabel("Ποσό (€)")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()


app = QApplication(sys.argv)

window = PayrollApp()
window.show()

sys.exit(app.exec())