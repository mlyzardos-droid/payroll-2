import sys
import os
import json
import shutil
import sqlite3
import re
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


if sys.platform == "win32":
    APP_DATA_ROOT = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
else:
    APP_DATA_ROOT = os.path.join(os.path.expanduser("~"), "Library", "Application Support")

APP_DATA_DIR = os.path.join(APP_DATA_ROOT, "PrometheusPayroll")

os.makedirs(APP_DATA_DIR, exist_ok=True)

DB_NAME = os.path.join(APP_DATA_DIR, "payroll.db")
SETTINGS_FILE = os.path.join(APP_DATA_DIR, "settings.json")


def migrate_legacy_user_files():
    """Copy legacy working-directory data only when the new user database is absent."""
    legacy_db = os.path.join(os.path.dirname(os.path.abspath(__file__)), "payroll.db")
    if not os.path.exists(DB_NAME) and os.path.isfile(legacy_db) and os.path.abspath(legacy_db) != os.path.abspath(DB_NAME):
        shutil.copy2(legacy_db, DB_NAME)
    legacy_settings = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
    if not os.path.exists(SETTINGS_FILE) and os.path.isfile(legacy_settings) and os.path.abspath(legacy_settings) != os.path.abspath(SETTINGS_FILE):
        shutil.copy2(legacy_settings, SETTINGS_FILE)


migrate_legacy_user_files()


DEFAULT_SETTINGS = {
    "theme": "dark",
    "view": "tabs"
}


class PayrollApp(QWidget):
    def __init__(self):
        super().__init__()

        self.settings = self.load_settings()
        self.editing_record_id = None
        self.last_deleted_employee = None
        self.last_edit_snapshot = None

        self.setWindowTitle("Prometheus Payroll v1.1.0 — Hourly Payroll Edition")
        self.resize(1700, 950)

        self.conn = sqlite3.connect(DB_NAME)
        self.cursor = self.conn.cursor()
        self.migration_summary = self.create_tables()

        self.apply_theme()

        main_layout = QVBoxLayout()

        if self.settings["view"] == "classic":
            main_layout.addLayout(self.build_classic_view())
        else:
            main_layout.addWidget(self.build_tabs_view())

        self.setLayout(main_layout)

        self.load_employees()
        self.load_history()
        self.load_monthly_summary()
        if self.migration_summary:
            QMessageBox.information(self, "Ολοκληρώθηκε η ενημέρωση", self.migration_summary)

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
        self.tabs = tabs

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

        self.employee_hourly_rate_input = QLineEdit()
        self.employee_hourly_rate_input.setPlaceholderText("Βασικό ωρομίσθιο (€), π.χ. 7 ή 7,50")
        layout.addWidget(self.employee_hourly_rate_input)

        add_btn = QPushButton("Προσθήκη Υπαλλήλου")
        add_btn.clicked.connect(self.add_employee)
        layout.addWidget(add_btn)

        self.employee_rate_save_btn = QPushButton("Επεξεργασία / Αποθήκευση Ωρομισθίου")
        self.employee_rate_save_btn.clicked.connect(self.save_employee_hourly_rate)
        layout.addWidget(self.employee_rate_save_btn)

        self.employee_list = QListWidget()
        self.employee_list.currentItemChanged.connect(self.load_selected_employee_rate)
        layout.addWidget(self.employee_list)

        remove_btn = QPushButton("Αφαίρεση Υπαλλήλου")
        remove_btn.clicked.connect(self.remove_employee)
        layout.addWidget(remove_btn)

        undo_remove_btn = QPushButton("Αναίρεση Διαγραφής Υπαλλήλου")
        undo_remove_btn.clicked.connect(self.undo_delete_employee)
        layout.addWidget(undo_remove_btn)

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

        layout.addWidget(QLabel("Το διάλειμμα αφαιρείται από τις πραγματικές ώρες εργασίας."))

        self.save_btn = QPushButton("Αποθήκευση Ωρών")
        self.save_btn.clicked.connect(self.save_attendance)
        layout.addWidget(self.save_btn)

        cancel_btn = QPushButton("Ακύρωση Επεξεργασίας")
        cancel_btn.clicked.connect(self.cancel_edit)
        layout.addWidget(cancel_btn)

        undo_edit_btn = QPushButton("Αναίρεση Τελευταίας Αλλαγής Μισθοδοσίας")
        undo_edit_btn.clicked.connect(self.undo_last_edit)
        layout.addWidget(undo_edit_btn)

        self.result_label = QLabel("")
        layout.addWidget(self.result_label)

        parent.setLayout(layout)

    def build_payroll_ui(self, parent):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Φίλτρα Payroll"))

        self.employee_search_input = QLineEdit()
        self.employee_search_input.setPlaceholderText("Αναζήτηση υπαλλήλου")
        self.employee_search_input.textChanged.connect(self.load_history)
        layout.addWidget(self.employee_search_input)

        self.filter_employee_select = QComboBox()
        self.filter_employee_select.currentIndexChanged.connect(self.sync_monthly_from_payroll_filters)
        layout.addWidget(self.filter_employee_select)

        layout.addWidget(QLabel("Από ημερομηνία"))
        self.from_date_input = QDateEdit()
        self.from_date_input.setCalendarPopup(True)
        self.from_date_input.setDate(QDate.currentDate().addMonths(-1))
        self.from_date_input.dateChanged.connect(self.sync_monthly_from_payroll_filters)
        layout.addWidget(self.from_date_input)

        layout.addWidget(QLabel("Έως ημερομηνία"))
        self.to_date_input = QDateEdit()
        self.to_date_input.setCalendarPopup(True)
        self.to_date_input.setDate(QDate.currentDate())
        self.to_date_input.dateChanged.connect(self.sync_monthly_from_payroll_filters)
        layout.addWidget(self.to_date_input)

        clear_btn = QPushButton("Καθαρισμός Φίλτρων")
        clear_btn.clicked.connect(self.clear_filters)
        layout.addWidget(clear_btn)

        self.summary_label = QLabel("")
        layout.addWidget(self.summary_label)

        self.employee_summary_table = QTableWidget()
        self.employee_summary_table.setColumnCount(7)
        self.employee_summary_table.setHorizontalHeaderLabels([
            "Υπάλληλος", "Πραγματικές Ώρες", "Πληρωμένες Ώρες",
            "Υπόλοιπο Ωρών", "Σύνολο Μισθοδοσίας", "Πληρωμένο Ποσό", "Οφειλόμενο Ποσό"
        ])
        layout.addWidget(self.employee_summary_table)

        layout.addWidget(QLabel("Μηνιαία μισθοδοσία"))
        self.monthly_employee_select = QComboBox()
        layout.addWidget(self.monthly_employee_select)
        monthly_period = QHBoxLayout()
        self.monthly_month_select = QComboBox()
        for month in range(1, 13):
            self.monthly_month_select.addItem(
                ("Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος",
                 "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος")[month - 1],
                month,
            )
        self.monthly_year_input = QLineEdit(str(QDate.currentDate().year()))
        self.monthly_year_input.setPlaceholderText("Έτος")
        monthly_period.addWidget(self.monthly_month_select)
        monthly_period.addWidget(self.monthly_year_input)
        layout.addLayout(monthly_period)
        self.monthly_paid_hours_input = QLineEdit()
        self.monthly_paid_hours_input.setPlaceholderText("Ώρες που έχουν ήδη πληρωθεί")
        layout.addWidget(self.monthly_paid_hours_input)
        self.monthly_hourly_rate_input = QLineEdit()
        self.monthly_hourly_rate_input.setPlaceholderText("Ωρομίσθιο (€)")
        layout.addWidget(self.monthly_hourly_rate_input)
        self.new_payment_amount_input = QLineEdit()
        self.new_payment_amount_input.setPlaceholderText("Πληρωμή τώρα (€), π.χ. 50 ή 50,00")
        layout.addWidget(self.new_payment_amount_input)
        add_payment_btn = QPushButton("Καταχώριση Πληρωμής")
        add_payment_btn.clicked.connect(self.add_monthly_payment)
        layout.addWidget(add_payment_btn)
        monthly_save = QPushButton("Αποθήκευση Μηνιαίας Μισθοδοσίας")
        monthly_save.clicked.connect(self.save_monthly_payroll)
        layout.addWidget(monthly_save)
        monthly_refresh = QPushButton("Προβολή Επιλεγμένου Μήνα")
        monthly_refresh.clicked.connect(self.load_monthly_summary)
        layout.addWidget(monthly_refresh)
        self.monthly_month_select.currentIndexChanged.connect(self.refresh_monthly_views)
        self.monthly_year_input.textChanged.connect(self.refresh_monthly_views)
        self.monthly_employee_select.currentIndexChanged.connect(self.refresh_monthly_views)
        self.monthly_summary_label = QLabel("")
        layout.addWidget(self.monthly_summary_label)

        parent.setLayout(layout)

    def build_reports_ui(self, parent):
        layout = QVBoxLayout()

        self.history_table = QTableWidget()
        self.history_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.history_table.setSelectionMode(QTableWidget.SingleSelection)
        self.history_table.setColumnCount(7)
        self.history_table.setHorizontalHeaderLabels([
            "ID", "Υπάλληλος", "Ημερομηνία", "Προσέλευση", "Αποχώρηση",
            "Διάλειμμα (λεπτά)", "Πραγματικές Ώρες"
        ])
        layout.addWidget(self.history_table)

        buttons = QHBoxLayout()

        for text, action in [
            ("Επεξεργασία", self.load_selected_record_for_edit),
            ("Διαγραφή", self.delete_selected_record),
            ("Export Excel", self.export_to_excel),
            ("Export PDF", self.export_to_pdf),
            ("Backup Database", self.backup_database),
            ("Charts", self.create_charts),
        ]:
            btn = QPushButton(text)
            btn.clicked.connect(action)
            buttons.addWidget(btn)

        layout.addLayout(buttons)
        employee_export_btn = QPushButton("Εξαγωγή Παρουσιών & Πληρωμών Υπαλλήλου (Excel)")
        employee_export_btn.clicked.connect(self.export_employee_payroll_to_excel)
        layout.addWidget(employee_export_btn)
        parent.setLayout(layout)

    def build_settings_ui(self, parent):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Prometheus Payroll v1.1.0 — Hourly Payroll Edition"))
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
                unpaid_overtime_rate REAL DEFAULT 0,
                unpaid_overtime_amount REAL DEFAULT 0,
                unpaid_overtime_paid INTEGER DEFAULT 0,
                overtime_payment_date TEXT
            )
        """)

        self.conn.commit()
        columns = self.table_columns("attendance")
        required = {"unpaid_overtime_amount", "unpaid_overtime_paid", "overtime_payment_date", "overtime_hours", "unpaid_overtime_rate"}
        employee_columns = self.table_columns("employees")
        employee_rate_missing = "hourly_rate" not in employee_columns
        monthly_exists = self.cursor.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='monthly_payments'"
        ).fetchone() is not None
        payments_exists = self.cursor.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='payroll_payments'"
        ).fetchone() is not None
        current_version = self.get_schema_version()
        if required.issubset(columns) and monthly_exists and payments_exists and not employee_rate_missing and current_version >= 4:
            return ""

        backup_path = self.create_pre_migration_backup()
        changed_rows = malformed_rows = 0
        recalculate_history = current_version < 2
        self.cursor.execute("BEGIN IMMEDIATE")
        try:
            if employee_rate_missing:
                self.cursor.execute("ALTER TABLE employees ADD COLUMN hourly_rate REAL NOT NULL DEFAULT 0")
            for name, definition in [
                ("unpaid_overtime_amount", "REAL DEFAULT 0"),
                ("unpaid_overtime_paid", "INTEGER DEFAULT 0"),
                ("overtime_payment_date", "TEXT"),
                ("overtime_hours", "REAL DEFAULT 0"),
                ("unpaid_overtime_rate", "REAL DEFAULT 0"),
            ]:
                if name not in self.table_columns("attendance"):
                    self.cursor.execute(f"ALTER TABLE attendance ADD COLUMN {name} {definition}")
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS monthly_payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_id INTEGER NOT NULL,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    paid_hours REAL NOT NULL DEFAULT 0,
                    hourly_rate REAL NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(employee_id, year, month),
                    FOREIGN KEY(employee_id) REFERENCES employees(id)
                )
            """)
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS payroll_payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_id INTEGER NOT NULL,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    amount REAL NOT NULL CHECK(amount > 0),
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(employee_id) REFERENCES employees(id)
                )
            """)
            if recalculate_history:
                rows = self.cursor.execute("SELECT id, date, check_in, check_out, break_minutes FROM attendance").fetchall()
                for row_id, date_value, check_in, check_out, break_minutes in rows:
                    try:
                        start = datetime.strptime(self.normalize_time(check_in or ""), "%H:%M")
                        end = datetime.strptime(self.normalize_time(check_out or ""), "%H:%M")
                        if end <= start:
                            from datetime import timedelta
                            end += timedelta(days=1)
                        actual = (end - start).total_seconds() / 3600 - int(break_minutes or 0) / 60
                        parsed_date = datetime.strptime(str(date_value), "%Y-%m-%d") if "-" in str(date_value) else datetime.strptime(str(date_value), "%Y/%m/%d")
                        if actual < 0:
                            raise ValueError
                    except (ValueError, TypeError):
                        malformed_rows += 1
                        continue
                    self.cursor.execute(
                        "UPDATE attendance SET date=?, total_hours=?, unpaid_hours=MAX(0, ? - COALESCE(paid_hours, 0)), overtime_hours=0 WHERE id=?",
                        (parsed_date.strftime("%Y-%m-%d"), actual, actual, row_id),
                    )
                    changed_rows += 1
            self.cursor.execute("PRAGMA user_version = 4")
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise
        migration_details = (
            f"Επαναϋπολογίστηκαν {changed_rows} παρουσίες και παραλείφθηκαν {malformed_rows} μη έγκυρες εγγραφές."
            if recalculate_history else "Το ιστορικό παρουσιών διατηρήθηκε ανέπαφο."
        )
        return f"Η βάση ενημερώθηκε. Αντίγραφο ασφαλείας: {backup_path}\n{migration_details}"

    def table_columns(self, table_name):
        self.cursor.execute(f"PRAGMA table_info({table_name})")
        return {column[1] for column in self.cursor.fetchall()}

    def get_schema_version(self):
        return self.cursor.execute("PRAGMA user_version").fetchone()[0]

    def create_pre_migration_backup(self):
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(APP_DATA_DIR, f"payroll_backup_before_v1_1_0_{timestamp}.db")
        self.conn.commit()
        destination = sqlite3.connect(backup_path)
        try:
            self.conn.backup(destination)
        finally:
            destination.close()
        return backup_path

    def add_column_if_missing(self, table_name, column_name, column_type):
        self.cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [column[1] for column in self.cursor.fetchall()]

        if column_name not in columns:
            self.cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")

    def normalize_time(self, value):
        value = value.strip()
        match = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?", value)
        if not match:
            raise ValueError
        hour, minute = int(match.group(1)), int(match.group(2) or 0)
        if hour > 23 or minute > 59:
            raise ValueError
        return f"{hour:02d}:{minute:02d}"

    def parse_number(self, value):
        value = value.strip().replace(",", ".")
        if value == "":
            return 0
        return float(value)

    def load_employees(self):
        current_filter_id = self.filter_employee_select.currentData() if self.filter_employee_select.count() > 0 else None
        current_monthly_id = self.monthly_employee_select.currentData() if hasattr(self, "monthly_employee_select") else None

        self.employee_list.clear()
        self.employee_select.clear()
        if hasattr(self, "monthly_employee_select"):
            self.monthly_employee_select.clear()

        self.filter_employee_select.blockSignals(True)
        self.filter_employee_select.clear()
        self.filter_employee_select.addItem("Όλοι οι υπάλληλοι", None)

        self.cursor.execute("SELECT id, name, hourly_rate FROM employees ORDER BY name")
        employees = self.cursor.fetchall()

        selected_index = 0
        monthly_selected_index = 0

        for employee_id, name, hourly_rate in employees:
            self.employee_list.addItem(f"{name} — {float(hourly_rate or 0):.2f} €/ώρα")
            self.employee_list.item(self.employee_list.count() - 1).setData(32, employee_id)
            self.employee_list.item(self.employee_list.count() - 1).setToolTip(f"Ωρομίσθιο: {float(hourly_rate or 0):.2f} €")
            self.employee_select.addItem(name, employee_id)
            self.filter_employee_select.addItem(name, employee_id)
            if hasattr(self, "monthly_employee_select"):
                self.monthly_employee_select.addItem(name, employee_id)
                if employee_id == current_monthly_id:
                    monthly_selected_index = self.monthly_employee_select.count() - 1
            if current_filter_id == employee_id:
                selected_index = self.filter_employee_select.count() - 1

        if self.employee_list.count() > 0 and self.employee_list.currentRow() < 0:
            self.employee_list.setCurrentRow(0)

        self.filter_employee_select.setCurrentIndex(selected_index)
        self.filter_employee_select.blockSignals(False)
        if hasattr(self, "monthly_employee_select"):
            self.monthly_employee_select.blockSignals(True)
            self.monthly_employee_select.setCurrentIndex(monthly_selected_index)
            self.monthly_employee_select.blockSignals(False)
        self.sync_monthly_from_payroll_filters()

    def load_selected_employee_rate(self, item, _previous=None):
        if item is None:
            self.employee_hourly_rate_input.clear()
            return
        row = self.cursor.execute(
            "SELECT hourly_rate FROM employees WHERE id=?", (item.data(32),)
        ).fetchone()
        self.employee_hourly_rate_input.setText(str(float(row[0] or 0)) if row else "0")

    def save_employee_hourly_rate(self):
        item = self.employee_list.currentItem()
        if item is None:
            QMessageBox.warning(self, "Σφάλμα", "Επίλεξε υπάλληλο για να αποθηκεύσεις το ωρομίσθιο.")
            return
        try:
            hourly_rate = self.parse_number(self.employee_hourly_rate_input.text())
            if hourly_rate < 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Σφάλμα", "Μη έγκυρο ωρομίσθιο. Χρησιμοποίησε αριθμό όπως 7, 7.5 ή 7,5.")
            return
        employee_id = item.data(32)
        employee_name = item.text().split(" — ", 1)[0]
        self.cursor.execute("UPDATE employees SET hourly_rate=? WHERE id=?", (hourly_rate, employee_id))
        self.conn.commit()
        self.load_employees()
        self.load_monthly_summary()
        self.load_history()
        QMessageBox.information(self, "Αποθηκεύτηκε", f"Ωρομίσθιο {hourly_rate:.2f} € αποθηκεύτηκε για τον/την {employee_name}.")

    def add_employee(self):
        name = self.employee_input.text().strip()

        if name == "":
            QMessageBox.warning(self, "Σφάλμα", "Δώσε όνομα υπαλλήλου")
            return

        try:
            hourly_rate = self.parse_number(self.employee_hourly_rate_input.text())
            if hourly_rate < 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Σφάλμα", "Μη έγκυρο ωρομίσθιο. Χρησιμοποίησε αριθμό όπως 7, 7.5 ή 7,5.")
            return

        self.cursor.execute("INSERT INTO employees (name, hourly_rate) VALUES (?, ?)", (name, hourly_rate))
        self.conn.commit()

        self.employee_input.clear()
        self.employee_hourly_rate_input.clear()
        self.load_employees()
        self.load_history()

    def remove_employee(self):
        selected_item = self.employee_list.currentItem()

        if selected_item is None:
            QMessageBox.warning(self, "Σφάλμα", "Επίλεξε υπάλληλο")
            return

        employee_id = selected_item.data(32)
        self.cursor.execute("SELECT id, name, hourly_rate FROM employees WHERE id = ?", (employee_id,))
        employee = self.cursor.fetchone()

        if employee is None:
            QMessageBox.warning(self, "Σφάλμα", "Δεν βρέθηκε ο υπάλληλος")
            return

        employee_id, employee_name, hourly_rate = employee

        self.cursor.execute("SELECT COUNT(*) FROM attendance WHERE employee_id = ?", (employee_id,))
        attendance_count = self.cursor.fetchone()[0]

        if attendance_count > 0:
            QMessageBox.warning(
                self,
                "Δεν επιτρέπεται",
                "Ο υπάλληλος έχει καταχωρήσεις μισθοδοσίας. Δεν διαγράφεται για να μη χαθεί ιστορικό."
            )
            return

        confirm = QMessageBox.question(self, "Επιβεβαίωση", f"Να αφαιρεθεί ο υπάλληλος {employee_name};")

        if confirm != QMessageBox.Yes:
            return

        self.last_deleted_employee = {
            "id": employee_id,
            "name": employee_name,
            "hourly_rate": hourly_rate,
        }

        self.cursor.execute("DELETE FROM employees WHERE id = ?", (employee_id,))
        self.conn.commit()

        self.load_employees()
        self.load_history()

        QMessageBox.information(
            self,
            "Διαγράφηκε",
            "Ο υπάλληλος διαγράφηκε. Μπορείς να πατήσεις Αναίρεση Διαγραφής Υπαλλήλου."
        )

    def undo_delete_employee(self):
        if not self.last_deleted_employee:
            QMessageBox.information(self, "Αναίρεση", "Δεν υπάρχει πρόσφατη διαγραφή υπαλλήλου για αναίρεση.")
            return

        employee_id = self.last_deleted_employee["id"]
        employee_name = self.last_deleted_employee["name"]

        self.cursor.execute("SELECT id FROM employees WHERE id = ?", (employee_id,))
        existing = self.cursor.fetchone()

        if existing:
            QMessageBox.warning(self, "Σφάλμα", "Ο υπάλληλος υπάρχει ήδη στη βάση.")
            self.last_deleted_employee = None
            return

        self.cursor.execute(
            "INSERT INTO employees (id, name, hourly_rate) VALUES (?, ?, ?)",
            (employee_id, employee_name, self.last_deleted_employee.get("hourly_rate", 0))
        )
        self.conn.commit()

        self.last_deleted_employee = None
        self.load_employees()
        self.load_history()

        QMessageBox.information(self, "Αναίρεση", f"Ο υπάλληλος {employee_name} επανήλθε.")

    def calculate_attendance_values(self):
        check_in_text = self.check_in_input.text().strip()
        check_out_text = self.check_out_input.text().strip()
        break_text = self.break_input.text().strip()

        if check_in_text == "" or check_out_text == "":
            raise ValueError("missing_time")

        try:
            check_in = self.normalize_time(check_in_text)
            check_out = self.normalize_time(check_out_text)
            start = datetime.strptime(check_in, "%H:%M")
            end = datetime.strptime(check_out, "%H:%M")
        except Exception:
            raise ValueError("invalid_time")

        try:
            break_minutes = int(break_text) if break_text else 0
        except Exception:
            raise ValueError("invalid_number")

        if break_minutes < 0:
            raise ValueError("negative_number")

        paid_hours = 0

        unpaid_overtime_paid = 0
        overtime_payment_date = None

        if end <= start:
            from datetime import timedelta
            end += timedelta(days=1)
        actual_hours = ((end - start).total_seconds() / 60 - break_minutes) / 60

        if actual_hours < 0:
            raise ValueError("negative_hours")

        total_hours = actual_hours
        overtime_hours = 0
        unpaid_hours = max(0, total_hours - paid_hours)
        unpaid_overtime_amount = 0

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
        except ValueError as error:
            error_text = str(error)

            if error_text == "missing_time":
                QMessageBox.warning(self, "Σφάλμα", "Συμπλήρωσε προσέλευση και αποχώρηση.")
            elif error_text == "invalid_time":
                QMessageBox.warning(self, "Σφάλμα", "Μη έγκυρη ώρα. Βάλε ώρα όπως 9, 09:00 ή 9:30.")
            elif error_text == "invalid_number":
                QMessageBox.warning(self, "Σφάλμα", "Μη έγκυρος αριθμός. Βάλε τιμή όπως 8, 8.5 ή 8,5.")
            elif error_text == "negative_number":
                QMessageBox.warning(self, "Σφάλμα", "Οι αριθμοί δεν μπορούν να είναι αρνητικοί.")
            elif error_text == "negative_hours":
                QMessageBox.warning(self, "Σφάλμα", "Η αποχώρηση πρέπει να είναι μετά την προσέλευση.")
            else:
                QMessageBox.warning(self, "Σφάλμα", "Έλεγξε ώρες και αριθμούς.")
            return

        if self.editing_record_id is None:
            self.cursor.execute("""
                INSERT INTO attendance (
                    employee_id, date, check_in, check_out, break_minutes,
                    total_hours, paid_hours, unpaid_hours, overtime_hours,
                    unpaid_overtime_rate, unpaid_overtime_amount, unpaid_overtime_paid,
                    overtime_payment_date
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                employee_id, date, values["check_in"], values["check_out"],
                values["break_minutes"], values["total_hours"], values["paid_hours"],
                values["unpaid_hours"], values["overtime_hours"], 0,
                values["unpaid_overtime_amount"], values["unpaid_overtime_paid"],
                values["overtime_payment_date"]
            ))
            message = "Αποθηκεύτηκε"
            self.last_edit_snapshot = None
        else:
            self.cursor.execute("""
                SELECT employee_id, date, check_in, check_out, break_minutes,
                       total_hours, paid_hours, unpaid_hours, overtime_hours,
                       unpaid_overtime_rate, unpaid_overtime_amount, unpaid_overtime_paid,
                       overtime_payment_date
                FROM attendance
                WHERE id = ?
            """, (self.editing_record_id,))

            previous_record = self.cursor.fetchone()

            if previous_record is not None:
                self.last_edit_snapshot = {
                    "id": self.editing_record_id,
                    "values": previous_record
                }

            self.cursor.execute("""
                UPDATE attendance
                SET employee_id = ?, date = ?, check_in = ?, check_out = ?,
                    break_minutes = ?, total_hours = ?, paid_hours = ?,
                    unpaid_hours = ?, overtime_hours = ?,
                    unpaid_overtime_rate = ?, unpaid_overtime_amount = ?, unpaid_overtime_paid = ?,
                    overtime_payment_date = ?
                WHERE id = ?
            """, (
                employee_id, date, values["check_in"], values["check_out"],
                values["break_minutes"], values["total_hours"], values["paid_hours"],
                values["unpaid_hours"], values["overtime_hours"], 0,
                values["unpaid_overtime_amount"], values["unpaid_overtime_paid"],
                values["overtime_payment_date"], self.editing_record_id
            ))
            message = "Ενημερώθηκε"

        self.conn.commit()

        self.filter_employee_select.setCurrentIndex(0)
        self.employee_search_input.clear()
        monthly_index = self.monthly_employee_select.findData(employee_id)
        if monthly_index >= 0:
            self.monthly_employee_select.blockSignals(True)
            self.monthly_employee_select.setCurrentIndex(monthly_index)
            self.monthly_employee_select.blockSignals(False)

        rate_row = self.cursor.execute(
            "SELECT hourly_rate FROM employees WHERE id=?", (employee_id,)
        ).fetchone()
        hourly_rate = float(rate_row[0] or 0) if rate_row else 0.0
        daily_amount = values["total_hours"] * hourly_rate
        self.result_label.setText(
            f"{message}: {values['total_hours']:.2f} πραγματικές ώρες × "
            f"{hourly_rate:.2f} €/ώρα = {daily_amount:.2f} €"
        )

        self.clear_attendance_form()
        self.load_employees()
        self.load_history()
        self.load_monthly_summary()

    def save_monthly_payroll(self):
        employee_id = self.monthly_employee_select.currentData()
        try:
            year = int(self.monthly_year_input.text().strip())
            month = int(self.monthly_month_select.currentData())
            paid_hours = self.parse_number(self.monthly_paid_hours_input.text())
            hourly_rate = self.parse_number(self.monthly_hourly_rate_input.text())
            if year < 1900 or year > 9999 or paid_hours < 0 or hourly_rate < 0:
                raise ValueError
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Σφάλμα", "Έλεγξε το έτος, τις πληρωμένες ώρες και το ωρομίσθιο. Χρησιμοποίησε τελεία ή κόμμα στα δεκαδικά.")
            return
        if employee_id is None:
            QMessageBox.warning(self, "Σφάλμα", "Επίλεξε υπάλληλο.")
            return
        if not self.monthly_hourly_rate_input.text().strip():
            rate_row = self.cursor.execute(
                "SELECT hourly_rate FROM employees WHERE id=?", (employee_id,)
            ).fetchone()
            hourly_rate = float(rate_row[0] or 0) if rate_row else 0.0
        else:
            hourly_rate = self.parse_number(self.monthly_hourly_rate_input.text())
        now = datetime.now().isoformat(timespec="seconds")
        self.cursor.execute("""
            INSERT INTO monthly_payments(employee_id, year, month, paid_hours, hourly_rate, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(employee_id, year, month) DO UPDATE SET
                paid_hours=excluded.paid_hours,
                hourly_rate=excluded.hourly_rate,
                updated_at=excluded.updated_at
        """, (employee_id, year, month, paid_hours, hourly_rate, now, now))
        self.conn.commit()
        self.load_monthly_summary()
        self.load_history()

    def refresh_monthly_views(self, *_args):
        self.fill_default_monthly_rate()
        self.load_monthly_summary()
        self.load_history()

    def sync_monthly_from_payroll_filters(self, *_args):
        if not hasattr(self, "monthly_employee_select") or self.monthly_employee_select.count() == 0:
            self.load_history()
            return

        employee_id = self.filter_employee_select.currentData()
        if employee_id is not None:
            employee_index = self.monthly_employee_select.findData(employee_id)
            if employee_index >= 0:
                self.monthly_employee_select.blockSignals(True)
                self.monthly_employee_select.setCurrentIndex(employee_index)
                self.monthly_employee_select.blockSignals(False)

        period_date = self.to_date_input.date()
        month_index = self.monthly_month_select.findData(period_date.month())
        self.monthly_month_select.blockSignals(True)
        if month_index >= 0:
            self.monthly_month_select.setCurrentIndex(month_index)
        self.monthly_month_select.blockSignals(False)
        self.monthly_year_input.blockSignals(True)
        self.monthly_year_input.setText(str(period_date.year()))
        self.monthly_year_input.blockSignals(False)
        self.load_monthly_summary()
        self.load_history()

    def fill_default_monthly_rate(self):
        if not hasattr(self, "monthly_employee_select") or self.monthly_employee_select.currentData() is None:
            return
        try:
            year = int(self.monthly_year_input.text().strip())
            month = int(self.monthly_month_select.currentData())
        except (ValueError, TypeError):
            return
        employee_id = self.monthly_employee_select.currentData()
        saved = self.cursor.execute(
            "SELECT hourly_rate FROM monthly_payments WHERE employee_id=? AND year=? AND month=?",
            (employee_id, year, month),
        ).fetchone()
        if saved:
            rate = saved[0]
        else:
            row = self.cursor.execute("SELECT hourly_rate FROM employees WHERE id=?", (employee_id,)).fetchone()
            rate = row[0] if row else 0
        self.monthly_hourly_rate_input.setText(str(float(rate or 0)))

    def calculate_monthly_payroll(self, employee_id, year, month):
        start = f"{year:04d}-{month:02d}-01"
        end = f"{year:04d}-{month + 1:02d}-01" if month < 12 else f"{year + 1:04d}-01-01"
        rows = self.cursor.execute(
            "SELECT date, total_hours, check_in, check_out, break_minutes FROM attendance WHERE employee_id=?",
            (employee_id,),
        ).fetchall()
        actual_hours = 0.0
        for date_value, stored_hours, check_in, check_out, break_minutes in rows:
            try:
                normalized_date = datetime.strptime(
                    str(date_value), "%Y-%m-%d" if "-" in str(date_value) else "%Y/%m/%d"
                ).strftime("%Y-%m-%d")
                if not (start <= normalized_date < end):
                    continue
                start_time = datetime.strptime(self.normalize_time(check_in or ""), "%H:%M")
                end_time = datetime.strptime(self.normalize_time(check_out or ""), "%H:%M")
                if end_time <= start_time:
                    from datetime import timedelta
                    end_time += timedelta(days=1)
                actual_hours += max(
                    0.0,
                    (end_time - start_time).total_seconds() / 3600 - int(break_minutes or 0) / 60,
                )
            except (TypeError, ValueError):
                try:
                    normalized_date = datetime.strptime(
                        str(date_value), "%Y-%m-%d" if "-" in str(date_value) else "%Y/%m/%d"
                    ).strftime("%Y-%m-%d")
                    if start <= normalized_date < end:
                        actual_hours += max(0.0, float(stored_hours or 0))
                except (TypeError, ValueError):
                    continue

        payment = self.cursor.execute(
            "SELECT paid_hours, hourly_rate FROM monthly_payments WHERE employee_id=? AND year=? AND month=?",
            (employee_id, year, month),
        ).fetchone()
        if payment:
            paid_hours, hourly_rate = float(payment[0] or 0), float(payment[1] or 0)
        else:
            paid_hours = 0.0
            rate_row = self.cursor.execute(
                "SELECT hourly_rate FROM employees WHERE id=?", (employee_id,)
            ).fetchone()
            hourly_rate = float(rate_row[0] or 0) if rate_row else 0.0
        cash_payments = float(self.cursor.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM payroll_payments WHERE employee_id=? AND year=? AND month=?",
            (employee_id, year, month),
        ).fetchone()[0] or 0)
        gross_pay = actual_hours * hourly_rate
        total_paid_amount = paid_hours * hourly_rate + cash_payments
        amount_due = max(0.0, gross_pay - total_paid_amount)
        extra_paid_amount = max(0.0, total_paid_amount - gross_pay)
        if hourly_rate > 0:
            paid_equivalent_hours = total_paid_amount / hourly_rate
            outstanding_hours = amount_due / hourly_rate
            extra_paid_hours = extra_paid_amount / hourly_rate
        else:
            paid_equivalent_hours = paid_hours
            outstanding_hours = max(0.0, actual_hours - paid_hours)
            extra_paid_hours = max(0.0, paid_hours - actual_hours)
        return {
            "actual_hours": actual_hours,
            "paid_hours": paid_hours,
            "paid_equivalent_hours": paid_equivalent_hours,
            "cash_payments": cash_payments,
            "total_paid_amount": total_paid_amount,
            "gross_pay": gross_pay,
            "outstanding_hours": outstanding_hours,
            "extra_paid_hours": extra_paid_hours,
            "extra_paid_amount": extra_paid_amount,
            "hourly_rate": hourly_rate,
            "amount_due": amount_due,
        }

    def load_monthly_summary(self):
        if not hasattr(self, "monthly_summary_label"):
            return
        employee_id = self.monthly_employee_select.currentData()
        try:
            year = int(self.monthly_year_input.text().strip())
            month = int(self.monthly_month_select.currentData())
        except (ValueError, TypeError):
            self.monthly_summary_label.setText("Έλεγξε μήνα και έτος.")
            return
        if employee_id is None:
            self.monthly_summary_label.setText("Πρόσθεσε ή επίλεξε υπάλληλο για μηνιαία σύνοψη.")
            return
        payroll = self.calculate_monthly_payroll(employee_id, year, month)
        self.monthly_paid_hours_input.setText(str(payroll["paid_hours"]))
        self.monthly_hourly_rate_input.setText(str(payroll["hourly_rate"]))
        self.monthly_summary_label.setText(
            f"Σύνολο μισθοδοσίας: {payroll['gross_pay']:.2f} € | "
            f"Πληρωμένο ποσό: {payroll['total_paid_amount']:.2f} € "
            f"(δόσεις: {payroll['cash_payments']:.2f} €) | "
            f"Πραγματικές ώρες: {payroll['actual_hours']:.2f} | "
            f"Πληρωμένες ώρες: {payroll['paid_equivalent_hours']:.2f} | "
            f"Υπόλοιπο: {payroll['outstanding_hours']:.2f} ώρες | Ωρομίσθιο: {payroll['hourly_rate']:.2f} € | "
            f"Οφειλόμενο: {payroll['amount_due']:.2f} €" +
            (f" | Επιπλέον πληρωμένα: {payroll['extra_paid_amount']:.2f} €" if payroll['extra_paid_amount'] else "")
        )

    def add_monthly_payment(self):
        employee_id = self.monthly_employee_select.currentData()
        try:
            year = int(self.monthly_year_input.text().strip())
            month = int(self.monthly_month_select.currentData())
            amount = self.parse_number(self.new_payment_amount_input.text())
            if year < 1900 or year > 9999 or amount <= 0:
                raise ValueError
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Σφάλμα", "Βάλε έγκυρο μήνα, έτος και ποσό πληρωμής μεγαλύτερο από 0.")
            return
        if employee_id is None:
            QMessageBox.warning(self, "Σφάλμα", "Επίλεξε υπάλληλο για την πληρωμή.")
            return
        self.cursor.execute(
            "INSERT INTO payroll_payments(employee_id, year, month, amount, created_at) VALUES (?, ?, ?, ?, ?)",
            (employee_id, year, month, amount, datetime.now().isoformat(timespec="seconds")),
        )
        self.conn.commit()
        self.new_payment_amount_input.clear()
        self.load_monthly_summary()
        self.load_history()

    def clear_attendance_form(self):
        self.editing_record_id = None
        self.save_btn.setText("Αποθήκευση Ωρών")
        self.check_in_input.clear()
        self.check_out_input.clear()
        self.break_input.clear()

    def cancel_edit(self):
        self.clear_attendance_form()
        self.result_label.setText("Η επεξεργασία ακυρώθηκε")

    def clear_filters(self):
        self.employee_search_input.clear()
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
            SELECT id, employee_id, date, check_in, check_out, break_minutes
            FROM attendance WHERE id = ?
        """, (record_id,))

        record = self.cursor.fetchone()

        if record is None:
            QMessageBox.warning(self, "Σφάλμα", "Δεν βρέθηκε η καταχώρηση")
            return

        (
            attendance_id, employee_id, date, check_in, check_out, break_minutes
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

        self.save_btn.setText("Ενημέρωση Καταχώρησης")
        self.result_label.setText(f"Επεξεργασία ID {attendance_id}")

        if hasattr(self, "tabs"):
            self.tabs.setCurrentIndex(1)

    def undo_last_edit(self):
        if not self.last_edit_snapshot:
            QMessageBox.information(self, "Αναίρεση", "Δεν υπάρχει πρόσφατη αλλαγή μισθοδοσίας για αναίρεση.")
            return

        record_id = self.last_edit_snapshot["id"]
        previous_values = self.last_edit_snapshot["values"]

        self.cursor.execute("""
            UPDATE attendance
            SET employee_id = ?,
                date = ?,
                check_in = ?,
                check_out = ?,
                break_minutes = ?,
                total_hours = ?,
                paid_hours = ?,
                unpaid_hours = ?,
                overtime_hours = ?,
                unpaid_overtime_rate = ?,
                unpaid_overtime_amount = ?,
                unpaid_overtime_paid = ?,
                overtime_payment_date = ?
            WHERE id = ?
        """, (*previous_values, record_id))

        self.conn.commit()
        self.last_edit_snapshot = None
        self.clear_attendance_form()
        self.load_employees()
        self.load_history()

        QMessageBox.information(self, "Αναίρεση", f"Η αλλαγή της καταχώρησης ID {record_id} αναιρέθηκε.")

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

    def get_filtered_records(self):
        employee_id = self.filter_employee_select.currentData()
        search = self.employee_search_input.text().strip()

        from_date = self.from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.to_date_input.date().toString("yyyy-MM-dd")

        query = """
            SELECT attendance.id, employees.name, attendance.date,
                   attendance.check_in, attendance.check_out,
                   attendance.break_minutes, attendance.total_hours
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

        query += " ORDER BY attendance.date DESC, attendance.id DESC"

        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def load_history(self):
        records = self.get_filtered_records()

        self.history_table.clearContents()
        self.history_table.setRowCount(len(records))

        totals = {}
        total_hours_sum = 0.0
        month = int(self.monthly_month_select.currentData())
        year = int(self.monthly_year_input.text())

        for row, record in enumerate(records):
            employee = record[1]
            total_hours = float(record[6] or 0)
            total_hours_sum += total_hours

            if employee not in totals:
                totals[employee] = {"id": self.cursor.execute("SELECT id FROM employees WHERE name=?", (employee,)).fetchone()[0], "hours": 0.0}
            totals[employee]["hours"] += total_hours
            for col, value in enumerate(record):
                item = QTableWidgetItem(str(value))
                self.history_table.setItem(row, col, item)

        self.history_table.resizeColumnsToContents()

        self.summary_label.setText(
            f"Σύνολα περιόδου\nΠραγματικές ώρες εργασίας: {total_hours_sum:.2f}"
        )

        self.employee_summary_table.clearContents()
        self.employee_summary_table.setRowCount(len(totals))

        for row, (employee, values) in enumerate(totals.items()):
            employee_id = values["id"]
            payroll = self.calculate_monthly_payroll(employee_id, year, month)
            month_hours = payroll["actual_hours"]
            row_values = [employee, f"{month_hours:.2f}", f"{payroll['paid_equivalent_hours']:.2f}",
                          f"{payroll['outstanding_hours']:.2f}" + (f" (επιπλέον {payroll['extra_paid_hours']:.2f})" if payroll['extra_paid_hours'] else ""),
                          f"{payroll['gross_pay']:.2f} €", f"{payroll['total_paid_amount']:.2f} €",
                          f"{payroll['amount_due']:.2f} €"]

            for col, value in enumerate(row_values):
                item = QTableWidgetItem(value)
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
            "Διάλειμμα (λεπτά)", "Πραγματικές Ώρες"
        ]

        sheet.append(headers)

        for record in records:
            row = list(record)
            sheet.append(row)

        workbook.save(file_path)
        QMessageBox.information(self, "Ολοκληρώθηκε", "Το Excel αποθηκεύτηκε")

    def export_employee_payroll_to_excel(self):
        employee_id = self.filter_employee_select.currentData()
        if employee_id is None:
            QMessageBox.warning(self, "Επίλεξε υπάλληλο", "Στα φίλτρα Payroll επίλεξε έναν συγκεκριμένο υπάλληλο πριν από την εξαγωγή.")
            return

        employee = self.cursor.execute(
            "SELECT name, hourly_rate FROM employees WHERE id=?", (employee_id,)
        ).fetchone()
        if employee is None:
            QMessageBox.warning(self, "Σφάλμα", "Δεν βρέθηκε ο επιλεγμένος υπάλληλος.")
            return
        employee_name, base_rate = employee
        from_date = self.from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.to_date_input.date().toString("yyyy-MM-dd")
        start_year, start_month = map(int, from_date[:7].split("-"))
        end_year, end_month = map(int, to_date[:7].split("-"))
        start_period = start_year * 100 + start_month
        end_period = end_year * 100 + end_month

        attendance = self.cursor.execute("""
            SELECT date, check_in, check_out, break_minutes, total_hours
            FROM attendance
            WHERE employee_id=? AND date>=? AND date<=?
            ORDER BY date, id
        """, (employee_id, from_date, to_date)).fetchall()
        month_payrolls = self.cursor.execute("""
            SELECT year, month, paid_hours, hourly_rate, updated_at
            FROM monthly_payments
            WHERE employee_id=? AND year*100+month BETWEEN ? AND ?
            ORDER BY year, month
        """, (employee_id, start_period, end_period)).fetchall()
        cash_payments = self.cursor.execute("""
            SELECT year, month, amount, created_at
            FROM payroll_payments
            WHERE employee_id=? AND year*100+month BETWEEN ? AND ?
            ORDER BY year, month, created_at, id
        """, (employee_id, start_period, end_period)).fetchall()

        safe_name = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in employee_name).strip("_")
        suggested = f"{safe_name}_payroll_{from_date}_to_{to_date}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Εξαγωγή Παρουσιών & Πληρωμών", suggested, "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        workbook = Workbook()
        summary = workbook.active
        summary.title = "Σύνοψη"
        summary.append(["Υπάλληλος", employee_name])
        summary.append(["Περίοδος παρουσιών", from_date, to_date])
        summary.append(["Πραγματικές ώρες", sum(float(row[4] or 0) for row in attendance)])
        summary.append(["Πληρωμένες ώρες payroll", sum(float(row[2] or 0) for row in month_payrolls)])
        monthly_paid_amount = sum(float(row[2] or 0) * float(row[3] or 0) for row in month_payrolls)
        cash_paid_amount = sum(float(row[2] or 0) for row in cash_payments)
        summary.append(["Ποσό πληρωμένων ωρών", monthly_paid_amount])
        summary.append(["Καταχωρημένες πληρωμές ποσού", cash_paid_amount])
        summary.append(["Συνολικά καταχωρημένο πληρωμένο ποσό", monthly_paid_amount + cash_paid_amount])
        summary.append(["Βασικό ωρομίσθιο", float(base_rate or 0)])

        attendance_sheet = workbook.create_sheet("Παρουσίες")
        attendance_sheet.append(["Ημερομηνία", "Προσέλευση", "Αποχώρηση", "Διάλειμμα (λεπτά)", "Πραγματικές ώρες", "Ωρομίσθιο (€)", "Αξία ωρών (€)"])
        gross_amount = 0.0
        for date_value, check_in, check_out, break_minutes, hours in attendance:
            year, month = map(int, date_value[:7].split("-"))
            rate_row = self.cursor.execute(
                "SELECT hourly_rate FROM monthly_payments WHERE employee_id=? AND year=? AND month=?",
                (employee_id, year, month),
            ).fetchone()
            rate = float(rate_row[0] or 0) if rate_row else float(base_rate or 0)
            hours = float(hours or 0)
            gross_amount += hours * rate
            attendance_sheet.append([date_value, check_in, check_out, break_minutes, hours, rate, hours * rate])
        summary.append(["Αξία παρουσιών στην περίοδο", gross_amount])

        payments_sheet = workbook.create_sheet("Πληρωμές")
        payments_sheet.append(["Έτος", "Μήνας", "Τύπος", "Ποσό (€)", "Ώρες", "Ωρομίσθιο (€)", "Ημερομηνία καταχώρισης"])
        for year, month, paid_hours, rate, updated_at in month_payrolls:
            payments_sheet.append([year, month, "Ώρες ήδη πληρωμένες", float(paid_hours or 0) * float(rate or 0), float(paid_hours or 0), float(rate or 0), updated_at])
        for year, month, amount, created_at in cash_payments:
            payments_sheet.append([year, month, "Μερική πληρωμή", float(amount), "", "", created_at])

        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions if sheet.max_row > 1 else None
            for column_cells in sheet.columns:
                width = min(max(max(len(str(cell.value or "")) for cell in column_cells) + 2, 12), 42)
                sheet.column_dimensions[column_cells[0].column_letter].width = width

        workbook.save(file_path)
        QMessageBox.information(self, "Ολοκληρώθηκε", f"Η εξαγωγή για τον/την {employee_name} αποθηκεύτηκε.")

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
                f"{record[3]}-{record[4]} | Break:{record[5]} min | "
                f"Actual hours:{record[6]}"
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
            if employee not in totals:
                totals[employee] = 0
            totals[employee] += float(record[6] or 0)

        names = list(totals.keys())
        amounts = list(totals.values())

        plt.figure(figsize=(10, 6))
        plt.bar(names, amounts)
        plt.title("Πραγματικές ώρες ανά υπάλληλο")
        plt.xlabel("Υπάλληλος")
        plt.ylabel("Ώρες")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()


app = QApplication(sys.argv)

window = PayrollApp()
window.show()

sys.exit(app.exec())
