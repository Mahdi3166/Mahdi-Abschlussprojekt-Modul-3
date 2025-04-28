from PyQt6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout,
    QPushButton, QMessageBox, QComboBox, QFormLayout, QApplication, QMainWindow
)
from PyQt6.QtCore import QDateTime, Qt
import datetime

class EditADUserWindow(QWidget):
    def __init__(self, caption, userid, db_handler):
        super().__init__()
        self.setWindowTitle(caption)
        self.userid = userid
        self.db_handler = db_handler

        # Fields
        self.firstname = QLineEdit()
        self.firstname.setReadOnly(True)
        self.lastname = QLineEdit()
        self.lastname.setReadOnly(True)
        self.phone = QLineEdit()
        self.department = QLineEdit()
        self.street = QLineEdit()
        self.postalcode = QLineEdit()
        self.city = QLineEdit()
        self.kuerzel = QComboBox()  # NEU: Kürzel jetzt als ComboBox

        self.status = QComboBox()
        self.course = QComboBox()
        self.created_label = QLabel()
        self.modified_label = QLabel()

        # Form Layout
        form_layout = QFormLayout()
        form_layout.addRow("Vorname:", self.firstname)
        form_layout.addRow("Nachname:", self.lastname)
        form_layout.addRow("Telefon:", self.phone)
        form_layout.addRow("Abteilung:", self.department)
        form_layout.addRow("Straße:", self.street)
        form_layout.addRow("PLZ:", self.postalcode)
        form_layout.addRow("Ort:", self.city)
        form_layout.addRow("Kürzel:", self.kuerzel)
        form_layout.addRow("Status:", self.status)
        form_layout.addRow("Kurs:", self.course)
        form_layout.addRow("Erstellt:", self.created_label)
        form_layout.addRow("Letzte Änderung:", self.modified_label)

        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Abbrechen")
        ok_button.clicked.connect(self.save_changes)
        cancel_button.clicked.connect(self.close)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)

        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

        self.load_user_data()

    def to_local_time(self, dt):
        if isinstance(dt, datetime.datetime):
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=datetime.timezone.utc)
            return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        return str(dt)

    def load_user_data(self):
        query = f"""
            SELECT firstname, lastname, phone, department, street, postalcode, city, city_code,
                   status_id_fk, ou_id_fk, created, modified
            FROM aduser
            WHERE id_pk = {self.userid}
        """
        try:
            result = self.db_handler.get_data(query)
            if result:
                (
                    firstname, lastname, phone, department, street, postalcode, city, city_code,
                    status_id, ou_id, created, modified
                ) = result[0]

                self.firstname.setText(firstname)
                self.lastname.setText(lastname)
                self.phone.setText(phone)
                self.department.setText(department)
                self.street.setText(street)
                self.postalcode.setText(postalcode)
                self.city.setText(city)

                self.created_label.setText(self.to_local_time(created))
                self.modified_label.setText(self.to_local_time(modified))

                # Kürzel ComboBox füllen mit allen vorhandenen city_code
                city_codes = self.db_handler.get_data("SELECT DISTINCT city_code FROM aduser WHERE city_code IS NOT NULL")
                for (code,) in city_codes:
                    if code:
                        self.kuerzel.addItem(code)

                # Aktuelles Kürzel setzen
                index = self.kuerzel.findText(city_code)
                if index >= 0:
                    self.kuerzel.setCurrentIndex(index)

                # Status ComboBox
                statuses = self.db_handler.get_data("SELECT id_pk, bezeichnung FROM aduser_status")
                for sid, name in statuses:
                    self.status.addItem(name, sid)
                index = self.status.findData(status_id)
                if index >= 0:
                    self.status.setCurrentIndex(index)

                # Kurs ComboBox
                courses = self.db_handler.get_data("SELECT id_pk, name FROM adou")
                for cid, name in courses:
                    self.course.addItem(name, cid)
                index = self.course.findData(ou_id)
                if index >= 0:
                    self.course.setCurrentIndex(index)

            else:
                QMessageBox.warning(self, "Fehler", "Benutzer nicht gefunden.")
                self.close()

        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Laden:\n{e}")
            self.close()

    def save_changes(self):
        try:
            query = """
                UPDATE aduser
                SET phone=%s, department=%s, street=%s, postalcode=%s, city=%s, city_code=%s,
                    status_id_fk=%s, ou_id_fk=%s, modified=NOW()
                WHERE id_pk=%s
            """
            values = (
                self.phone.text(),
                self.department.text(),
                self.street.text(),
                self.postalcode.text(),
                self.city.text(),
                self.kuerzel.currentText(),  # NEU: Kürzel aus ComboBox speichern
                self.status.currentData(),
                self.course.currentData(),
                self.userid,
            )
            self.db_handler.change_data(query, values)

            QMessageBox.information(self, "Erfolg", "Benutzerdaten gespeichert.")

            # MainWindow aktualisieren
            for widget in QApplication.topLevelWidgets():
                if isinstance(widget, QMainWindow):
                    if hasattr(widget, 'load_ad_users'):
                        widget.load_ad_users()
                    break

            self.close()

        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Speichern:\n{e}")
