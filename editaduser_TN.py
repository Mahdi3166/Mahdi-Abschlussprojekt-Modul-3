from PyQt6.QtWidgets import (
    QWidget, QLabel, QLineEdit, QVBoxLayout, QHBoxLayout,
    QPushButton, QComboBox, QMessageBox
)
from PyQt6.QtCore import Qt
from database import DatabaseHandler


class EditADUserWindow(QWidget):
    def __init__(self, title, user_id, db_handler):
        super().__init__()
        self.setWindowTitle(title)
        self.user_id = user_id
        self.db_handler = db_handler

        main_layout = QVBoxLayout(self)

        # Formfelder
        self.firstname = QLineEdit()
        self.lastname = QLineEdit()
        self.phone = QLineEdit()
        self.department = QLineEdit()
        self.street = QLineEdit()
        self.zipcode = QLineEdit()
        self.city = QLineEdit()
        self.status = QComboBox()
        self.course = QComboBox()
        self.created_label = QLabel()
        self.modified_label = QLabel()

        # Einheitliche Breite für Felder
        self._set_fixed_sizes()

        # Zeilenaufbau
        main_layout.addLayout(self._row("Vorname:", self.firstname))
        main_layout.addLayout(self._row("Nachname:", self.lastname))
        main_layout.addLayout(self._row("Telefon:", self.phone))
        main_layout.addLayout(self._row("Abteilung:", self.department))
        main_layout.addLayout(self._row("Straße:", self.street))
        main_layout.addLayout(self._row("PLZ:", self.zipcode))
        main_layout.addLayout(self._row("Ort:", self.city))
        main_layout.addLayout(self._row("Status:", self.status))
        main_layout.addLayout(self._row("Kurs:", self.course))
        main_layout.addLayout(self._row("erstellt:", self.created_label))
        main_layout.addLayout(self._row("letzte Änderung:", self.modified_label))

        # Buttons nebeneinander
        button_layout = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.cancel_btn = QPushButton("Abbrechen")
        self.ok_btn.clicked.connect(self.save_changes)
        self.cancel_btn.clicked.connect(self.close)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)
        self.load_user_data()

    def _set_fixed_sizes(self):
        self.label_width = 100
        self.input_width = 200

        # Einheitliche Breite für alle Eingabefelder
        for widget in [
            self.firstname, self.lastname, self.phone, self.department,
            self.street, self.zipcode, self.city, self.status, self.course
        ]:
            widget.setFixedWidth(self.input_width)

    def _row(self, label_text, widget):
        row = QHBoxLayout()
        label = QLabel(label_text)
        label.setFixedWidth(self.label_width)
        label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)  # linksbündig
        row.addWidget(label)
        row.addWidget(widget)
        return row

    def load_user_data(self):
        query = f"""
            SELECT firstname, lastname, phone, department, street, postalcode, city,
                   status_id_fk, ou_id_fk, created, modified
            FROM aduser WHERE id_pk = {self.user_id}
        """
        result = self.db_handler.get_data(query)
        if not result:
            QMessageBox.critical(self, "Fehler", "Benutzer nicht gefunden.")
            self.close()
            return

        (
            firstname, lastname, phone, department, street, plz, city,
            status_id, ou_id, created, modified
        ) = result[0]

        self.firstname.setText(firstname)
        self.lastname.setText(lastname)
        self.phone.setText(phone)
        self.department.setText(department)
        self.street.setText(street)
        self.zipcode.setText(plz)
        self.city.setText(city)
        self.created_label.setText(str(created))
        self.modified_label.setText(str(modified))

        # Status laden
        self.status.clear()
        status_list = self.db_handler.get_data("SELECT id_pk, bezeichnung FROM aduser_status")
        for id_, name in status_list:
            self.status.addItem(name, id_)
            if id_ == status_id:
                self.status.setCurrentText(name)

        # Kurs laden
        self.course.clear()
        course_list = self.db_handler.get_data("SELECT id_pk, name FROM adou")
        for id_, name in course_list:
            self.course.addItem(name, id_)
            if id_ == ou_id:
                self.course.setCurrentText(name)

    def save_changes(self):
        try:
            query = """
                UPDATE aduser SET firstname=%s, lastname=%s, phone=%s, department=%s,
                    street=%s, postalcode=%s, city=%s,
                    status_id_fk=%s, ou_id_fk=%s, modified=NOW()
                WHERE id_pk=%s
            """
            values = (
                self.firstname.text(),
                self.lastname.text(),
                self.phone.text(),
                self.department.text(),
                self.street.text(),
                self.zipcode.text(),
                self.city.text(),
                self.status.currentData(),
                self.course.currentData(),
                self.user_id
            )
            self.db_handler.change_data(query, values)
            QMessageBox.information(self, "Erfolg", "Daten erfolgreich gespeichert.")
            if self.parent():
                try:
                    self.parent().load_ad_users()
                except:
                    pass
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Speichern:\n{e}")
