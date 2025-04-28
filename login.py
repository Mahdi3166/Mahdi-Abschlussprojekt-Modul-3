from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QMessageBox
)
from database import DatabaseHandler  # DatabaseHandler senin verdiğin sınıf

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Datenbankanmeldung")
        self.setFixedSize(300, 200)

        self.db_handler = None

        # Layout
        layout = QVBoxLayout()

        # Database Auswahl
        self.db_label = QLabel("Database:")
        self.db_select = QComboBox()
        self.db_select.addItems(["AD", "interessenten_modul3"])
        layout.addWidget(self.db_label)
        layout.addWidget(self.db_select)

        # Benutzername
        self.user_label = QLabel("User name:")
        self.user_input = QLineEdit()
        layout.addWidget(self.user_label)
        layout.addWidget(self.user_input)

        # Passwort
        self.pass_label = QLabel("Password:")
        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.pass_label)
        layout.addWidget(self.pass_input)

        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")
        self.ok_button.clicked.connect(self.try_login)
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def try_login(self):
        db = self.db_select.currentText()
        user = self.user_input.text()
        password = self.pass_input.text()

        try:
            self.db_handler = DatabaseHandler(
                host="mariadb",  # veya senin ortamına göre ayarla
                user=user,
                password=password,
                database=db
            )
            QMessageBox.information(self, "Erfolg", "Verbindung erfolgreich!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Verbindung fehlgeschlagen:\n{e}")
            self.db_handler = None

    def get_db_handler(self):
        return self.db_handler

 