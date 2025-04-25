from editaduser_TN import EditADUserWindow
from database import DatabaseHandler  
from login import LoginDialog
from editaduser_TN import EditADUserWindow
import sys
import os
import csv
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QDockWidget,
    QToolBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QMessageBox,
    QAbstractItemView,
    QFileDialog
)
           
class MainWindow(QMainWindow):
 
    def __init__(self):
        super().__init__()
       
        self.mainmenue = {1: "&Datei", 2: "&Active Directory", 4:"&Hilfe"}
        self.menueoptions= {11:"Import von CSV", 12:"Transfer nach AD",  13: "Einloggen", 14:"Ausloggen", 0: "separator", 19: "&Beenden", 21:"Benutzer bearbeiten", 22:"Lösche AD-User",  23:"Inaktiv AD-User", 41:"&Über", 42:"&Hilfe"}
        self.toolbarbuttons=  {13: "Einloggen", 11:"Import von CSV", 12:"Transfer nach AD",  0: "separator", 21:"Benutzer bearbeiten", 22:"Lösche AD-User", 23:"Inaktiv AD-User",  0: "separator", 42:"&Hilfe"}
        self.initUI()
       
    def initUI(self):
        self.setWindowTitle("myAdmin Center")
        self.setWindowIcon(QIcon(".\\images\\logo-zm.png"))
        # Menüleiste erstellen für jeden Eintrag im Dictionary self.mainmenu. Dort steht die Beschreibung, die für das Menü und
        # für ToolTip Texte verwendet wird. zusätzlich ein numerischer Code, der abgefragt werden kann
        menubar = self.menuBar()
        for menu_id, menu_title in self.mainmenue.items():
            menu = menubar.addMenu(menu_title)
            for action_id, action_title in self.menueoptions.items():
                if action_id ==0:
                    menu.addSeparator()
                else:
                    if action_id // 10 == menu_id:
                        action=QAction(action_title, self)
                        action.setProperty("command", (action_id, action_title))
                        action.triggered.connect(self.menue_clicked)
                        menu.addAction(action)
        # Toolbar erstellen auf demselben Weg, mit dem Dictionary self.toolbarbuttoms
        toolbar = QToolBar("Hauptwerkzeugleiste")
        self.addToolBar(toolbar)
       
        for command, caption in self.toolbarbuttons.items():
            if command == 0:
                toolbar.addSeparator()
            else:
                btn = QPushButton()
                icon = ".\\images\\tb_" + str(command)+ ".png"
                if os.path.exists(icon):
                    btn.setIcon(QIcon(icon))
                    btn.setIconSize(QSize(32, 32))
                    btn.setToolTip(caption)
                else:
                    btn.setText(caption)
                btn.setProperty("command", (command, caption))
                btn.clicked.connect(self.menue_clicked)
                toolbar.addWidget(btn)  
       
        # Statusbar
        self.statusBar().showMessage("Ausgeloggt")
       
        # DockWidget anlegen
        self.dock = QDockWidget("Dock", self)
        self.dock.setWidget(QTextEdit("Zeigt Hilfe"))
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.dock)
        self.dock.setVisible(False)
       
        # Zentrales Widget + Layout
        central_widget = QWidget(self)
        # Zentrales Widget dem QMainWindow zuweisen
        self.setCentralWidget(central_widget)
        central_layout = QVBoxLayout(central_widget)
       
       
 
 
        # TableWidget für die Interessenten aus der Datenbank
        self.table_interessenten = QTableWidget()
        self.table_interessenten.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_interessenten.doubleClicked.connect(self.editaduser)        
        self.table_interessenten.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_interessenten.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
             
        # BEGINN TabWidget Demo - Löschen und durch Zugriff auf die Datenbank ersetzen
       
       
        # ENDE TabWidget Demo
 
        central_layout.addWidget(self.table_interessenten)
        self.resize(800, 600)  # Individuell zu setzen. Notwendig, wenn das TableWidget zu Beginn nicht sichtbar ist
        self.show()
 
    def editaduser(self):
        selection = self.table_interessenten.selectedItems()
        if selection:
            row = selection[0].row()
            userid = self.table_interessenten.item(row, 0).text()
            self.editaduserwindow = EditADUserWindow(self.menueoptions[21], userid, self.db_handler)
            self.editaduserwindow.setWindowModality(Qt.WindowModality.ApplicationModal)
            self.editaduserwindow.show()
        else:
            QMessageBox.warning(self, "Fehler", "Kein Eintrag ausgewählt!")
 
   
    def menue_clicked(self):                # slot für Menü und Toolbar. Ein Slot für alle Widgets
        men = self.sender()                 # Welches Widget hat das Ergeignis ausgelöst?
        print(f"Menu {men.property("command")} selected")
        # Command (int) und Beschreibung werden in der property "command" geliefert
        # Abfragen auf die ID des Kommandos
        match men.property("command")[0]:
            case 21:                        # edit AD-User
                self.editaduser()
            case 42:
                self.menue_help_help()      # Aufruf der Funktion für den Menüpunkt "Help->Help"
            case 41:
                self.menue_help_about()     # Aufruf der Funktion für den Menüpunkt "Help->About" (keine User Story)
            case 13:  # Einloggen
                self.menu_login()
            case 14:
                self.logout_database()
 
            case 11:  # CSV-Import
                self.menue_csv_import()
 
#           case 11:                        # Beispiel für weitere Funktionsaufrufe
#               self.menue_csv_import()                
#           usw.
#
#buradan sonra
    def menu_login(self):
        dlg = LoginDialog(self)
        if dlg.exec():
            self.db_handler = dlg.get_db_handler()
            self.statusBar().showMessage("Eingeloggt")
            self.load_ad_users() #  veri yükleme çağrısı
    def logout_database(self):
        if self.db_handler:
            self.db_handler.close_connection()
            self.db_handler = None
        self.table_interessenten.clear()
        self.table_interessenten.setRowCount(0)
        self.table_interessenten.setColumnCount(0)
        self.statusBar().showMessage("Ausgeloggt")
 
           
    def load_ad_users(self):
        if not hasattr(self, 'db_handler') or self.db_handler is None:
            QMessageBox.warning(self, "Fehler", "Keine Datenbankverbindung!")
            return
 
        try:
            query = "SELECT * FROM view_aduser_details"
            results = self.db_handler.get_data(query)
 
            # Tabloyu sıfırla
            self.table_interessenten.clear()
            self.table_interessenten.setRowCount(0)
 
            # Başlıkları ayarla (kolon adları otomatik değilse elle yaz)
            headers = [desc[0] for desc in self.db_handler.cursor.description]
            self.table_interessenten.setColumnCount(len(headers))
            self.table_interessenten.setHorizontalHeaderLabels(headers)
 
            # Verileri tabloya ekle
            for row_num, row_data in enumerate(results):
                self.table_interessenten.insertRow(row_num)
                for col_num, value in enumerate(row_data):
                    self.table_interessenten.setItem(row_num, col_num, QTableWidgetItem(str(value)))
 
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Laden der Daten:\n{e}")
    def menue_csv_import(self):
        if not hasattr(self, 'db_handler') or self.db_handler is None:
            QMessageBox.warning(self, "Fehler", "Bitte zuerst einloggen!")
            return
 
        file_path, _ = QFileDialog.getOpenFileName(self, "CSV-Datei auswählen", "", "CSV-Dateien (*.csv)")
        if not file_path:
            return
 
        try:
            with open(file_path, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
 
                for row in reader:
                    firstname = row['firstname']
                    lastname = row['lastname']
                    username = (firstname[0] + lastname).lower()
                    email = f"{firstname.lower()}.{lastname.lower()}@M-zukunftsmotor.local"
                    kurs_id = int(row['kurs'])
                    status_id = int(row['status_id_fk'])
 
                    check_query = "SELECT id_pk FROM aduser WHERE username = '{username}'"
                    existing = self.db_handler.get_data(check_query)
 
                    if existing:
                        update_query = """
                            UPDATE aduser SET firstname=%s, lastname=%s, email=%s, phone=%s,
                                department=%s, street=%s, city=%s, city_code=%s, postalcode=%s,
                                status_id_fk=%s, ou_id_fk=%s, modified=NOW()
                            WHERE username=%s
                        """
                        values = (
                            firstname, lastname, email, row['phone'], row['abteilung'],
                            row['street'], row['city'], row['city_code'], row['postalcode'],
                            status_id, kurs_id, username
                        )
                        self.db_handler.change_data(update_query, values)
                    else:
                        insert_query = """
                            INSERT INTO aduser (firstname, lastname, username, email, phone, department, street,
                                city, city_code, postalcode, status_id_fk, ou_id_fk)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """
                        values = (
                            firstname, lastname, username, email, row['phone'], row['abteilung'],
                            row['street'], row['city'], row['city_code'], row['postalcode'],
                            status_id, kurs_id
                        )
                        self.db_handler.insert_data(insert_query, values)
 
            QMessageBox.information(self, "Erfolg", "CSV-Import abgeschlossen!")
            self.load_ad_users()
 
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Fehler beim Import:\n{e}")
 
   
#buradan önce
    def menue_help_about(self):
        print("Missing function!")
    def menue_help_help(self):
        self.dock.setVisible(True)
#           usw.
 
 
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())
 
if __name__ == "__main__":
    main()
 
 
 