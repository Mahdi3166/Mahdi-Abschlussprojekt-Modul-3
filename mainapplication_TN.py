# Import custom modules for editing AD users, database handling, and login dialog
from editaduser_TN import EditADUserWindow
from database import DatabaseHandler
from login import LoginDialog
# Import standard Python libraries for system operations, file handling, CSV processing, and file copying
import sys, os, csv, shutil
# Import PyQt6 modules for GUI components and functionality
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import *

# MainWindow class inherits from QMainWindow to create the main application window
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Define main menu structure with IDs and titles
        # - 1: File menu, 2: Active Directory menu, 4: Help menu
        self.mainmenue = {1: "&Datei", 2: "&Active Directory", 4: "&Hilfe"}
        # Define menu options with IDs and captions
        # - IDs are structured to map to menus (e.g., 11-19 for File, 21-23 for AD, 41-42 for Help)
        # - 0 represents a separator
        self.menueoptions = {
            11: "Import von CSV", 12: "Transfer nach AD", 13: "Einloggen", 14: "Ausloggen",
            0: "separator", 19: "&Beenden", 21: "Benutzer bearbeiten", 22: "Lösche AD-User",
            23: "Inaktiv AD-User", 41: "&Über", 42: "&Hilfe"
        }
        # Define toolbar buttons with IDs and captions
        # - Matches some menu options for quick access
        # - 0 represents a separator
        self.toolbarbuttons = {
            13: "Einloggen", 11: "Import von CSV", 12: "Transfer nach AD",
            0: "separator", 21: "Benutzer bearbeiten", 22: "Lösche AD-User", 23: "Inaktiv AD-User",
            0: "separator", 42: "&Hilfe"
        }

        # Initialize the user interface
        self.initUI()

    def initUI(self):
        # Set window title and icon
        self.setWindowTitle("myAdmin Center")
        self.setWindowIcon(QIcon(".\\images\\logo-zm.png"))

        # Create and populate the menu bar
        menubar = self.menuBar()
        # Iterate through main menu items
        for menu_id, menu_title in self.mainmenue.items():
            menu = menubar.addMenu(menu_title)  # Add menu to menubar
            # Iterate through menu options
            for action_id, action_title in self.menueoptions.items():
                if action_id == 0:
                    menu.addSeparator()  # Add separator if ID is 0
                elif action_id // 10 == menu_id:
                    # Create QAction for menu item if it belongs to the current menu
                    action = QAction(action_title, self)
                    # Store command ID and title as property for handling clicks
                    action.setProperty("command", (action_id, action_title))
                    # Connect action to menu click handler
                    action.triggered.connect(self.menue_clicked)
                    menu.addAction(action)

        # Create and populate the toolbar
        toolbar = QToolBar("Hauptwerkzeugleiste")
        self.addToolBar(toolbar)
        # Iterate through toolbar buttons
        for command, caption in self.toolbarbuttons.items():
            if command == 0:
                toolbar.addSeparator()  # Add separator if ID is 0
            else:
                # Create a QPushButton for the toolbar
                btn = QPushButton()
                # Set icon if available, otherwise use text
                icon = f".\\images\\tb_{command}.png"
                if os.path.exists(icon):
                    btn.setIcon(QIcon(icon))
                    btn.setIconSize(QSize(32, 32))
                    btn.setToolTip(caption)
                else:
                    btn.setText(caption)
                # Store command ID and caption as property
                btn.setProperty("command", (command, caption))
                # Connect button to menu click handler
                btn.clicked.connect(self.menue_clicked)
                toolbar.addWidget(btn)

        # Initialize status bar with default message
        self.statusBar().showMessage("Ausgeloggt")

        # Create a dock widget for help content
        self.dock = QDockWidget("Dock", self)
        # Add a QTextBrowser to display HTML help content
        self.dock.setWidget(QTextBrowser())
        # Set HTML content describing the application and its features
        self.dock.widget().setHtml(("""
    <h2 style='color: navy;'>Herzlich willkommen im <i>myAdmin Center</i>!</h2>
    <p>Dieses Verwaltungswerkzeug unterstützt Sie bei der effizienten und sicheren Verwaltung von Benutzerdaten. Es richtet sich vor allem an Administratorinnen und Administratoren, die regelmäßig mit dem Active Directory (AD) arbeiten.</p>

    <p>Mit dem <b>myAdmin Center</b> haben Sie folgende Möglichkeiten:</p>
    <ul>
        <li>Importieren Sie Benutzerdaten komfortabel aus vorbereiteten <b>CSV-Dateien</b>.</li>
        <li>Bearbeiten Sie bestehende Benutzerinformationen direkt in der Benutzeroberfläche.</li>
        <li>Übertragen Sie die gepflegten Daten automatisiert in Ihre Active Directory-Umgebung.</li>
        <li>Verwalten Sie Ihre AD-Konten durch Deaktivieren oder endgültiges Löschen von Benutzern.</li>
    </ul>

    <p>Verwenden Sie dazu die Optionen in der <b>Menüleiste</b> oder in der <b>Toolbar</b>, um alle Funktionen intuitiv aufzurufen.</p>

    <p style='color: darkred;'><b>Wichtiger Hinweis:</b> Bitte stellen Sie sicher, dass Sie sich erfolgreich eingeloggt haben, bevor Sie mit der Bearbeitung oder Übertragung von Daten beginnen. Ohne gültige Anmeldung sind die meisten Aktionen aus Sicherheitsgründen gesperrt.</p>

    <p>Wir wünschen Ihnen viel Erfolg bei Ihrer Arbeit mit dem myAdmin Center!</p>
"""))
        # Add dock to the left side and hide it initially
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.dock)
        self.dock.setVisible(False)

        # Create central widget and layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        central_layout = QVBoxLayout(central_widget)

        # Create table for displaying AD user data
        self.table_interessenten = QTableWidget()
        # Disable editing to prevent direct modifications
        self.table_interessenten.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        # Connect double-click to edit user function
        self.table_interessenten.doubleClicked.connect(self.editaduser)
        # Select entire rows and allow single selection
        self.table_interessenten.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_interessenten.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        # Add table to central layout
        central_layout.addWidget(self.table_interessenten)
        # Set window size and show the window
        self.resize(800, 600)
        self.show()

    def editaduser(self):
        # Handle double-click on table to edit a user
        selection = self.table_interessenten.selectedItems()
        if selection:
            # Get row and user ID from the first column
            row = selection[0].row()
            userid = self.table_interessenten.item(row, 0).text()
            # Open edit user window with user ID and database handler
            self.editaduserwindow = EditADUserWindow(self.menueoptions[21], userid, self.db_handler)
            # Make window modal to block other interactions
            self.editaduserwindow.setWindowModality(Qt.WindowModality.ApplicationModal)
            self.editaduserwindow.show()
        else:
            # Show warning if no row is selected
            QMessageBox.warning(self, "Fehler", "Kein Eintrag ausgewählt!")

    def delete_ad_user(self):
        # Delete selected AD user from database
        selection = self.table_interessenten.selectedItems()
        if not selection:
            # Show warning if no row is selected
            QMessageBox.warning(self, "Fehler", "Kein Eintrag ausgewählt!")
            return

        # Get row and user ID
        row = selection[0].row()
        userid = self.table_interessenten.item(row, 0).text()

        # Confirm deletion with user
        confirm = QMessageBox.question(
            self, "Bestätigung",
            f"Soll der Benutzer mit ID {userid} wirklich gelöscht werden?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            try:
                # Execute delete query
                query = "DELETE FROM aduser WHERE id_pk = %s"
                self.db_handler.change_data(query, (userid,))
                # Show success message
                QMessageBox.information(self, "Erfolg", f"Benutzer mit ID {userid} wurde gelöscht.")
                # Refresh table
                self.load_ad_users()
            except Exception as e:
                # Show error message if deletion fails
                QMessageBox.critical(self, "Fehler", f"Fehler beim Löschen:\n{e}")

    def deactivate_ad_user(self):
        # Deactivate selected AD user by updating status
        selection = self.table_interessenten.selectedItems()
        if not selection:
            # Show warning if no row is selected
            QMessageBox.warning(self, "Fehler", "Kein Eintrag ausgewählt!")
            return

        # Get row and user ID
        row = selection[0].row()
        userid = self.table_interessenten.item(row, 0).text()

        # Confirm deactivation with user
        confirm = QMessageBox.question(
            self, "Bestätigung",
            f"Soll der Benutzer mit ID {userid} deaktiviert werden?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            try:
                # Update status to inactive (status_id_fk = 2)
                query = "UPDATE aduser SET status_id_fk = 2 WHERE id_pk = %s"
                self.db_handler.change_data(query, (userid,))
                # Show success message
                QMessageBox.information(self, "Erfolg", f"Benutzer mit ID {userid} wurde deaktiviert.")
                # Refresh table
                self.load_ad_users()
            except Exception as e:
                # Show error message if deactivation fails
                QMessageBox.critical(self, "Fehler", f"Fehler beim Deaktivieren:\n{e}")

    def transfer_to_ad(self):
        # Export AD user data to CSV and copy to network path
        if not hasattr(self, 'db_handler') or self.db_handler is None:
            # Check if logged in
            QMessageBox.warning(self, "Fehler", "Bitte zuerst einloggen!")
            return

        try:
            # Fetch all user details from view
            query = "SELECT * FROM view_aduser_details"
            results = self.db_handler.get_data(query)
            # Get column headers
            headers = [desc[0] for desc in self.db_handler.cursor.description]
            # Write data to local CSV file
            local_file = "ad_export.csv"
            with open(local_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(headers)
                writer.writerows(results)
            # Copy file to network path
            network_path = r"\\Admin-Server\Logging\ad_export.csv"
            shutil.copy(local_file, network_path)
            # Show success message
            QMessageBox.information(self, "Erfolg", "Daten erfolgreich übertragen.")
        except Exception as e:
            # Show error message if transfer fails
            QMessageBox.critical(self, "Fehler", f"Fehler beim Transfer:\n{e}")

    def menue_clicked(self):
        # Handle menu and toolbar button clicks
        men = self.sender()
        # Match command ID to corresponding function
        match men.property("command")[0]:
            case 21: self.editaduser()  # Edit user
            case 22: self.delete_ad_user()  # Delete user
            case 23: self.deactivate_ad_user()  # Deactivate user
            case 12: self.transfer_to_ad()  # Transfer to AD
            case 42: self.menue_help_help()  # Show help
            case 41: self.menue_help_about()  # Show about
            case 13: self.menu_login()  # Login
            case 14: self.logout_database()  # Logout
            case 11: self.menue_csv_import()  # Import CSV

    def menu_login(self):
        # Open login dialog and establish database connection
        dlg = LoginDialog(self)
        if dlg.exec():
            # Get database handler from dialog
            self.db_handler = dlg.get_db_handler()
            # Update status bar
            self.statusBar().showMessage("Eingeloggt")
            # Load user data
            self.load_ad_users()

    def logout_database(self):
        # Close database connection and clear table
        if self.db_handler:
            self.db_handler.close_connection()
            self.db_handler = None
        # Clear table contents
        self.table_interessenten.clear()
        self.table_interessenten.setRowCount(0)
        self.table_interessenten.setColumnCount(0)
        # Update status bar
        self.statusBar().showMessage("Ausgeloggt")

    def load_ad_users(self):
        # Load AD user data into table
        if not hasattr(self, 'db_handler') or self.db_handler is None:
            # Check if logged in
            QMessageBox.warning(self, "Fehler", "Keine Datenbankverbindung!")
            return

        try:
            # Fetch user details from view
            query = "SELECT * FROM view_aduser_details"
            results = self.db_handler.get_data(query)
            # Clear existing table
            self.table_interessenten.clear()
            self.table_interessenten.setRowCount(0)
            # Set column headers
            headers = [desc[0] for desc in self.db_handler.cursor.description]
            self.table_interessenten.setColumnCount(len(headers))
            self.table_interessenten.setHorizontalHeaderLabels(headers)
            # Populate table with data
            for row_num, row_data in enumerate(results):
                self.table_interessenten.insertRow(row_num)
                for col_num, value in enumerate(row_data):
                    self.table_interessenten.setItem(row_num, col_num, QTableWidgetItem(str(value)))
        except Exception as e:
            # Show error message if loading fails
            QMessageBox.critical(self, "Fehler", f"Fehler beim Laden:\n{e}")

    def menue_csv_import(self):
        # Import user data from CSV file
        if not hasattr(self, 'db_handler') or self.db_handler is None:
            # Check if logged in
            QMessageBox.warning(self, "Fehler", "Bitte zuerst einloggen!")
            return

        # Open file dialog to select CSV
        file_path, _ = QFileDialog.getOpenFileName(self, "CSV-Datei auswählen", "", "CSV-Dateien (*.csv)")
        if not file_path:
            return

        try:
            # Read CSV file
            with open(file_path, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    # Extract and process user data
                    firstname = row['firstname']
                    lastname = row['lastname']
                    username = (firstname[0] + lastname).lower()
                    email = f"{firstname.lower()}.{lastname.lower()}@M-zukunftsmotor.local"
                    kurs_id = int(row['kurs'])
                    status_id = int(row['status_id_fk'])

                    # Check if user already exists
                    check_query = f"SELECT id_pk FROM aduser WHERE username = '{username}'"
                    existing = self.db_handler.get_data(check_query)

                    if existing:
                        # Update existing user
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
                        # Insert new user
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

            # Show success message and refresh table
            QMessageBox.information(self, "Erfolg", "CSV-Import abgeschlossen!")
            self.load_ad_users()

        except Exception as e:
            # Show error message if import fails
            QMessageBox.critical(self, "Fehler", f"Fehler beim Import:\n{e}")

    def menue_help_about(self):
        # Placeholder for about dialog (not implemented)
        print("Missing function!")

    def menue_help_help(self):
        # Show help dock widget
        self.dock.setVisible(True)

# Main function to run the application
def main():
    # Create QApplication instance
    app = QApplication(sys.argv)
    # Create and show main window
    window = MainWindow()
    # Start event loop and exit on close
    sys.exit(app.exec())

# Run main function if script is executed directly
if __name__ == "__main__":
    main()