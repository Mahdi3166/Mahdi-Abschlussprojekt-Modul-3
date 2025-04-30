import mysql.connector
class DatabaseHandler:
    def __init__(self, host, user, password, database):
        # Verbindung zur Datenbank herstellen
        self.connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database )
         # Cursor erstellen, um SQL-Abfragen auszuführen
        self.cursor = self.connection.cursor()
    def __del__(self):
        self.connection.close()
    def close_connection(self):
        # Verbindung zur Datenbank schließen
        self.connection.close()
    def get_data(self, query):
        # SQL-Abfrage ausführen
        self.cursor.execute(query)
        # Ergebnisse abrufen (falls vorhanden)
        result = self.cursor.fetchall()
        return result
    def insert_data(self, query, insert_data):
        # Daten in die Datenbank einfügen
        self.cursor.execute(query, insert_data)
        # Änderungen bestätigen
        self.connection.commit()
    def change_data(self, query, change_data):
        # Datensätze aus der Datenbank löschen
        self.cursor.execute(query, change_data)
        # Änderungen bestätigen
        self.connection.commit()