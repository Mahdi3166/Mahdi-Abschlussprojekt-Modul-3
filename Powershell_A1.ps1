# Dieses PowerShell-Skript importiert Benutzerdaten aus einer CSV-Datei und verwaltet Active Directory (AD)-Benutzer.
# Es erstellt, aktualisiert, deaktiviert oder löscht Benutzer basierend auf dem Status_code in der CSV-Datei.
# Es prüft und erstellt OUs, verschiebt Benutzer zwischen OUs und protokolliert alle Aktionen in einer Log-Datei.

# Parameterdefinition
# - $csvFolderPath: Pfad zum Ordner mit der CSV-Datei, Standardwert ist "\\Admin-Server\logging"
param(
    [string]$csvFolderPath = "\\Admin-Server\logging"
)

# Einstellungen
# - $domain: AD-Domäne, in der die Benutzer verwaltet werden
$domain = "M-zukunftsmotor.local"
# - $baseOU: Basis-OU-Pfad, unter dem untergeordnete OUs erstellt werden
$baseOU = "OU=M-zukunftsmotor,DC=M-zukunftsmotor,DC=local"
# - $logFile: Pfad zur Log-Datei für Protokollierung
$logFile = "$csvFolderPath\user-import-log.txt"

# Neueste CSV-Datei im angegebenen Ordner finden
# - Get-ChildItem sucht nach CSV-Dateien im Ordner
# - -Filter *.csv beschränkt die Suche auf CSV-Dateien
# - Sort-Object LastWriteTime -Descending sortiert nach Änderungsdatum (neueste zuerst)
# - Select-Object -First 1 wählt die neueste Datei aus
$csvFile = Get-ChildItem -Path $csvFolderPath -Filter *.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1
# Prüfen, ob eine CSV-Datei gefunden wurde
if (-not $csvFile) {
    # Fehlermeldung ausgeben und Skript beenden, wenn keine Datei gefunden wurde
    Write-Host " Keine gültige CSV-Datei gefunden im Pfad: $csvFolderPath"
    exit
}

# CSV-Datei laden
# - Import-Csv liest die CSV-Datei ein
# - -Path $csvFile.FullName gibt den vollständigen Pfad der Datei an
# - -Delimiter ',' definiert das Trennzeichen (Komma)
$users = Import-Csv -Path $csvFile.FullName -Delimiter ','

# Logfunktion definieren
# - Protokolliert Nachrichten mit Zeitstempel in Datei und Konsole
function Log($text) {
    # Erstelle Zeitstempel im Format "YYYY-MM-DD HH:MM:SS"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Schreibe Nachricht mit Zeitstempel in Log-Datei (UTF-8)
    "$timestamp - $text" | Out-File -FilePath $logFile -Append -Encoding utf8
    # Gib Nachricht in Konsole aus
    Write-Host $text
}

# Schleife durch alle Benutzer in der CSV-Datei
foreach ($user in $users) {
    # Pflichtfelder prüfen
    # - Überprüfe, ob Vorname, Nachname, Kurs und Status_code vorhanden sind
    if (-not $user.Vorname -or -not $user.Nachname -or -not $user.Kurs -or -not $user.Status_code) {
        # Protokolliere Warnung und überspringe Benutzer, wenn Pflichtfelder fehlen
        Log " WARNUNG: Fehlende Pflichtfelder für $($user.Vorname) $($user.Nachname): Kurs oder Status_code"
        continue
    }

    # Benutzerdaten vorbereiten
    # - $username: Kombination aus erstem Buchstaben des Vornamens und Nachname, kleingeschrieben, max. 20 Zeichen
    $username = ($user.Vorname.Substring(0,1) + $user.Nachname).ToLower().Substring(0, [Math]::Min(20, ($user.Vorname.Substring(0,1) + $user.Nachname).Length))
    # - $email: E-Mail-Adresse im Format vorname.nachname@domain
    $email = "$($user.Vorname.ToLower()).$($user.Nachname.ToLower())@$domain"
    # - $ouName: Name der OU, basierend auf dem Kursfeld
    $ouName = $user.Kurs
    # - $ouPath: Pfad der OU im AD (unter $baseOU)
    $ouPath = "OU=$ouName,$baseOU"
    # - $userPrincipalName: UPN des Benutzers (username@domain)
    $userPrincipalName = "$username@$domain"

    # OU prüfen / erstellen
    # - Prüfe, ob die OU existiert, mit LDAP-Filter für distinguishedName
    # - -ErrorAction SilentlyContinue unterdrückt Fehler, wenn OU nicht existiert
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouPath)" -ErrorAction SilentlyContinue)) {
        try {
            # Erstelle die OU, wenn sie nicht existiert
            # - New-ADOrganizationalUnit erstellt die OU
            # - -Name $ouName gibt den Namen der OU an
            # - -Path $baseOU definiert den übergeordneten Pfad
            New-ADOrganizationalUnit -Name $ouName -Path $baseOU
            Log " OU erstellt: $ouPath"
        } catch {
            # Protokolliere Fehler beim Erstellen der OU und überspringe Benutzer
            Log " FEHLER beim Erstellen der OU: $ouPath - $($_.Exception.Message)"
            continue
        }
    }

    # Benutzerprüfung
    # - Suche nach existierendem Benutzer anhand des UserPrincipalName
    # - -Properties DistinguishedName lädt den DN für OU-Vergleiche
    $existingUser = Get-ADUser -Filter { UserPrincipalName -eq $userPrincipalName } -Properties DistinguishedName -ErrorAction SilentlyContinue

    # Verarbeite Benutzer basierend auf Status_code
    switch ($user.Status_code) {
        1 {
            # Status_code 1: Benutzer aktivieren oder erstellen
            if (-not $existingUser) {
                # Benutzer existiert nicht -> erstellen
                try {
                    # Erstelle neuen AD-Benutzer mit New-ADUser
                    # - Setze Attribute wie Name, E-Mail, Abteilung, Adresse usw.
                    # - -Path $ouPath definiert die Ziel-OU
                    # - -AccountPassword setzt ein Standardpasswort
                    # - -Enabled $true aktiviert den Benutzer
                    New-ADUser -Name "$($user.Vorname) $($user.Nachname)" `
                        -GivenName $user.Vorname `
                        -Surname $user.Nachname `
                        -UserPrincipalName $userPrincipalName `
                        -SamAccountName $username `
                        -EmailAddress $email `
                        -Department $user.Abteilung `
                        -StreetAddress $user.Strasse `
                        -PostalCode $user.PLZ `
                        -City $user.Ort `
                        -OfficePhone $user.Telefon `
                        -Title $user.Kursbezeichnung `
                        -Path $ouPath `
                        -AccountPassword (ConvertTo-SecureString "Passw0rd!" -AsPlainText -Force) `
                        -Enabled $true
                    Log " Benutzer erstellt: $userPrincipalName"
                } catch {
                    # Protokolliere Fehler beim Erstellen
                    Log " FEHLER beim Erstellen: $userPrincipalName - $($_.Exception.Message)"
                }
            } else {
                # Benutzer existiert -> aktualisieren
                # Prüfe, ob Benutzer in korrekter OU ist
                $currentOU = ($existingUser.DistinguishedName -split ",", 2)[1]
                if ($currentOU -ne $ouPath) {
                    try {
                        # Verschiebe Benutzer in neue OU
                        Move-ADObject -Identity $existingUser.DistinguishedName -TargetPath $ouPath
                        Log " Benutzer in neue OU verschoben: $userPrincipalName -> $ouPath"
                    } catch {
                        # Protokolliere Fehler beim Verschieben
                        Log " FEHLER beim Verschieben des Benutzers: $userPrincipalName - $($_.Exception.Message)"
                    }
                }
                try {
                    # Aktualisiere Benutzerattribute
                    Set-ADUser $existingUser `
                        -Department $user.Abteilung `
                        -StreetAddress $user.Strasse `
                        -PostalCode $user.PLZ `
                        -City $user.Ort `
                        -EmailAddress $email `
                        -OfficePhone $user.Telefon `
                        -Title $user.Kursbezeichnung
                    # Aktiviere Benutzerkonto
                    Enable-ADAccount $existingUser
                    Log " Benutzer aktualisiert & aktiviert: $userPrincipalName"
                } catch {
                    # Protokolliere Fehler beim Aktualisieren
                    Log "❌ FEHLER beim Aktualisieren: $userPrincipalName - $($_.Exception.Message)"
                }
            }
        }
        2 {
            # Status_code 2: Benutzer deaktivieren oder inaktiv erstellen
            if (-not $existingUser) {
                # Benutzer existiert nicht -> inaktiv erstellen
                try {
                    # Erstelle neuen AD-Benutzer mit deaktiviertem Konto
                    New-ADUser -Name "$($user.Vorname) $($user.Nachname)" `
                        -GivenName $user.Vorname `
                        -Surname $user.Nachname `
                        -UserPrincipalName $userPrincipalName `
                        -SamAccountName $username `
                        -EmailAddress $email `
                        -Department $user.Abteilung `
                        -StreetAddress $user.Strasse `
                        -PostalCode $user.PLZ `
                        -City $user.Ort `
                        -OfficePhone $user.Telefon `
                        -Title $user.Kursbezeichnung `
                        -Path $ouPath `
                        -AccountPassword (ConvertTo-SecureString "Passw0rd!" -AsPlainText -Force) `
                        -Enabled $false
                    Log "➖ Inaktiver Benutzer erstellt: $userPrincipalName"
                } catch {
                    # Protokolliere Fehler beim Erstellen
                    Log " FEHLER beim Erstellen (inaktiv): $userPrincipalName - $($_.Exception.Message)"
                }
            } else {
                # Benutzer existiert -> deaktivieren
                # Prüfe, ob Benutzer in korrekter OU ist
                $currentOU = ($existingUser.DistinguishedName -split ",", 2)[1]
                if ($currentOU -ne $ouPath) {
                    try {
                        # Verschiebe Benutzer in neue OU
                        Move-ADObject -Identity $existingUser.DistinguishedName -TargetPath $ouPath
                        Log " Benutzer in neue OU verschoben: $userPrincipalName -> $ouPath"
                    } catch {
                        # Protokolliere Fehler beim Verschieben
                        Log " FEHLER beim Verschieben des Benutzers: $userPrincipalName - $($_.Exception.Message)"
                    }
                }
                try {
                    # Aktualisiere Benutzerattribute
                    Set-ADUser $existingUser `
                        -Department $user.Abteilung `
                        -StreetAddress $user.Strasse `
                        -PostalCode $user.PLZ `
                        -City $user.Ort `
                        -EmailAddress $email `
                        -OfficePhone $user.Telefon `
                        -Title $user.Kursbezeichnung
                    # Deaktiviere Benutzerkonto
                    Disable-ADAccount $existingUser
                    Log " Benutzer deaktiviert: $userPrincipalName"
                } catch {
                    # Protokolliere Fehler beim Deaktivieren
                    Log " FEHLER beim Deaktivieren: $userPrincipalName - $($_.Exception.Message)"
                }
            }
        }
        3 {
            # Status_code 3: Benutzer löschen
            if ($existingUser) {
                try {
                    # Lösche Benutzer aus AD
                    Remove-ADUser $existingUser -Confirm:$false
                    Log " Benutzer gelöscht: $userPrincipalName"
                } catch {
                    # Protokolliere Fehler beim Löschen
                    Log " FEHLER beim Löschen: $userPrincipalName - $($_.Exception.Message)"
                }
            } else {
                # Protokolliere, wenn Benutzer nicht existiert
                Log " Benutzer nicht gefunden (zum Löschen): $userPrincipalName"
            }
        }
        default {
            # Unbekannter Status_code
            Log " WARNUNG: Unbekannter Status_code ($($user.Status_code)) für $userPrincipalName"
        }
    }
}