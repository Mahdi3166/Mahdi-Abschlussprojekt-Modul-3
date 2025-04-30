param(
    [string]$csvFolderPath = "\\Admin-Server\logging"
)

# Einstellungen
$domain = "M-zukunftsmotor.local"
$baseOU = "OU=M-zukunftsmotor,DC=M-zukunftsmotor,DC=local"
$logFile = "$csvFolderPath\user-import-log.txt"

# Neueste CSV-Datei finden
$csvFile = Get-ChildItem -Path $csvFolderPath -Filter *.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $csvFile) {
    Write-Host "❌ Keine gültige CSV-Datei gefunden im Pfad: $csvFolderPath"
    exit
}

# CSV-Datei laden
$users = Import-Csv -Path $csvFile.FullName -Delimiter ','

# Logfunktion
function Log($text) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $text" | Out-File -FilePath $logFile -Append -Encoding utf8
    Write-Host $text
}

foreach ($user in $users) {
    # Pflichtfelder prüfen
    if (-not $user.Vorname -or -not $user.Nachname -or -not $user.Kurs -or -not $user.Status_code) {
        Log "⚠️ WARNUNG: Fehlende Pflichtfelder für $($user.Vorname) $($user.Nachname): Kurs oder Status_code"
        continue
    }

    # Benutzerdaten vorbereiten
    $username = ($user.Vorname.Substring(0,1) + $user.Nachname).ToLower()
    $email = "$($user.Vorname.ToLower()).$($user.Nachname.ToLower())@$domain"
    $ouName = $user.Kurs
    $ouPath = "OU=$ouName,$baseOU"
    $userPrincipalName = "$username@$domain"

    # OU prüfen / erstellen
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$ouPath)" -ErrorAction SilentlyContinue)) {
        try {
            New-ADOrganizationalUnit -Name $ouName -Path $baseOU
            Log "✅ OU erstellt: $ouPath"
        } catch {
            Log "❌ FEHLER beim Erstellen der OU: $ouPath - $($_.Exception.Message)"
            continue
        }
    }

    # Benutzerprüfung
    $existingUser = Get-ADUser -Filter { UserPrincipalName -eq $userPrincipalName } -Properties DistinguishedName -ErrorAction SilentlyContinue

    switch ($user.Status_code) {
        1 {
            if (-not $existingUser) {
                try {
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
                    Log "✅ Benutzer erstellt: $userPrincipalName"
                } catch {
                    Log "❌ FEHLER beim Erstellen: $userPrincipalName - $($_.Exception.Message)"
                }
            } else {
                # OU-Verschiebung prüfen und durchführen
                $currentOU = ($existingUser.DistinguishedName -split ",", 2)[1]
                if ($currentOU -ne $ouPath) {
                    try {
                        Move-ADObject -Identity $existingUser.DistinguishedName -TargetPath $ouPath
                        Log "📁 Benutzer in neue OU verschoben: $userPrincipalName -> $ouPath"
                    } catch {
                        Log "❌ FEHLER beim Verschieben des Benutzers: $userPrincipalName - $($_.Exception.Message)"
                    }
                }
                try {
                    Set-ADUser $existingUser `
                        -Department $user.Abteilung `
                        -StreetAddress $user.Strasse `
                        -PostalCode $user.PLZ `
                        -City $user.Ort `
                        -EmailAddress $email `
                        -OfficePhone $user.Telefon `
                        -Title $user.Kursbezeichnung
                    Enable-ADAccount $existingUser
                    Log "🔄 Benutzer aktualisiert & aktiviert: $userPrincipalName"
                } catch {
                    Log "❌ FEHLER beim Aktualisieren: $userPrincipalName - $($_.Exception.Message)"
                }
            }
        }
        2 {
            if (-not $existingUser) {
                try {
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
                    Log "❌ FEHLER beim Erstellen (inaktiv): $userPrincipalName - $($_.Exception.Message)"
                }
            } else {
                # OU-Verschiebung prüfen und durchführen
                $currentOU = ($existingUser.DistinguishedName -split ",", 2)[1]
                if ($currentOU -ne $ouPath) {
                    try {
                        Move-ADObject -Identity $existingUser.DistinguishedName -TargetPath $ouPath
                        Log "📁 Benutzer in neue OU verschoben: $userPrincipalName -> $ouPath"
                    } catch {
                        Log "❌ FEHLER beim Verschieben des Benutzers: $userPrincipalName - $($_.Exception.Message)"
                    }
                }
                try {
                    Set-ADUser $existingUser `
                        -Department $user.Abteilung `
                        -StreetAddress $user.Strasse `
                        -PostalCode $user.PLZ `
                        -City $user.Ort `
                        -EmailAddress $email `
                        -OfficePhone $user.Telefon `
                        -Title $user.Kursbezeichnung
                    Disable-ADAccount $existingUser
                    Log "⛔ Benutzer deaktiviert: $userPrincipalName"
                } catch {
                    Log "❌ FEHLER beim Deaktivieren: $userPrincipalName - $($_.Exception.Message)"
                }
            }
        }
        3 {
            if ($existingUser) {
                try {
                    Remove-ADUser $existingUser -Confirm:$false
                    Log "🗑️ Benutzer gelöscht: $userPrincipalName"
                } catch {
                    Log "❌ FEHLER beim Löschen: $userPrincipalName - $($_.Exception.Message)"
                }
            } else {
                Log "ℹ️ Benutzer nicht gefunden (zum Löschen): $userPrincipalName"
            }
        }
        default {
            Log "⚠️ WARNUNG: Unbekannter Status_code ($($user.Status_code)) für $userPrincipalName"
        }
    }
}
