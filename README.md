# eCall Log Analyzer

PowerShell-Script zur monatlichen Verrechnung der [eCall SMS/FAX-Logfiles](https://help.ecall-messaging.com/de/article/logfiles-1fa1qpk/) auf Kostenstellen und Innenaufträge.

eCall stellt die Logfiles des Vormonats als ZIP-Anhang per E-Mail zu. Das Script holt diese Mails per **Microsoft Graph** aus einem dedizierten Postfach ab, liest die In-/Out-Logfiles beider eCall-Accounts ein, verrechnet die Punkte anhand einer Stammdaten-Tabelle und erzeugt daraus ein Verrechnungs-CSV. Zum Abschluss wird ein HTML-Report per Mail versendet.

> **Wichtig — die Logfiles kommen nicht automatisch.** Der monatliche Versand der Logfiles muss beim [eCall-Support](https://www.ecall.ch/support/) aktiv beantragt werden. Erst wenn eCall die Einlieferung für die betroffenen Accounts eingerichtet hat, treffen die ZIP-Anhänge im konfigurierten Postfach ein. Ohne diese Freischaltung findet das Script keine Logfiles vor und meldet nur `WARN`-Zeilen.

> **Warum Graph statt EWS?** Exchange Web Services (EWS) wird für Exchange Online am **1. Oktober 2026** abgeschaltet. Diese Version verwendet durchgängig Microsoft Graph mit App-Only-Authentifizierung (Zertifikat) — es werden keine Passwörter mehr im Script oder auf der Platte benötigt.

## Funktionsweise

1. **Abholen:** Alle Mails im Posteingang des Log-Postfachs werden verarbeitet; ZIP-Anhänge werden in einen Import-Ordner heruntergeladen und entpackt. Verarbeitete Mails wandern in „Gelöschte Elemente".
2. **Einlesen:** Die entpackten Logfiles werden anhand ihres Dateinamens den vier Datensätzen zugeordnet (Account 1/2, jeweils In/Out). Mehrere Dateien pro Datensatz werden akkumuliert.
3. **Verrechnen:** Jede Punktesumme wird über den `Buchungstext` der Stammdaten einer Kostenstelle / einem Innenauftrag zugeordnet:
   - **Account 1:** vollständige Absender- bzw. Antwortadresse
   - **Account 2:** Maildomain des Absenders bzw. der Antwortadresse
   - **Alle Accounts:** `ExterneID` (API-Meldungen); IDs mit Präfix `eCallURL` laufen gesammelt auf die Position `eCallURL`
   - Nicht zuordenbare Punkte landen auf dem konfigurierbaren **Default-Eintrag**.
4. **Kontrolle:** Die Gesamtpunkte laut Logfiles werden gegen die verrechneten Punkte geprüft; eine Restdifferenz wird auf den Default-Eintrag gebucht und geloggt.
5. **Output:** Verrechnungs-CSV im Reporting-Ordner + HTML-Report-Mail mit Logfile und CSV als Anhang.

## Voraussetzungen

- Windows Server mit **PowerShell 5.1 oder 7.x**
- **Microsoft Graph PowerShell SDK** (nur die drei benötigten Module):

  ```powershell
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Mail, Microsoft.Graph.Users.Actions -Scope AllUsers
  ```

- Exchange-Online-Postfach, in dem die eCall-Logfiles eintreffen (dediziertes Shared Mailbox empfohlen)
- Schreibzugriff des ausführenden Kontos auf die konfigurierten Netzwerkfreigaben

## Einrichtung

### 1. Entra-App-Registrierung

1. [Entra Admin Center](https://entra.microsoft.com) → **App registrations** → **New registration**
   - Name: z. B. `eCall-Log-Analyzer`
   - Supported account types: *Accounts in this organizational directory only*
2. Unter **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**:
   - `Mail.ReadWrite` (Logfiles lesen, Mails in den Papierkorb verschieben)
   - `Mail.Send` (Report-Mail versenden)
3. **Grant admin consent** ausführen.
4. **Tenant-ID** und **Application (client) ID** von der Übersichtsseite notieren → im Script bei `$TenantId` und `$GraphClientId` eintragen.

### 2. Zertifikat erstellen und hinterlegen

Auf dem Server, der das Script ausführt (im Kontext des Task-Kontos bzw. in `LocalMachine`, wenn ein Dienstkonto verwendet wird):

```powershell
$cert = New-SelfSignedCertificate -Subject 'CN=eCall-Log-Analyzer' `
    -CertStoreLocation 'Cert:\LocalMachine\My' `
    -KeyExportPolicy NonExportable -KeySpec Signature `
    -KeyLength 2048 -NotAfter (Get-Date).AddYears(2)

# Public Key exportieren (wird in Entra hochgeladen)
Export-Certificate -Cert $cert -FilePath .\eCall-Log-Analyzer.cer

$cert.Thumbprint   # -> im Script bei $GraphCertThumbprint eintragen
```

Dann in der App-Registrierung unter **Certificates & secrets** → **Certificates** → **Upload certificate** die `.cer`-Datei hochladen.

> Das ausführende Konto (Scheduled-Task-Benutzer) braucht Leserecht auf den privaten Schlüssel: Zertifikat in `certlm.msc` → Rechtsklick → *All Tasks* → *Manage Private Keys*.
>
> Vor Ablauf des Zertifikats (hier: 2 Jahre) rechtzeitig ein neues erstellen, hochladen und den Thumbprint im Script aktualisieren.

### 3. Zugriff auf die benötigten Postfächer einschränken

`Mail.ReadWrite`/`Mail.Send` als Application Permission gelten sonst **tenant-weit**. Mit einer Application Access Policy wird die App auf das Log- und das Absenderpostfach begrenzt (Exchange Online PowerShell, einmalig):

```powershell
# Mail-aktivierte Sicherheitsgruppe mit den beiden Postfächern anlegen,
# z. B. "eCall-Log-Analyzer-Mailboxes", dann:
New-ApplicationAccessPolicy -AppId '<Application-ID>' `
    -PolicyScopeGroupId 'eCall-Log-Analyzer-Mailboxes@example.com' `
    -AccessRight RestrictAccess `
    -Description 'eCall Log Analyzer: nur Log- und Reporting-Postfach'

# Wirksamkeit pruefen:
Test-ApplicationAccessPolicy -AppId '<Application-ID>' -Identity 'ecall-logs@example.com'      # Granted
Test-ApplicationAccessPolicy -AppId '<Application-ID>' -Identity 'irgendwer@example.com'       # Denied
```

### 4. Script konfigurieren

Alle Einstellungen stehen gebündelt im Block `#region Konfiguration` von [`billing.ps1`](billing.ps1):

| Variable | Bedeutung |
|---|---|
| `$TenantId`, `$GraphClientId`, `$GraphCertThumbprint` | Werte aus Schritt 1 und 2 |
| `$LogMailbox` | Postfach, in dem die eCall-Logfiles eintreffen |
| `$MailFrom`, `$MailTo` | Absender und Empfänger der Report-Mail |
| `$SourceFolder` | Ordner mit `Stammdaten.csv` und `Leistungsarten.csv` |
| `$ShareFolder`, `$ImportRoot`, `$ReportingFolder`, `$LockFolder` | Ablage für Logs, entpackte Logfiles, Verrechnungs-CSV und Monats-Lockfile |
| `$DefaultBuchungstext` | Stammdaten-Eintrag, auf den Unzuordenbares gebucht wird (muss existieren, sonst Abbruch) |
| `$Zuweisungsgruppe`, `$ServiceName` | Texte für Report-Mail und Dateinamen |
| `$LogFileMap` | Dateinamensmuster → Datensatz; weitere eCall-Accounts = eine Zeile mehr |

### 5. Eingabedateien

Beide Dateien liegen im `$SourceFolder` auf dem Server (nicht im Repo). Als Vorlage mit dem exakten Format liegen sie unter [`examples/`](examples/) — die realen Dateien mit echten Kostenstellen und Mailadressen gehören **nicht** ins Repository (per `.gitignore` geschützt).

**`Stammdaten.csv`** (Semikolon-getrennt, mit Kopfzeile) — die Zuordnungstabelle:

```csv
Innenauftrag;Kostenstelle;Kundenauftrag;Auftragsposition;Buchungstext;Menge
IA100001;KST5000;KA200001;10;alerting@example.com;0
IA100002;KST5100;KA200002;10;subdomain.example.com;0
IA100003;KST5200;KA200003;10;MeineExterneID;0
IA100009;KST5900;KA200009;10;eCallURL;0
IA100010;KST5999;KA200010;10;default.example.com;0
```

Der `Buchungstext` ist der Zuordnungsschlüssel: je nach Account eine vollständige Mailadresse, eine Domain oder eine ExterneID. Die Einträge `eCallURL` und der Default-Eintrag müssen vorhanden sein.

**`Leistungsarten.csv`** (Semikolon-getrennt, mit Kopfzeile) — wird 1:1 in den Kopf des Verrechnungsfiles übernommen:

```csv
Leistungsart;Bezeichnung
L001;SMS Versand
L002;SMS Empfang
L003;FAX Versand
L004;FAX Empfang
```

### 6. Geplante Ausführung

Das Script ist für einen monatlichen Lauf (Anfang Monat, nach Zustellung der eCall-Logfiles) als geplante Aufgabe gedacht:

```powershell
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument '-NoProfile -ExecutionPolicy Bypass -File "D:\Temp\SCRIPTS\Log_Analyzer\billing.ps1"'
$trigger = New-ScheduledTaskTrigger -At 06:00 -Daily
Register-ScheduledTask -TaskName 'eCall-Verrechnung' -Action $action -Trigger $trigger `
    -User 'DOMAIN\svc-ecall' -Password (Read-Host 'Passwort Dienstkonto')
```

Ein eingebauter **Monats-Lock** (`Check_<JJJJMM>.txt` im `$LockFolder`) verhindert Doppelverrechnung: Läuft das Script ein zweites Mal im selben Monat, bricht es mit Fehler ab (inkl. Eintrag im Windows-Eventlog). Der Trigger darf also ruhig täglich feuern — verrechnet wird trotzdem nur einmal pro Monat, sobald Logfiles im Postfach liegen. Bewusster zweiter Lauf:

```powershell
.\billing.ps1 -Force
```

## Output

| Artefakt | Ablage |
|---|---|
| Verrechnungs-CSV `<ServiceName>_SMS_FAX_<JJJJMM>.csv` | `$ReportingFolder` |
| Tageslogfile `logfile_<JJJJMMTT>.txt` | `$ShareFolder\logs` |
| Entpackte Roh-Logfiles (ein Unterordner pro Mail, inkl. `metadata.json`) | `$ImportRoot\<Zeitstempel>` |
| HTML-Report-Mail (mit CSV und Logfile als Anhang) | an `$MailTo` |

## Troubleshooting

| Symptom | Ursache / Lösung |
|---|---|
| `Access is denied` / `ErrorAccessDenied` bei Graph-Aufrufen | Admin Consent fehlt, oder die Application Access Policy schließt das Postfach aus (`Test-ApplicationAccessPolicy` prüfen) |
| `Certificate with thumbprint ... not found` | Zertifikat liegt nicht im Store des ausführenden Kontos, oder das Task-Konto hat kein Leserecht auf den privaten Schlüssel |
| Abbruch „bereits ausgeführt" | Gewollt (Monats-Lock). Bewusster zweiter Lauf: `-Force` |
| WARN `Kein Stammdaten-Eintrag fuer '...'` im Log | Neuer Absender/Domain/ExterneID ohne Stammdaten-Zeile — Punkte liefen auf den Default-Eintrag. Zeile in `Stammdaten.csv` ergänzen |
| WARN `Kein Logfile fuer Account..._...` | eCall-Mail noch nicht eingetroffen oder Dateinamensmuster in `$LogFileMap` passt nicht |
| Hohe Menge auf dem Default-Eintrag | Erwartet für alles Unzuordenbare — im Tageslogfile stehen die betroffenen Buchungstexte |

## Hinweise zur Verrechnungslogik

- Zeilen, die **sowohl** eine Absender-/Antwortadresse **als auch** eine `ExterneID` enthalten, werden in beiden Durchgängen gebucht. Die Schlusskontrolle gleicht die Gesamtsumme über den Default-Eintrag wieder aus, die einzelnen Kostenstellen wären in diesem Fall aber überzeichnet. Bei eCall ist pro Meldungstyp üblicherweise nur eines der Felder gefüllt.
- Punktedifferenzen zwischen Logfile-Summe und verrechneter Summe werden immer auf den Default-Eintrag gebucht und im Log ausgewiesen — das Verrechnungsfile geht in der Gesamtsumme also immer auf.
