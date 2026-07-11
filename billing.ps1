#requires -Version 5.1

<#
.SYNOPSIS
    Monatliche Verrechnung der eCall SMS/FAX-Logfiles (Microsoft Graph).

.DESCRIPTION
    Verarbeitet die von eCall per E-Mail zugestellten monatlichen Logfiles
    (Format: https://help.ecall-messaging.com/de/article/logfiles-1fa1qpk/):

      1. ZIP-Anhaenge per Microsoft Graph aus dem Postfach herunterladen und entpacken
      2. In-/Out-Logfiles der beiden eCall-Accounts einlesen
      3. Punkte anhand der Stammdaten auf Kostenstellen/Innenauftraege verrechnen
         - Account 1: Zuordnung ueber die vollstaendige Absender-/Antwortadresse
         - Account 2: Zuordnung ueber die Maildomain
         - Zusaetzlich: Zuordnung ueber ExterneID (API-Meldungen) sowie
           Sammelposition "eCallURL"
      4. Nicht zuordenbare Restpunkte auf den Default-Eintrag buchen
      5. Verrechnungs-CSV erzeugen und Report-Mail via Graph versenden

    Authentifizierung: App-Only (Entra-App-Registrierung + Zertifikat).
    Einrichtung siehe README.md.

.PARAMETER Force
    Fuehrt das Skript aus, auch wenn es in diesem Monat bereits gelaufen ist.

.NOTES
    Benoetigte Module : Microsoft.Graph.Authentication, Microsoft.Graph.Mail,
                        Microsoft.Graph.Users.Actions
    Benoetigte Rechte : Mail.ReadWrite, Mail.Send (Application, per
                        ApplicationAccessPolicy auf die Postfaecher eingeschraenkt)
#>
[CmdletBinding()]
param(
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Konfiguration ----------------------------------------------------------

# Entra ID / Microsoft Graph (App-Only-Authentifizierung, siehe README.md)
$TenantId            = '00000000-0000-0000-0000-000000000000'
$GraphClientId       = '00000000-0000-0000-0000-000000000000'
$GraphCertThumbprint = '0000000000000000000000000000000000000000'

$LogMailbox = 'ecall-logs@example.com'    # Postfach, in dem die eCall-Logfiles eintreffen
$MailFrom   = 'reporting@example.com'     # Absenderpostfach der Report-Mail
$MailTo     = 'recipient@example.com'     # Empfaenger der Report-Mail

$SourceFolder    = 'D:\Temp\SCRIPTS\Log_Analyzer'          # Stammdaten.csv, Leistungsarten.csv
$ShareFolder     = '\\FILESERVER01\Source'
$ImportRoot      = '\\FILESERVER01\Source\Import'
$ReportingFolder = '\\FILESERVER01\Source\Verrechnungsfiles'
$LockFolder      = '\\FILESERVER01\Source\Check'           # Merker "diesen Monat bereits gelaufen"

$DefaultBuchungstext = 'default.example.com'               # Auffangposition fuer Unzuordenbares
$Zuweisungsgruppe    = 'IT-SUPPORT-GROUP'
$ServiceName         = 'SERVICE_XYZ'

# Erwartete Logfiles und ihre Zuordnung zu den vier Datensaetzen
$LogFileMap = @(
    @{ Pattern = 'Account1_*_Out-LogFile.txt'; Key = 'Account1_Out'; Richtung = 'Out' }
    @{ Pattern = 'Account1_*_In-LogFile.txt';  Key = 'Account1_In';  Richtung = 'In'  }
    @{ Pattern = 'Account2_*_Out-LogFile.txt'; Key = 'Account2_Out'; Richtung = 'Out' }
    @{ Pattern = 'Account2_*_In-LogFile.txt';  Key = 'Account2_In';  Richtung = 'In'  }
)

#endregion

#region Abgeleitete Werte ------------------------------------------------------

$Vormonat  = (Get-Date).AddMonths(-1)
$Monat     = $Vormonat.ToString('MM')
$Jahr      = $Vormonat.Year
$MonatLang = $Vormonat.ToString('Y')                        # z.B. "Juni 2026"
$Heute     = Get-Date -Format 'dd.MM.yyyy'

$LogFile    = Join-Path $ShareFolder ('logs\logfile_{0}.txt' -f (Get-Date -Format 'yyyyMMdd'))
$LockFile   = Join-Path $LockFolder ('Check_{0}{1}.txt' -f $Jahr, $Monat)
$ReportFile = Join-Path $ReportingFolder ('{0}_SMS_FAX_{1}{2}.csv' -f $ServiceName, $Jahr, $Monat)

# Spaltenkoepfe der eCall-Logfiles gemaess F24/eCall-Spezifikation "Datenfelder
# bei der Logfile-Zustellung" v1.4. Die Dateien werden Semikolon-getrennt und
# OHNE Kopfzeile geliefert; die Zuordnung erfolgt daher rein ueber die Position.
#
# ACHTUNG: Die Meldungsfelder (OUT Feld 3 "Meldung"; IN Felder 3+4 "Gesendete"/
# "Empfangene Meldung") koennen von eCall auf Wunsch unterdrueckt werden. Werden
# dabei ganze Spalten entfernt, verschieben sich alle folgenden Felder und die
# Verrechnung liest falsche Spalten. Diese Header setzen voraus, dass ALLE Spalten
# geliefert werden (leere Werte sind ok). Details siehe README.md.
$OutHeader = 'Referenz', 'Startdatum', 'Meldung', 'Resultatcode', 'Absender',
             'Empfaengernummer', 'ExterneID', 'Punkte', 'Empfaengername'
$InHeader  = 'Referenz', 'Startdatum', 'GesendeteMeldung', 'EmpfangeneMeldung', 'Resultatcode',
             'SendeAdresse', 'eCallAdresse', 'AntwortAdresse', 'AntwortInfo', 'ExterneID', 'Punkte'

#endregion

#region Funktionen -------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'FATAL', 'DEBUG')][string]$Level = 'INFO'
    )
    $line = '{0} {1,-5} {2}' -f (Get-Date -Format 'yyyy/MM/dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $line
    $farbe = if ($Level -in 'ERROR', 'FATAL') { 'Red' } elseif ($Level -eq 'WARN') { 'Yellow' } else { 'Green' }
    Write-Host $line -ForegroundColor $farbe
}

function Get-MailDomain {
    # Liefert den Domainteil einer Mailadresse (leerzeichenbereinigt), sonst $null.
    param([string]$Adresse)
    (($Adresse -replace '\s', '') -split '@')[1]
}

function Get-PunkteSumme {
    param($Zeilen)
    $summe = ($Zeilen | Measure-Object -Property Punkte -Sum).Sum
    if ($null -eq $summe) { [decimal]0 } else { [decimal]$summe }
}

function Add-Verrechnungspunkte {
    # Bucht Punkte auf den Stammdaten-Eintrag mit passendem Buchungstext;
    # ohne Treffer auf den Default-Eintrag.
    param(
        [Parameter(Mandatory)][string]$Buchungstext,
        [Parameter(Mandatory)][decimal]$Punkte
    )
    if ($Punkte -eq 0) { return }

    $eintrag = $script:StammdatenIndex[$Buchungstext]
    if (-not $eintrag) {
        Write-Log -Level WARN -Message "Kein Stammdaten-Eintrag fuer '$Buchungstext' - $Punkte Punkte gehen auf '$DefaultBuchungstext'"
        $eintrag = $script:DefaultEintrag
    }
    $eintrag.Menge += $Punkte
    Write-Log -Message ('{0}: {1} Punkte (KST {2}, IA {3}, KA {4}, Pos {5})' -f
        $Buchungstext, $Punkte, $eintrag.Kostenstelle, $eintrag.Innenauftrag, $eintrag.Kundenauftrag, $eintrag.Auftragsposition)
}

function Import-ECallLogFile {
    param(
        [Parameter(Mandatory)][string]$Pfad,
        [Parameter(Mandatory)][ValidateSet('In', 'Out')][string]$Richtung
    )
    $header = if ($Richtung -eq 'Out') { $script:OutHeader } else { $script:InHeader }
    Import-Csv -Path $Pfad -Delimiter ';' -Header $header | ForEach-Object {
        $_.Punkte = if ($_.Punkte) { [decimal]$_.Punkte } else { [decimal]0 }
        $_
    }
}

function Get-GraphLogAttachments {
    # Laedt alle ZIP-Anhaenge aus dem Posteingang des Log-Postfachs, entpackt sie
    # pro Mail in einen eigenen Unterordner und verschiebt verarbeitete Mails in
    # "Geloeschte Elemente".
    param(
        [Parameter(Mandatory)][string]$Zielordner
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $mailNr = 0
    $maxDurchlaeufe = 50
    do {
        $maxDurchlaeufe--
        $messages = @(Get-MgUserMessage -UserId $LogMailbox -Top 100 `
            -Property id, subject, receivedDateTime, from, hasAttachments)
        Write-Log -Message "$($messages.Count) Mails im Posteingang gefunden"

        foreach ($msg in $messages) {
            $mailNr++
            $mailOrdner = Join-Path $Zielordner $mailNr
            New-Item -Path $mailOrdner -ItemType Directory | Out-Null

            try {
                Write-Log -Message "Mail $mailNr - $($msg.Subject)"

                # Metadaten der Mail zur Nachvollziehbarkeit ablegen
                [ordered]@{
                    Subject          = $msg.Subject
                    ReceivedDateTime = $msg.ReceivedDateTime
                    From             = $msg.From.EmailAddress.Address
                } | ConvertTo-Json | Set-Content -Path (Join-Path $mailOrdner 'metadata.json')

                $anhaenge = @(Get-MgUserMessageAttachment -UserId $LogMailbox -MessageId $msg.Id |
                    Where-Object {
                        $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.fileAttachment' -and
                        $_.Name -like '*.zip'
                    })

                foreach ($att in $anhaenge) {
                    $zipDatei = Join-Path $mailOrdner ($att.Name -replace '/', '')

                    # Rohinhalt streamen statt Base64 aus der Auflistung - funktioniert
                    # auch bei grossen Anhaengen zuverlaessig
                    $uri = 'https://graph.microsoft.com/v1.0/users/{0}/messages/{1}/attachments/{2}/$value' -f
                        $LogMailbox, $msg.Id, $att.Id
                    Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $zipDatei
                    Write-Log -Message ('Anhang gespeichert: {0} ({1} KB)' -f $zipDatei, [Math]::Round($att.Size / 1KB))

                    try {
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipDatei, $mailOrdner)
                    }
                    catch {
                        Write-Log -Level ERROR -Message "Entpacken fehlgeschlagen: $zipDatei - $_"
                    }
                }

                # Verarbeitete Mail in "Geloeschte Elemente" verschieben
                Move-MgUserMessage -UserId $LogMailbox -MessageId $msg.Id `
                    -DestinationId 'deleteditems' | Out-Null
            }
            catch {
                Write-Log -Level ERROR -Message "Mail konnte nicht verarbeitet werden: $_"
            }
        }
    } while ($messages.Count -eq 100 -and $maxDurchlaeufe -gt 0)
}

function Send-ReportMail {
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$HtmlBody,
        [string[]]$Attachments = @()
    )

    $mailAttachments = foreach ($pfad in $Attachments) {
        @{
            '@odata.type' = '#microsoft.graph.fileAttachment'
            name          = Split-Path $pfad -Leaf
            contentBytes  = [Convert]::ToBase64String([IO.File]::ReadAllBytes($pfad))
        }
    }

    $body = @{
        message         = @{
            subject      = $Subject
            body         = @{ contentType = 'HTML'; content = $HtmlBody }
            toRecipients = @(@{ emailAddress = @{ address = $MailTo } })
            attachments  = @($mailAttachments)
        }
        saveToSentItems = $true
    }

    Send-MgUserMail -UserId $MailFrom -BodyParameter $body
}

#endregion

#region Hauptablauf ------------------------------------------------------------

# Logdatei sicherstellen
$logDir = Split-Path $LogFile -Parent
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory | Out-Null }
if (-not (Test-Path $LogFile)) { New-Item -Path $LogFile -ItemType File | Out-Null }

$graphVerbunden = $false
try {
    Write-Log -Message "=== Start Verrechnung $ServiceName fuer $MonatLang ==="

    # Doppelausfuehrung im selben Monat verhindern
    if (Test-Path $LockFile) {
        if (-not $Force) {
            $meldung = "Das Verrechnungsskript wurde diesen Monat bereits ausgefuehrt ($LockFile). Erneuter Lauf nur mit -Force."
            Write-Log -Level ERROR -Message $meldung
            Write-EventLog -LogName Application -Source 'Application Error' -EventId 1 -EntryType Error -Message $meldung
            throw $meldung
        }
        Write-Log -Level WARN -Message 'Monats-Check per -Force uebersteuert'
    }
    New-Item -Path $LockFile -ItemType File -Force | Out-Null

    # Stammdaten einlesen, Menge als Dezimalwert initialisieren
    $Stammdaten = @(Import-Csv -Path (Join-Path $SourceFolder 'Stammdaten.csv') -Delimiter ';')
    foreach ($eintrag in $Stammdaten) {
        if ($eintrag.PSObject.Properties['Menge']) {
            $eintrag.Menge = if ($eintrag.Menge) { [decimal]$eintrag.Menge } else { [decimal]0 }
        }
        else {
            $eintrag | Add-Member -MemberType NoteProperty -Name Menge -Value ([decimal]0)
        }
    }
    Write-Log -Message "Stammdaten.csv eingelesen ($($Stammdaten.Count) Eintraege)"

    # Index Buchungstext -> Eintrag (bei Duplikaten gewinnt der erste)
    $StammdatenIndex = @{}
    foreach ($eintrag in $Stammdaten) {
        if (-not $StammdatenIndex.ContainsKey($eintrag.Buchungstext)) {
            $StammdatenIndex[$eintrag.Buchungstext] = $eintrag
        }
    }

    $DefaultEintrag = $StammdatenIndex[$DefaultBuchungstext]
    if (-not $DefaultEintrag) { throw "Default-Eintrag '$DefaultBuchungstext' fehlt in den Stammdaten" }

    $Leistungsarten = @(Import-Csv -Path (Join-Path $SourceFolder 'Leistungsarten.csv') -Delimiter ';')
    Write-Log -Message 'Leistungsarten.csv eingelesen'

    # Graph-Verbindung (App-Only mit Zertifikat)
    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Mail, Microsoft.Graph.Users.Actions
    Connect-MgGraph -TenantId $TenantId -ClientId $GraphClientId `
        -CertificateThumbprint $GraphCertThumbprint -NoWelcome
    $graphVerbunden = $true
    Write-Log -Message "Mit Microsoft Graph verbunden (App $GraphClientId)"

    # Logfiles aus dem Postfach abholen
    $importFolder = Join-Path $ImportRoot (Get-Date -Format 'yyyyMMdd HHmmss')
    New-Item -Path $importFolder -ItemType Directory | Out-Null
    Get-GraphLogAttachments -Zielordner $importFolder

    # Entpackte Logfiles den vier Datensaetzen zuordnen und einlesen
    $Logs = @{}
    foreach ($map in $LogFileMap) { $Logs[$map.Key] = @() }

    foreach ($datei in Get-ChildItem -Path $importFolder -Recurse -Filter '*.txt') {
        $map = $LogFileMap | Where-Object { $datei.Name -like $_.Pattern } | Select-Object -First 1
        if (-not $map) {
            Write-Log -Level WARN -Message "Unbekanntes Logfile wird ignoriert: $($datei.Name)"
            continue
        }
        $Logs[$map.Key] += @(Import-ECallLogFile -Pfad $datei.FullName -Richtung $map.Richtung)
        Write-Log -Message "$($datei.Name) eingelesen ($($map.Key), gesamt $($Logs[$map.Key].Count) Zeilen)"
    }

    foreach ($map in $LogFileMap) {
        if ($Logs[$map.Key].Count -eq 0) {
            Write-Log -Level WARN -Message "Kein Logfile fuer $($map.Key) gefunden"
        }
    }

    # --- Verrechnung nach Absender / Antwortadresse -----------------------------

    Write-Log -Message 'Verrechnung Account 1 (Zuordnung ueber Mailadresse)'
    $Logs.Account1_Out | Where-Object { $_.Absender } | Group-Object -Property Absender | ForEach-Object {
        Add-Verrechnungspunkte -Buchungstext $_.Name -Punkte (Get-PunkteSumme $_.Group)
    }
    $Logs.Account1_In | Where-Object { $_.AntwortAdresse } | Group-Object -Property AntwortAdresse | ForEach-Object {
        Add-Verrechnungspunkte -Buchungstext $_.Name -Punkte (Get-PunkteSumme $_.Group)
    }

    Write-Log -Message 'Verrechnung Account 2 (Zuordnung ueber Maildomain)'
    $Logs.Account2_Out | Where-Object { $_.Absender } | Group-Object -Property { Get-MailDomain $_.Absender } | ForEach-Object {
        $key = if ($_.Name) { $_.Name } else { $DefaultBuchungstext }
        Add-Verrechnungspunkte -Buchungstext $key -Punkte (Get-PunkteSumme $_.Group)
    }
    # Beim In-File zaehlen auch Zeilen ohne Antwortadresse - sie gehen auf den Default-Eintrag
    $Logs.Account2_In | Group-Object -Property { Get-MailDomain $_.AntwortAdresse } | ForEach-Object {
        $key = if ($_.Name) { $_.Name } else { $DefaultBuchungstext }
        Add-Verrechnungspunkte -Buchungstext $key -Punkte (Get-PunkteSumme $_.Group)
    }

    # --- Verrechnung nach ExterneID (API-Meldungen) -----------------------------

    foreach ($key in $LogFileMap.Key) {
        Write-Log -Message "Verrechnung ExterneID ($key)"

        $Logs[$key] |
            Where-Object { $_.ExterneID -and $_.ExterneID -notlike '*ecallURL*' } |
            Group-Object -Property ExterneID | ForEach-Object {
                Add-Verrechnungspunkte -Buchungstext $_.Name -Punkte (Get-PunkteSumme $_.Group)
            }

        # eCall-URL-Meldungen laufen gesammelt auf die Position "eCallURL"
        $urlPunkte = Get-PunkteSumme ($Logs[$key] | Where-Object { $_.ExterneID -like 'eCallURL*' })
        Add-Verrechnungspunkte -Buchungstext 'eCallURL' -Punkte $urlPunkte
    }

    # --- Kontrollsummen und Restpunkte ------------------------------------------

    $PunkteJe = @{}
    foreach ($key in $LogFileMap.Key) { $PunkteJe[$key] = Get-PunkteSumme $Logs[$key] }

    $PunkteAccount1 = $PunkteJe['Account1_Out'] + $PunkteJe['Account1_In']
    $PunkteAccount2 = $PunkteJe['Account2_Out'] + $PunkteJe['Account2_In']
    $Gesamtpunkte   = $PunkteAccount1 + $PunkteAccount2

    $VerrechnetePunkte = [decimal](($Stammdaten | Measure-Object -Property Menge -Sum).Sum)
    $Differenz = $Gesamtpunkte - $VerrechnetePunkte
    Write-Log -Message "Gesamtpunkte laut Logfiles: $Gesamtpunkte / verrechnet: $VerrechnetePunkte / Differenz: $Differenz"

    # Nicht zugeordnete Restpunkte auf den Default-Eintrag buchen
    if ($Differenz -ne 0) {
        $DefaultEintrag.Menge += $Differenz
        Write-Log -Message "Restdifferenz von $Differenz Punkten auf '$DefaultBuchungstext' gebucht"
    }

    $Kontrolle = $Gesamtpunkte - [decimal](($Stammdaten | Measure-Object -Property Menge -Sum).Sum)
    if ($Kontrolle -ne 0) {
        Write-Log -Level ERROR -Message "Kontrollsumme geht nicht auf (Restdifferenz $Kontrolle)"
    }

    # --- Verrechnungsfile erzeugen ----------------------------------------------

    $kopfzeilen = @(
        'SKST;XXXX Communication Services'
        'Modul;XXXXXX SMS/FAX'
        "Erstellung;$Heute"
        'Leistungsart;Leistungsgroessen'
    )
    $kopfzeilen += $Leistungsarten | ForEach-Object { '{0};{1}' -f $_.Leistungsart, $_.Bezeichnung }

    Set-Content -Path $ReportFile -Value $kopfzeilen -Encoding UTF8
    $Stammdaten | ConvertTo-Csv -Delimiter ';' -NoTypeInformation | Add-Content -Path $ReportFile -Encoding UTF8
    Write-Log -Message "Verrechnungsfile erstellt: $ReportFile"

    # --- Report-Mail versenden ---------------------------------------------------

    $mailBody = @"
<h1 style="color:#5e9ca0;">$ServiceName Periode <span style="color:#2b2301;">$MonatLang</span></h1>
<h2 style="color:#2e6c80;">Quick-Infos:</h2>
<table style="width:606px;" border="2">
<tbody>
<tr><td style="width:265px;"><strong>Ausl&ouml;sedatum</strong></td><td style="width:337px;">$Heute</td></tr>
<tr><td><strong>Generierender Server</strong></td><td>$env:COMPUTERNAME</td></tr>
<tr><td><strong>Punktezahl Gesamt</strong></td><td>$Gesamtpunkte</td></tr>
<tr><td><strong>Punktezahl Account ACCOUNT1</strong></td><td>$PunkteAccount1</td></tr>
<tr><td><strong>Punktezahl Account ACCOUNT2</strong></td><td>$PunkteAccount2</td></tr>
<tr><td><strong>Probleme?</strong></td>
<td>Bitte einen <a href="https://example.service-now.com/incident.do?sys_id=-1&amp;sysparm_query=active=true">Incident er&ouml;ffnen</a>
und der Zuweisungsgruppe $Zuweisungsgruppe zuweisen.</td></tr>
</tbody>
</table>
"@

    Send-ReportMail -Subject "$ServiceName SMS FAX $MonatLang" -HtmlBody $mailBody `
        -Attachments $LogFile, $ReportFile

    Write-Log -Message '=== Verrechnung abgeschlossen, Report-Mail versendet ==='
}
catch {
    Write-Log -Level FATAL -Message "Abbruch: $_"
    throw
}
finally {
    if ($graphVerbunden) { Disconnect-MgGraph | Out-Null }
}

#endregion
