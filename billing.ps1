$hostname = hostname
$month = (get-date).AddMonths(-1).ToString("MM")
$lastmonthfull = (get-date).AddMonths(-1).ToString("Y")
$year = (get-date).AddMonths(-1).Year
$DDMMYYY = Get-Date -format dd.MM.yyyy

# SMTP / Mail settings (anonymized)
$smtpServer = "mail.example.com" # SMTP Server
$smtpFrom   = "reporting@example.com" # Sender
$smtpTo     = "recipient@example.com" # Recipient of report mails

$messageSubject = "SERVICE_XYZ SMS FAX $lastmonthfull"
$securitycheck  = "false"
$message        = $null
$UnknownDomains = @()
$Zuweisungsgruppe = "IT-SUPPORT-GROUP" # Ticket assignment group

# Destination folder (anonymized shares)
$sourcefolder     = "D:\Temp\SCRIPTS\Log_Analyzer"
$sharefolder      = "\\FILESERVER01\Source"
$destinationFolder = "\\FILESERVER01\Source\Import"
$ReportingFolder   = "\\FILESERVER01\Source\Verrechnungsfiles"

# Log function 
$logfile = "{0}\{1}{2}{3}" -f $sharefolder, "\logs\logfile_", (Get-Date -Format "yyyyMMdd"), ".txt"
$logfilecheck = test-path $logfile

if ($logfilecheck -ne "true") {
    new-item $logfile -itemtype file
}

Function write-log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory=$True)]
        [String]
        $Message
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line  = "$Stamp $Level $Message"
    Add-Content -Path $logfile -Value $Line
    write-host $line -foregroundcolor green
}

# Master data - CSV import

$Stammdaten = @{}
$Stammdaten = Import-Csv -path "$sourcefolder\Stammdaten.csv" -Delimiter ";"
write-log -message "Stammdaten.csv wurde eingelesen"
$Stammdaten = [Collections.Generic.List[Object]]$Stammdaten
$Countingtable = Import-Csv -path "$sourcefolder\Stammdaten.csv" -Delimiter ";" | select Innenauftrag, Kostenstelle, Kundenauftrag, Auftragsposition -unique
$Countingtable | Add-Member -MemberType NoteProperty -Name Punkte -value 0

$Nichtverrechenbar = @{}

# Check if the script has already been executed this month
if ($securitycheck -eq "true") {
    $pathcheck = test-path "\\FILESERVER02\Source\Check\Check_$year$month.txt"
    if ($pathcheck -eq "true") {
        write-log -message "The script has already been executed this month. If you want to overrule this check, change the value of the ""securitycheck"" variable in the script settings to anything else than ""true""." -level "ERROR"
        Write-EventLog -LogName "Application" -Source "Application Error" -EventID 1 -Message "The billing script has already been executed this month. If you want to overrule this check, change the value of the ""securitycheck"" variable in the script settings to anything else than ""true""." -EntryType Error
        throw "The script has already been executed this month. If you want to overrule this check, change the value of the ""securitycheck"" variable in the script settings to anything else than ""true""."
    } 
    new-item "\\FILESERVER02\Source\Check\Check_$year$month.txt" -itemtype file
}
	
# Credentials (anonymized / empty placeholders)
$email    = ""
$username = ""
$password = ""
 
# File extensions to download
$extensions = "zip"
 
# load the assembly
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

# Load unzip function from System.IO.Compression.ZipFile
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    param([string]$zipfile, [string]$outpath)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

# Create Exchange Service object
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
$s.Credentials = New-Object Net.NetworkCredential($username, $password)
$s.TraceEnabled = $true
write-log -message "Trying AutoDiscover... "
$s.AutodiscoverUrl($email, {$true})
 
if(!$s.Url) {
    Write-Log -message "AutoDiscover failed" -level "ERROR"
    Write-Error "AutoDiscover failed"
    return;
} else {
    write-log -message "AutoDiscover succeeded - $($s.Url)"
}
 
# Create destination folder
$destinationFolder = "{0}\{1}" -f $destinationFolder, (Get-Date -Format "yyyyMMdd HHmmss")
mkdir $destinationFolder | Out-Null
 
# get a handle to the inbox
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
 
#create a property set (to let us access the body & other details not available from the FindItems call)
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;
 
# Find the items
$inc = 0;
$maxRepeat = 50;
do {
    $maxRepeat -= 1;
 
    write-log -message "Searching for items in mailbox... "
    $items = $inbox.FindItems(100)
    write-log -message "found $($items.items.Count) items in the inbox"
 
    foreach ($item in $items.Items) {
        # Create mail folder
        $inc += 1
        $mailFolder = "{0}\{1}" -f $destinationFolder, $inc;
        mkdir $mailFolder | Out-Null
 
        # load the property set to allow us to get to the body
        try {
            $item.load($psPropertySet)
            write-log -message ("$inc - $($item.Subject)")
 
            # save the metadata to a file
            $item | Export-Clixml ("{0}\metadata.xml" -f $mailFolder)
 
            # save all attachments
            foreach($attachment in $item.Attachments) {
                if(($attachment.Name -split "\." | select -last 1) -in $extensions) {
                    $fileName = ("{0}\{1}" -f $mailFolder, $attachment.Name) -replace "/",""
                    write-log -message "File has been downloaded: $filename - $([Math]::Round($attachment.Size / 1024))KB"
                    $attachment.Load($fileName)
					
                    # Extract files		
                    Try {
                        write-log -message "Unzipping $filename"
                        Unzip $filename $mailfolder 
                    } catch [Exception] {
                        Write-Log -message "Unable to extract item: $filename" -level "ERROR"
                        Write-Error "Unable to extract item: $filename"
                    }
                }
            }
 
            # delete the mail item
            $item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $true)
            write-log -message "Moving processed items to the mailbox dumpster"
        } catch [Exception] {
            Write-Log -message "Unable to load item: $($_)" -level "ERROR"
            Write-Error "Unable to load item: $($_)"
        }	
	
        # CSV Import
	
        $searchinfolder = Get-ChildItem $mailFolder *.txt
        if ($searchinfolder.name -like "Account1_*_Out-LogFile.txt") {
            write-log -message "Einlesen von CSV für OUT-File von Account ACCOUNT1 wurde gestartet"
            $Account1_Out_Logfile = @()
            $Account1_Out_Logfile | Add-Member -MemberType NoteProperty -Name Punkte -value 0		
            $Account1_Out_LogFile = Import-Csv `
                -path "$mailfolder\Account1_*_Out-LogFile.txt" `
                -Delimiter ";"`
                -header "Referenz","Startdatum","Meldung","Resultatcode","Absender","Empfaengernummer","ExterneID","Punkte","Empfaengername" `
                | Select @{Name="Referenz";Expression={[decimal]$_.Referenz}}, @{Name="Startdatum";Expression={[datetime]$_.Startdatum}}, Meldung, `
                @{Name="Resultatcode";Expression={[decimal]$_.Resultatcode}}, Absender, @{Name="Empfaengernummer";Expression={[decimal]$_.Empfaengernummer}}, `
                ExterneID , @{Name="Punkte";Expression={[decimal]$_.Punkte}}, Empfaengername
								
            write-log -message "CSV für OUT-File von Account ACCOUNT1 wurde eingelesen"
        }	
		
        if ($searchinfolder.name -like "Account1_*_In-LogFile.txt") {
            write-log -message "Einlesen von CSV für IN-File von Account ACCOUNT1 wurde gestartet"
            $Account1_In_Logfile = @()
            $Account1_In_Logfile | Add-Member -MemberType NoteProperty -Name Punkte -value 0
            $Account1_In_LogFile = Import-Csv `
                -path "$mailfolder\Account1_*_In-LogFile.txt" `
                -Delimiter ";"`
                -header "Referenz","Startdatum","GesendeteMeldung","EmpfangeneMeldung","Resultatcode","Empfaengernummer","eCallNummer","AntwortAdresse","AntwortInfo","ExterneID","Punkte" `
                | Select @{Name="Referenz";Expression={[decimal]$_.Referenz}}, @{Name="Startdatum";Expression={[datetime]$_.Startdatum}}, GesendeteMeldung, `
                EmpfangeneMeldung, @{Name="Resultatcode";Expression={[decimal]$_.Resultatcode}}, @{Name="Empfaengernummer";Expression={[decimal]$_.Empfaengernummer}}, `
                @{Name="eCallNummer";Expression={[decimal]$_.ecallNummer}} , Antwortadresse, Antwortinfo, ExterneID, @{Name="Punkte";Expression={[decimal]$_.Punkte}}	
								
            write-log -message "CSV für IN-File von Account ACCOUNT1 wurde eingelesen"
        }
		
        if ($searchinfolder.name -like "Account2_*_Out-LogFile.txt") {
            write-log -message "Einlesen von CSV für OUT-File von Account ACCOUNT2 wurde gestartet"
            $Account2_Out_LogFile = @()
            $Account2_Out_LogFile | Add-Member -MemberType NoteProperty -Name Punkte -value 0
            $Account2_Out_LogFile = Import-Csv `
                -path "$mailfolder\Account2_*_Out-LogFile.txt" `
                -Delimiter ";"`
                -header "Referenz","Startdatum","Meldung","Resultatcode","Absender","Empfaengernummer","ExterneID","Punkte","Empfaengername" `
                | Select @{Name="Referenz";Expression={[decimal]$_.Referenz}}, @{Name="Startdatum";Expression={[datetime]$_.Startdatum}}, Meldung, `
                @{Name="Resultatcode";Expression={[decimal]$_.Resultatcode}}, Absender, @{Name="Empfaengernummer";Expression={[decimal]$_.Empfaengernummer}}, `
                ExterneID , @{Name="Punkte";Expression={[decimal]$_.Punkte}}, Empfaengername
								
            write-log -message "CSV für OUT-File von Account ACCOUNT2 wurde eingelesen"
        }
		
        if ($searchinfolder.name -like "Account2_*_In-LogFile.txt") {
            write-log -message "Einlesen von CSV für IN-File von Account ACCOUNT2 wurde gestartet"
            $Account2_In_LogFile = @()
            $Account2_In_LogFile | Add-Member -MemberType NoteProperty -Name Punkte -value 0
            $Account2_In_LogFile = Import-Csv `
                -path "$mailfolder\Account2_*_In-LogFile.txt" `
                -Delimiter ";"`
                -header "Referenz","Startdatum","GesendeteMeldung","EmpfangeneMeldung","Resultatcode","Empfaengernummer","eCallNummer","AntwortAdresse","AntwortInfo","ExterneID","Punkte" `
                | Select @{Name="Referenz";Expression={[decimal]$_.Referenz}}, @{Name="Startdatum";Expression={[datetime]$_.Startdatum}}, GesendeteMeldung, `
                EmpfangeneMeldung, @{Name="Resultatcode";Expression={[decimal]$_.Resultatcode}}, @{Name="Empfaengernummer";Expression={[decimal]$_.Empfaengernummer}}, `
                @{Name="eCallNummer";Expression={[decimal]$_.ecallNummer}} , Antwortadresse, Antwortinfo, ExterneID, @{Name="Punkte";Expression={[decimal]$_.Punkte}}	
																
            write-log -message "CSV für IN-File von Account ACCOUNT2 wurde eingelesen"
        }
    }
} while($items.MoreAvailable -and $maxRepeat -ge 0)

# Import Leistungsarten

$Leistungsarten = Import-Csv -path "$sourcefolder\Leistungsarten.csv" -Delimiter ";"
write-log -message "Leistungsarten.csv wurde eingelesen"

# Start sender / reply-address logic

$Account1_Absender = $Account1_Out_LogFile | select Absender -unique | where-object {$_.Absender -ne ""}
$Account1_Antwortadresse = $Account1_In_LogFile | select AntwortAdresse -unique | where-object {$_.AntwortAdresse -ne ""}

$Account2_Absender = $Account2_Out_LogFile | select Absender -unique | where-object {$_.Absender -ne ""}
$Account2_Domain_Out =  foreach ($item in $Account2_Absender.Absender) {($item -replace '\s','' -split '@')[1]}
$Account2_Domain_Out_Unique = $Account2_Domain_Out | select -unique

$Account2_Antwortadresse = $Account2_In_LogFile | select Antwortadresse -unique # keep empty addresses as well
$Account2_Domain_In =  foreach ($item in $Account2_Antwortadresse.Antwortadresse) {($item -replace '\s','' -split '@')[1]}
$Account2_Domain_In_Unique = $Account2_Domain_In | select -unique
$Account2_Domain_In_Unique += "" # include empty domain for checking

# Account1 sender

foreach ($Account1_Absender_Unique_Out in $Account1_Absender) {
    write-host $Account1_Absender_Unique_Out.Absender
	
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account1_Absender_Unique_Out.Absender} )
	
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
	
    $Punkte = ""		
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account1_Out_LogFile | where {$_.Absender -eq $Account1_Absender_Unique_Out.Absender}
        ).Punkte
    )
    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
	
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}

foreach ($Account1_Antwortadresse_Unique_In in $Account1_Antwortadresse) {
    write-host $Account1_Antwortadresse_Unique_In.AntwortAdresse
	
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account1_Antwortadresse_Unique_In.AntwortAdresse} )
	
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
	
    $Punkte = ""
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account1_In_LogFile | where {$_.Antwortadresse -eq $Account1_Antwortadresse_Unique_In.AntwortAdresse}
        ).Punkte
    )
	
    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
	
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}	
	
foreach ($Account2_Domain_Out_SingleItem in $Account2_Domain_Out_Unique) {
    write-host $Account2_Domain_Out_SingleItem
	
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account2_Domain_Out_SingleItem} )
	
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
	
    $Punkte = ""	
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account2_Out_LogFile | where {$_.Absender -like "*$Account2_Domain_Out_SingleItem*"}
        ).Punkte
    )
	
    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
	
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}
		
foreach ($Account2_Domain_In_SingleItem in $Account2_Domain_In_Unique) {
    write-host $Account2_Domain_In_SingleItem
	
    $Index = $Stammdaten.findindex( {$args[0].Antwortadresse -eq $Account2_Domain_Out_SingleItem} )
		
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
		
    if ($Account2_Domain_In_SingleItem -ne "") {
        $Punkte = ""
        $Punkte = [Linq.Enumerable]::Sum(
            [decimal[]] (
                $Account2_In_LogFile | where {$_.Antwortadresse -like "*$Account2_Domain_In_SingleItem*"}
            ).Punkte
        )
        $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
			
        $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
        $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
        $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
        $Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
        write-host $Innenauftrag
        write-host $Kostenstelle
        write-host $Kundenauftrag
        write-host $Auftragsposition
        write-host $Punkte
			
    } else {
        write-host $Account2_Domain_In_SingleItem
        $Punkte = ""
        $Punkte = [Linq.Enumerable]::Sum(
            [decimal[]] (
                $Account2_In_LogFile | where {$_.Antwortadresse -eq ""}
            ).Punkte
        )
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )		
		
        $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
			
        $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
        $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
        $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
        $Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
        write-host $Innenauftrag
        write-host $Kostenstelle
        write-host $Kundenauftrag
        write-host $Auftragsposition
        write-host $Punkte
    }
}
		
# Start ExterneID billing

$Account1_ExterneID_Out = $Account1_Out_LogFile | select ExterneID -unique | where-object {$_.ExterneID -ne "" -and $_.ExterneID -notlike "*ecallURL*"}
$Account1_ExterneID_In  = $Account1_In_LogFile  | select ExterneID -unique | where-object {$_.ExterneID -ne "" -and $_.ExterneID -notlike "*ecallURL*"}

$Account2_ExterneID_Out = $Account2_Out_LogFile | select ExterneID -unique | where-object {$_.ExterneID -ne "" -and $_.ExterneID -notlike "*ecallURL*"}
$Account2_ExterneID_In  = $Account2_In_LogFile  | select ExterneID -unique | where-object {$_.ExterneID -ne "" -and $_.ExterneID -notlike "*ecallURL*"}

foreach ($Account1_ExterneID_Unique_Out in $Account1_ExterneID_Out) {
    write-host $Account1_ExterneID_Unique_Out.ExterneID -foregroundcolor green
		
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account1_ExterneID_Unique_Out.ExterneID} )
		
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
		
    $Punkte = ""
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account1_Out_LogFile | where {$_.ExterneID -eq $Account1_ExterneID_Unique_Out.ExterneID}
        ).Punkte
    )
		
    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}
		
# eCall URL points add (Account1_Out_LogFile)		
write-host "eCallURL Addition" -foregroundcolor yellow
		
$Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "eCallURL"} )
$Punkte = ""
$Punkte = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Account1_Out_LogFile | where {$_.ExterneID -like "eCallURL*"}
    ).Punkte
)
		
$Kostenstelle   = $Stammdaten[$Index].Kostenstelle
$Innenauftrag   = $Stammdaten[$Index].Innenauftrag
$Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
$Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
write-host $Innenauftrag
write-host $Kostenstelle
write-host $Kundenauftrag
write-host $Auftragsposition
write-host $Punkte
$Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge

foreach ($Account1_ExterneID_Unique_In in $Account1_ExterneID_In) {
    write-host $Account1_ExterneID_Unique_In.ExterneID -foregroundcolor green
		
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account1_ExterneID_Unique_In.ExterneID} )
		
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
		
    $Punkte = ""
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account1_In_LogFile | where {$_.ExterneID -eq $Account1_ExterneID_Unique_In.ExterneID}
        ).Punkte
    )
		
    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}

# eCall URL points add (Account1_In_LogFile)	
write-host "eCallURL Addition" -foregroundcolor yellow
		
$Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "eCallURL"} )
$Punkte = ""
$Punkte = [Linq.Enumerable]::Sum(
    [decimal[]] @(
        $Account1_In_LogFile | where {$_.ExterneID -like "eCallURL*"}
    ).Punkte
)
		
$Kostenstelle   = $Stammdaten[$Index].Kostenstelle
$Innenauftrag   = $Stammdaten[$Index].Innenauftrag
$Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
$Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
write-host $Innenauftrag
write-host $Kostenstelle
write-host $Kundenauftrag
write-host $Auftragsposition
$Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
write-host $Punkte		
	
foreach ($Account2_ExterneID_Unique_Out in $Account2_ExterneID_Out) {
    write-host $Account2_ExterneID_Unique_Out.ExterneID -foregroundcolor green
		
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account2_ExterneID_Unique_Out.ExterneID} )
		
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
		
    $Punkte = ""		
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account2_Out_LogFile | where {$_.ExterneID -eq $Account2_ExterneID_Unique_Out.ExterneID}
        ).Punkte
    )

    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}

# eCall URL points add (Account2_Out_LogFile)	
write-host "eCallURL Addition" -foregroundcolor yellow
		
$Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "eCallURL"} )
$Punkte = ""
$Punkte = [Linq.Enumerable]::Sum(
    [decimal[]] @(
        $Account2_Out_LogFile | where {$_.ExterneID -like "eCallURL*"}
    ).Punkte
)
		
$Kostenstelle   = $Stammdaten[$Index].Kostenstelle
$Innenauftrag   = $Stammdaten[$Index].Innenauftrag
$Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
$Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
write-host $Innenauftrag
write-host $Kostenstelle
write-host $Kundenauftrag
write-host $Auftragsposition
$Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
write-host $Punkte		
	
foreach ($Account2_ExterneID_Unique_In in $Account2_ExterneID_In) {
    write-host $Account2_ExterneID_Unique_In.ExterneID -foregroundcolor green
		
    $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq $Account2_ExterneID_Unique_In.ExterneID} )
		
    if ($Index -eq "-1") {
        $Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )
    }
		
    $Punkte = ""		
    $Punkte = [Linq.Enumerable]::Sum(
        [decimal[]] (
            $Account2_In_LogFile | where {$_.ExterneID -eq $Account2_ExterneID_Unique_In.ExterneID}
        ).Punkte
    )
		
    $Kostenstelle   = $Stammdaten[$Index].Kostenstelle
    $Innenauftrag   = $Stammdaten[$Index].Innenauftrag
    $Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
    $Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
    write-host $Innenauftrag
    write-host $Kostenstelle
    write-host $Kundenauftrag
    write-host $Auftragsposition
    $Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
    write-host $Punkte
}
	
# eCall URL points add (Account2_In_LogFile)	
write-host "eCallURL Addition" -foregroundcolor yellow
		
$Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "eCallURL"} )
$Punkte = ""
$Punkte = [Linq.Enumerable]::Sum(
    [decimal[]] @(
        $Account2_In_LogFile | where {$_.ExterneID -like "eCallURL*"}
    ).Punkte
)
		
$Kostenstelle   = $Stammdaten[$Index].Kostenstelle
$Innenauftrag   = $Stammdaten[$Index].Innenauftrag
$Kundenauftrag  = $Stammdaten[$Index].Kundenauftrag
$Auftragsposition = $Stammdaten[$Index].Auftragsposition
		
write-host $Innenauftrag
write-host $Kostenstelle
write-host $Kundenauftrag
write-host $Auftragsposition
$Stammdaten[$Index].Menge = $Punkte + $Stammdaten[$Index].Menge
write-host $Punkte		
	
# Total points checks

$GesamtpunkteAccount1Out = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Account1_Out_LogFile
    ).Punkte
)		

$GesamtpunkteAccount1In = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Account1_In_LogFile
    ).Punkte
)		

$GesamtpunkteAccount2In = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Account2_In_LogFile
    ).Punkte
)		
		
$GesamtpunkteAccount2Out = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Account2_Out_LogFile
    ).Punkte
)	
		
$Gesamtpunkte = $GesamtpunkteAccount1Out + $GesamtpunkteAccount1In + $GesamtpunkteAccount2In + $GesamtpunkteAccount2Out
$PunktzahlAccount1 = $GesamtpunkteAccount1Out + $GesamtpunkteAccount1In
$PunktezahlAccount2 = $GesamtpunkteAccount2Out + $GesamtpunkteAccount2In

write-host $Gesamtpunkte -foregroundcolor yellow
		
# Compare with script points
		
$Scriptpunkte = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Stammdaten
    ).Menge
)	
		
write-host $Scriptpunkte -foregroundcolor yellow

$stammdaten | ft

$Punktedifferenz = $Gesamtpunkte - $Scriptpunkte

write-host "Punktedifferenz: $Punktedifferenz" -foregroundcolor yellow
		
# Assign difference to default.example.com

$Index = $Stammdaten.findindex( {$args[0].Buchungstext -eq "default.example.com"} )		
$Stammdaten[$Index].Menge = $Punktedifferenz + $Stammdaten[$Index].Menge

$Scriptpunkte = [Linq.Enumerable]::Sum(
    [decimal[]] (
        $Stammdaten
    ).Menge
)	

$PunktedifferenzNEU = $Gesamtpunkte - $Scriptpunkte

write-host "Punktedifferenz NEU: $PunktedifferenzNEU" -foregroundcolor green

# Billing CSV "metadata" header
new-item "$reportingfolder\SERVICE_XYZ_SMS_FAX_$year$month.csv" -itemtype file -force -value "SKST;XXXX Communication Services
Modul;XXXXXX SMS/FAX
Erstellung;$DDMMYYY
Leistungsart;Leistungsgroessen
$($Leistungsarten.Leistungsart[0]);$($Leistungsarten.Bezeichnung[0])
$($Leistungsarten.Leistungsart[1]);$($Leistungsarten.Bezeichnung[1])
$($Leistungsarten.Leistungsart[2]);$($Leistungsarten.Bezeichnung[2])
$($Leistungsarten.Leistungsart[3]);$($Leistungsarten.Bezeichnung[3])
"

$Stammdaten | Export-Csv "$reportingfolder\Temp\Stammdaten_$year$month.csv" ";" -NoTypeInformation
$Stammdaten_Werte = get-content "$reportingfolder\Temp\Stammdaten_$year$month.csv"

# Append data rows to billing CSV
add-content -Path "$reportingfolder\SERVICE_XYZ_SMS_FAX_$year$month.csv" -Value $Stammdaten_Werte

# Build HTML mail body (anonymized)
$messagebody = 
"<h1 style='color: #5e9ca0;'>SERVICE_XYZ Periode <span style='color: #2b2301;'>$lastmonthfull</span></h1>
<h2 style='color: #2e6c80;'>Quick-Infos:</h2>
<table style='height: 105px; width: 606px;' border='2'>
<tbody>
<tr>
<td style='width: 265px;'><strong>Auslösedatum</strong></td>
<td style='width: 337px;'>$DDMMYYY</td>
</tr>
<tr>
<td style='width: 265px;'><strong>Generierender Server</strong></td>
<td style='width: 337px;'>$hostname</td>
</tr>
<tr>
<td style='width: 265px;'><strong>Punktezahl Gesamt</strong></td>
<td style='width: 337px;'>$Gesamtpunkte</td>
</tr>
<tr>
<td style='width: 265px;'><strong>Punktezahl Account&nbsp;ACCOUNT1</strong></td>
<td style='width: 337px;'>$PunktzahlAccount1</td>
</tr>
<tr>
<td style='width: 265px;'><strong>Punktezahl Account ACCOUNT2</strong></td>
<td style='width: 337px;'>$PunktezahlAccount2</td>
</tr>
<tr>
<td style='width: 265px;'><strong>Probleme?</strong></td>
<td style='width: 337px;'>Bitte einen <a href='https://example.service-now.com/incident.do?sys_id=-1&amp;sysparm_query=active=true&amp;sysparm_stack=incident_list.do?sysparm_query=active=true'>Incident eröffnen</a>&nbsp;und der Zuweisungsgruppe $Zuweisungsgruppe zuweisen.</td>
</tr>
</tbody>
</table>
<h2 style='color: #2e6c80;'>&nbsp;</h2>"

# Sending Email

Send-MailMessage -To $smtpTo -From $smtpFrom -Subject $messageSubject -Body $messageBody -BodyAsHTML -SmtpServer $smtpServer -encoding UTF8 -attachments $logfile, "$reportingfolder\SERVICE_XYZ_SMS_FAX_$year$month.csv"
