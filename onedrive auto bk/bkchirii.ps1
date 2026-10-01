# ==============================================================================
# 1. CONFIGURARE CĂI & PARAMETRI
# ==============================================================================
# Calea către fișierul sursă al bazei de date Chirii
$FisierSursa = "D:\MX_CHIRII\date\MXCHIRII.FDB"

# Calea destinație din OneDrive
$DestinatieLocala = "C:\Users\dasac\OneDrive\CHIRII"

# Folderul local de log-uri
$LogDir = "C:\Logs_Backup_Chirii"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogFile = "$LogDir\copy_chirii_log_$(Get-Date -Format 'yyyy-MM-dd').log"

# Colectare log pentru rularea curentă
$Script:LogCurent = @()

function Write-Log {
    param ([string]$Mesaj)
    $TextCuOra = "$(Get-Date -Format 'HH:mm:ss') - $Mesaj"
    $Script:LogCurent += $TextCuOra
    Add-Content -Path $LogFile -Value $TextCuOra -Encoding UTF8
}

# Ne asigurăm că directorul de destinație există
if (!(Test-Path $DestinatieLocala)) { 
    New-Item -ItemType Directory -Path $DestinatieLocala -Force | Out-Null 
}

# ==============================================================================
# 2. CONFIGURARE E-MAIL INSTITUȚIONAL (mail.dasil.ro)
# ==============================================================================
$SmtpServer            = "nume.domeniu.ro"
$SmtpPort              = 587
$Expeditor             = '"DASIL - Automatizari IT" <nume@dasil.ro>'

# Destinatari (modifică sau adaugă adresele dorite, separate prin virgulă)
$Destinatar            = "test1@dasil.ro, test2@dasil.ro"

$NecesitaAutentificare = $true
$SmtpUser              = "nume@dasil.ro"
$SmtpPass              = "xxx" # Introdu parola contului automatizari@dasil.ro
$UseSSL                = $true

function Send-ReportEmail {
    param (
        [string]$Subiect,
        [string]$MesajStare,
        [bool]$IsSuccess
    )
    
    try {
        Write-Log "INFO: Se trimite e-mailul de notificare către: $Destinatar"
        
        $TextLogCurent = $Script:LogCurent -join "`r`n"
        
        $CorpEmail = @"
=== RAPORT AUTOMATIZAT BACKUP CHIRII - MICROSOFT CLOUD ===
Data/Ora: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Status: $(if($IsSuccess){"SUCCES"}else{"EROARE"})

Detalii executie:
$MesajStare

==================== LOG EXECUTIE RULARE CURENTĂ ====================
$TextLogCurent

-- 
Sistem Automatizat de Backup & Disaster Recovery
Script dezvoltat de: Inspector Informatică Sorin Marian
Departamentul Informatică
"@

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $MailMessage = New-Object System.Net.Mail.MailMessage
        $MailMessage.From = $Expeditor
        $MailMessage.Subject = $Subiect
        $MailMessage.Body = $CorpEmail
        $MailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

        $ListaDestinatari = $Destinatar.Split(',')
        foreach ($Adresa in $ListaDestinatari) {
            $AdresaCurata = $Adresa.Trim()
            if ($AdresaCurata -ne "") {
                $MailMessage.To.Add($AdresaCurata)
            }
        }

        $SmtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $SmtpClient.EnableSsl = $UseSSL

        if ($NecesitaAutentificare) {
            $SmtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPass)
        }

        $SmtpClient.Send($MailMessage)
        Write-Log "INFO: E-mailul de raportare a fost trimis cu succes."
    }
    catch {
        Write-Log "EROARE SMTP: Nu s-a putut trimite e-mailul de raportare: $_"
    }
}

# ==============================================================================
# 3. EXECUȚIE PROCES BACKUP
# ==============================================================================
Write-Log "====================================================================="
Write-Log "Pornire copiere baza de date Chirii la: $(Get-Date)"

$ProcesReusit = $false
$DetaliuRezultat = ""

try {
    # Step 1: Verificare existență fișier sursă
    if (Test-Path $FisierSursa) {
        
        # Denumire fișier destinație cu Timestamp
        $DataTimeStamp = Get-Date -Format "yyyy-MM-dd_HHmm"
        $NumeFisierDestinatie = "MXCHIRII_$DataTimeStamp.FDB"

        Write-Log "Fișier sursă identificat: $FisierSursa"
        Write-Log "Se inițiază copierea în folderul OneDrive..."

        # Copiere fișier cu Robocopy
        $FolderSursa = Split-Path -Path $FisierSursa -Parent
        $NumeFisierSursa = Split-Path -Path $FisierSursa -Leaf

        robocopy "$FolderSursa" "$DestinatieLocala" "$NumeFisierSursa" /R:2 /W:3 /NP /LOG+:"$LogFile" | Out-Null
        
        # Redenumim fișierul proaspăt copiat pentru a include timestamp-ul
        $CaleCopieSimpla = Join-Path -Path $DestinatieLocala -ChildPath $NumeFisierSursa
        if (Test-Path $CaleCopieSimpla) {
            Rename-Item -Path $CaleCopieSimpla -NewName $NumeFisierDestinatie -Force
            Write-Log "Copiere și redenumire finalizate cu SUCCES!"
        }

        # Step 2: POLITICĂ RETENȚIE - Păstrare doar o singură copie (cea mai recentă)
        Write-Log "Se verifică directoarele OneDrive pentru păstrarea unei singure copii..."
        
        $ToateCopiile = Get-ChildItem -Path $DestinatieLocala -Filter "MXCHIRII_*.FDB" | Sort-Object LastWriteTime -Descending

        if ($ToateCopiile.Count -gt 1) {
            # Selectăm toate fișierele cu excepția ultimului adăugat (Skip 1)
            $CopiiDeSters = $ToateCopiile | Select-Object -Skip 1
            foreach ($FisierVechi in $CopiiDeSters) {
                Write-Log "Ștergere copie veche: $($FisierVechi.Name)"
                Remove-Item -Path $FisierVechi.FullName -Force
            }
        }

        $DetaliuRezultat = "Fișierul [$NumeFisierDestinatie] a fost copiat local și transmis în OneDrive. Copiile vechi au fost eliminate."
        $ProcesReusit = $true
    } else {
        Write-Log "EROARE: Fișierul sursă nu a fost găsit la calea: $FisierSursa"
        $DetaliuRezultat = "Baza de date $FisierSursa nu există pe disc."
    }
}
catch {
    Write-Log "EROARE CRITICĂ: $_"
    $DetaliuRezultat = "A apărut o excepție neașteptată în timpul execuției: $_"
}

Write-Log "====================================================================="

# ==============================================================================
# 4. TRIMITERE RAPORT E-MAIL
# ==============================================================================
$DataFormatata = Get-Date -Format 'dd.MM.yyyy HH:mm'

if ($ProcesReusit) {
    Send-ReportEmail -Subiect "[SUCCESS] Backup Chirii sincronizat ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess $true
} else {
    Send-ReportEmail -Subiect "[EROARE] Backup Chirii EȘUAT ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess $false
}