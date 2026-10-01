# ==============================================================================
# 1. CONFIGURARE CĂI & LOGARE (Compatibil PS 3.0 / Windows Server 2012)
# ==============================================================================
$ExePath = "C:\Program Files (x86)\Adi-Com SOFT\Copia de Siguranta\AcsBk.exe"

$LogDir  = "C:\Logs_Backup_Server"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogFile = "$LogDir\generate_log_$(Get-Date -Format 'yyyy-MM-dd').log"

# Array simplu compatibil cu versiuni vechi de PowerShell
$Script:LogCurent = @()

function Write-ServerLog {
    param ([string]$Mesaj)
    $TextCuOra = "$(Get-Date -Format 'HH:mm:ss') - $Mesaj"
    $Script:LogCurent += $TextCuOra
    Add-Content -Path $LogFile -Value $TextCuOra -Encoding UTF8
}

# ==============================================================================
# 2. CONFIGURARE E-MAIL INSTITUȚIONAL (cPanel / mail.dasil.ro)
# ==============================================================================
$SmtpServer            = "nume.domeniu.ro"
$SmtpPort              = 587

# Modificat: Nume afișat + Noua adresă de e-mail
$Expeditor             = '"DAS-Automatizari IT" <nume@dasil.ro>'

# Destinatari multipli separați prin virgulă
$Destinatar            = "test1@dasil.ro, test2@dasil.ro"

$NecesitaAutentificare = $true

# Modificat: Contul de autentificare SMTP (noua adresă)
$SmtpUser              = "nume@dasil.ro"
$SmtpPass              = "xxx" # Introdu noua parolă a contului 
$UseSSL                = $true

function Send-ServerReportEmail {
    param (
        [string]$Subiect,
        [string]$MesajStare,
        [bool]$IsSuccess
    )
    
    try {
        Write-ServerLog "INFO: Se trimite e-mailul de notificare către: $Destinatar"
        
        $TextLogCurent = $Script:LogCurent -join "`r`n"
        
        $CorpEmail = @"
=== RAPORT GENERARE BACKUP ADICOM (SERVER 2012) ===
Data/Ora: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Server: $env:COMPUTERNAME
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

        # Setare TLS 1.2 pentru Windows Server 2012
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
        Write-ServerLog "INFO: E-mailul de raportare a fost trimis cu succes."
    }
    catch {
        Write-ServerLog "EROARE SMTP: Nu s-a putut trimite e-mailul de raportare: $_"
    }
}

# ==============================================================================
# 3. EXECUȚIE AUTOMATIZARE INTERFAȚĂ (GUI)
# ==============================================================================
Write-ServerLog "====================================================================="
Write-ServerLog "Inițiere generare automată backup Adi-Com pe Server 2012..."

$ProcesReusit = $false
$DetaliuRezultat = ""

try {
    if (!(Test-Path $ExePath)) {
        throw "Executabilul nu a fost găsit la calea: $ExePath"
    }

    # 1. Lansare proces
    Write-ServerLog "Lansare executabil AcsBk.exe..."
    $process = Start-Process -FilePath $ExePath -PassThru
    
    if (-not $process -or -not $process.Id) {
        throw "Nu s-a putut obține ID-ul de proces pentru AcsBk.exe."
    }

    Start-Sleep -Seconds 3

    $wshell = New-Object -ComObject WScript.Shell

    # Setăm focus pe fereastră
    $wshell.AppActivate($process.Id)
    Start-Sleep -Milliseconds 500

    # --- FEREASTRA 1 (Bun găsit) ---
    Write-ServerLog "Apasare Enter pe Fereastra 1..."
    $wshell.SendKeys('{ENTER}')

    # Pauză de 7 secunde pentru procesarea de după primul pas
    Start-Sleep -Seconds 7

    # --- FEREASTRA 2 ---
    Write-ServerLog "Navigare și confirmare Fereastra 2..."
    $wshell.AppActivate($process.Id)
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys('{TAB}')
    Start-Sleep -Milliseconds 200
    $wshell.SendKeys('{ENTER}')

    Start-Sleep -Seconds 2

    # --- FEREASTRA 3 ---
    Write-ServerLog "Navigare și confirmare Fereastra 3..."
    $wshell.AppActivate($process.Id)
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys('{TAB}')
    Start-Sleep -Milliseconds 200
    $wshell.SendKeys('{ENTER}')

    Start-Sleep -Seconds 2

    # --- FEREASTRA 4 ---
    Write-ServerLog "Navigare și confirmare Fereastra 4..."
    $wshell.AppActivate($process.Id)
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys('{TAB}')
    Start-Sleep -Milliseconds 200
    $wshell.SendKeys('{ENTER}')

    # --- PASUL FINAL ---
    Start-Sleep -Milliseconds 500
    Write-ServerLog "Confirmare finală (Enter)... Generarea arhivelor a fost inițiată."
    $wshell.SendKeys('{ENTER}')

    # ==========================================================================
    # 4. MONITORIZARE PROCESARE & ÎNCHIDERE APLICAȚIE
    # ==========================================================================
    Write-ServerLog "Așteptare finalizare procesare baze de date (AcsBk.exe)..."
    
    # Oferim o pauză inițială de 30 secunde până când programul intră în lucru intens
    Start-Sleep -Seconds 30

    $MaxWaitMinutes = 45 # Limită de siguranță (45 de minute)
    $WaitCounter = 0

    # Așteptăm până la 45 min, verificând la fiecare minut
    while ((Get-Process -Name "AcsBk" -ErrorAction SilentlyContinue) -and ($WaitCounter -lt $MaxWaitMinutes)) {
        Start-Sleep -Seconds 60
        $WaitCounter++
        if ($WaitCounter % 5 -eq 0) {
            Write-ServerLog "Procesare în curs... (Minutul $WaitCounter)"
        }
    }

    # Închidere curată a procesului dacă încă mai există
    $ProcessCheck = Get-Process -Name "AcsBk" -ErrorAction SilentlyContinue
    if ($ProcessCheck) {
        Write-ServerLog "Procesarea s-a încheiat. Închidere aplicație AcsBk.exe..."
        
        # 1. Încercăm închidere politicoasă din fereastră
        $ProcessCheck.CloseMainWindow() | Out-Null
        Start-Sleep -Seconds 3
        
        # 2. Dacă nu s-a închis, o oprim direct
        if (Get-Process -Name "AcsBk" -ErrorAction SilentlyContinue) {
            Stop-Process -Name "AcsBk" -Force -ErrorAction SilentlyContinue
        }
    }

    Write-ServerLog "Generare backup Adi-Com finalizată și aplicație închisă cu succes!"
    $DetaliuRezultat = "Aplicația AcsBk.exe a procesat complet fișierele și a fost închisă automat pe Server 2012."
    $ProcesReusit = $true
}
catch {
    Write-ServerLog "EROARE la generarea backup-ului GUI: $_"
    $DetaliuRezultat = "Procesul de generare GUI pe server a eșuat: $_"
}

Write-ServerLog "====================================================================="

# ==============================================================================
# 5. TRIMITERE RAPORT E-MAIL
# ==============================================================================
$DataFormatata = Get-Date -Format 'dd.MM.yyyy HH:mm'

if ($ProcesReusit) {
    Send-ServerReportEmail -Subiect "[SERVER SUCCESS] Generare Backup Adicom ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess $true
} else {
    Send-ServerReportEmail -Subiect "[SERVER EROARE] Generare Backup Adicom ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess $false
}