# ==============================================================================
# 1. CONFIGURARE CĂI & LOGARE (Compatibil PS 3.0 / Windows Server 2012)
# ==============================================================================
$ExePath       = "C:\Program Files (x86)\VendorApp\Backup\AcsBk.exe"
$BackupRootDir = "D:\Copia de siguranta"

$LogDir  = "C:\Logs_Backup_Server"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force \vert{} Out-Null }$LogFile = "$LogDir\generate_log_$(Get-Date -Format 'yyyy-MM-dd').log"

# Array compatibil cu versiuni vechi de PowerShell
$Script:LogCurent = @()

function Write-ServerLog {
    param ([string]$Mesaj)$TextCuOra = "$(Get-Date -Format 'HH:mm:ss') -$Mesaj"
    $Script:LogCurent +=$TextCuOra
    Add-Content -Path $LogFile -Value$TextCuOra -Encoding UTF8
}

# ==============================================================================
# 2. CONFIGURARE E-MAIL INSTITUȚIONAL (SMTP Configuration)
# ==============================================================================
$SmtpServer            = "mail.yourdomain.com"
$SmtpPort              = 587$Expeditor             = '"Backup Service" <noreply@yourdomain.com>'
$Destinatar            = "admin@yourdomain.com, management@yourdomain.com"
$NecesitaAutentificare =$true

# Recomandare securitate: Folosiți variabile de mediu sau un fișier securizat/DPAPI în producție
$SmtpUser              = "noreply@yourdomain.com"
$SmtpPass              = "YOUR_SECURE_PASSWORD_HERE"
$UseSSL                =$true

function Send-ServerReportEmail {
    param (
        [string]$Subiect,
        [string]$MesajStare,
        [bool]$IsSuccess
    )
    
    try {
        Write-ServerLog "INFO: Se trimite e-mailul de notificare către: $Destinatar"
        
        $TextLogCurent =$Script:LogCurent -join "`r`n"
        
        $CorpEmail = @"
=== RAPORT GENERARE BACKUP AUTOMATIZAT ===
Data/Ora: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Server: $env:COMPUTERNAME
Status: $(if($IsSuccess){"SUCCES"}else{"EROARE"})

Detalii executie:
$MesajStare

==================== LOG EXECUTIE RULARE CURENTĂ ====================
$TextLogCurent

-- 
Sistem Automatizat de Backup & Disaster Recovery
Departamentul IT
"@

        # Forțare protocol TLS 1.2 pentru compatibilitate Windows Server 2012
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $MailMessage = New-Object System.Net.Mail.MailMessage
        $MailMessage.From = $Expeditor$MailMessage.Subject = $Subiect$MailMessage.Body = $CorpEmail$MailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

        $ListaDestinatari =$Destinatar.Split(',')
        foreach ($Adresa in$ListaDestinatari) {
            $AdresaCurata =$Adresa.Trim()
            if ($AdresaCurata -ne "") {
                $MailMessage.To.Add($AdresaCurata)
            }
        }

        $SmtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
        $SmtpClient.EnableSsl =$UseSSL

        if ($NecesitaAutentificare) {$SmtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser,$SmtpPass)
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
Write-ServerLog "Inițiere generare automată backup pe $env:COMPUTERNAME..."

$ProcesReusit = $false$DetaliuRezultat = ""

# Marcăm ora de start a procesului pentru detectarea folderului nou creat
$ScriptStartTime = Get-Date

try {
    if (!(Test-Path $ExePath)) {
        throw "Executabilul nu a fost găsit la calea: $ExePath"
    }

    if (!(Test-Path $BackupRootDir)) {
        throw "Calea de backup '$BackupRootDir' nu există pe server!"
    }

    # 1. Lansare proces
    Write-ServerLog "Lansare executabil aplicatie..."
    $process = Start-Process -FilePath$ExePath -PassThru
    
    if (-not $process -or -not$process.Id) {
        throw "Nu s-a putut obține ID-ul de proces pentru executabil."
    }

    Start-Sleep -Seconds 3

    $wshell = New-Object -ComObject WScript.Shell

    # Focus pe fereastra aplicației
    $wshell.AppActivate($process.Id)
    Start-Sleep -Milliseconds 500

    # --- FEREASTRA 1 ---
    Write-ServerLog "Apasare Enter pe Fereastra 1..."
    $wshell.SendKeys('{ENTER}')

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
    Write-ServerLog "Așteptare finalizare procesare baze de date..."
    
    Start-Sleep -Seconds 30

    $MaxWaitMinutes = 45$WaitCounter = 0

    $ProcessName = [System.IO.Path]::GetFileNameWithoutExtension($ExePath)

    while ((Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) -and ($WaitCounter -lt$MaxWaitMinutes)) {
        Start-Sleep -Seconds 60
        $WaitCounter++
        if ($WaitCounter % 5 -eq 0) {
            Write-ServerLog "Procesare în curs... (Minutul $WaitCounter)"
        }
    }

    # Închidere curată a procesului
    $ProcessCheck = Get-Process -Name$ProcessName -ErrorAction SilentlyContinue
    if ($ProcessCheck) {
        Write-ServerLog "Procesarea s-a încheiat. Închidere aplicație..."
        $ProcessCheck.CloseMainWindow() | Out-Null
        Start-Sleep -Seconds 3
        
        if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
            Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
        }
    }

    # ==========================================================================
    # 5. VERIFICARE EXISTENȚĂ DIRECTOR NOU ÎN DESTINAȚIA DE BACKUP
    # ==========================================================================
    Write-ServerLog "Verificare existență director nou în: $BackupRootDir..."

    # Căutăm directoare create/modificate DUPĂ începerea rulării scriptului
    $NewFolder = Get-ChildItem -Path$BackupRootDir -Directory | 
                 Where-Object { $_.CreationTime -ge$ScriptStartTime -or $_.LastWriteTime -ge$ScriptStartTime } | 
                 Sort-Object CreationTime -Descending | 
                 Select-Object -First 1

    if ($NewFolder) {
        Write-ServerLog "CONFIRMARE: Folderul nou de backup a fost găsit: '$($NewFolder.Name)' (Creat la: $($NewFolder.CreationTime.ToString('HH:mm:ss')))."
        $DetaliuRezultat = "Aplicația a rulat cu succes, iar folderul de backup '$($NewFolder.Name)' a fost creat corect în '$BackupRootDir'."
        $ProcesReusit =$true
    } else {
        Write-ServerLog "EROARE CRITICĂ: Aplicația s-a închis, dar NU a fost găsit niciun folder nou creat în '$BackupRootDir' după ora $($ScriptStartTime.ToString('HH:mm:ss'))!"
        $DetaliuRezultat = "EROARE: Interfața GUI a procesat comenzile, dar la final NU s-a generat niciun director nou de backup în '$BackupRootDir'."
        $ProcesReusit =$false
    }

}
catch {
    Write-ServerLog "EROARE la generarea backup-ului GUI: $_"
    $DetaliuRezultat = "Procesul de generare GUI pe server a eșuat: $_"
    $ProcesReusit =$false
}

Write-ServerLog "====================================================================="

# ==============================================================================
# 6. TRIMITERE RAPORT E-MAIL
# ==============================================================================
$DataFormatata = Get-Date -Format 'dd.MM.yyyy HH:mm'

if ($ProcesReusit) {
    Send-ServerReportEmail -Subiect "[SERVER SUCCESS] Generare Backup Automata ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess$true
} else {
    Send-ServerReportEmail -Subiect "[SERVER EROARE] Generare Backup Automata ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess$false
}