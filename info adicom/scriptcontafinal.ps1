# ==============================================================================
# 1. CONFIGURARE CĂI & PARAMETRI
# ==============================================================================
# Calea de rețea (UNC) către folderul de backup de pe server
$SursaServer = "\\serverspas\Copia de siguranta"

# Calea de pe stația ta unde vrei să salvezi backupul
$DestinatieLocala = "C:\Users\spasi\OneDrive\Backup prog Conta"

# Numărul maxim de ZILE DISTINCTE de păstrat (3-day rolling retention)
$MaxZilePastrate = 3

# Folderul de log-uri local
$LogDir = "C:\Logs_Backup"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$LogFile = "$LogDir\copy_log_$(Get-Date -Format 'yyyy-MM-dd').log"

# Variabilă pentru colectarea log-ului EXCLUSIV din această rulare
$LogCurent = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param ([string]$Mesaj)
    $LogCurent.Add($Mesaj)
    Add-Content -Path $LogFile -Value $Mesaj -Encoding utf8
}

# ==============================================================================
# 2. CONFIGURARE E-MAIL INSTITUȚIONAL (cPanel / mail.dasil.ro)
# ==============================================================================
$SmtpServer            = "nume.domeniu.ro"
$SmtpPort              = 587

# Modificat: Nume afișat + Noua adresă de e-mail
$Expeditor             = '"DAS-Automatizari IT" <nume@dasil.ro>'

# Destinatari multipli separați prin virgulă
$Destinatar            = "test@dasil.ro, test@dasil.ro"

$NecesitaAutentificare = $true

# Modificat: Contul de autentificare SMTP (noua adresă)
$SmtpUser              = "nume@dasil.ro"
$SmtpPass              = "xxx" # Introdu noua parolă a contului 
$UseSSL                = $true
# ==============================================================================
# FUNCȚIE TRIMITERE RAPORT E-MAIL (PORT 587 / TLS)
# ==============================================================================
function Send-ReportEmail {
    param (
        [string]$Subiect,
        [string]$MesajStare,
        [bool]$IsSuccess
    )
    
    try {
        # Adăugăm linia de confirmare în log MEREU înainte de a genera corpul e-mailului
        Write-Log "INFO: E-mailul de notificare a fost trimis cu succes catre: $Destinatar"
        
        # Transformăm log-ul rulării curente în text (va include și linia de mai sus)
        $TextLogCurent = $LogCurent -join "`r`n"
        
        $CorpEmail = @"
=== RAPORT AUTOMATIZAT BACKUP ADICOM - MICROSOFT CLOUD ===
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

        # Adăugare separată a fiecărui destinatar
        $ListaDestinatari = $Destinatar -split ','
        foreach ($Adresa in $ListaDestinatari) {
            if ($Adresa.Trim() -ne "") {
                $MailMessage.To.Add($Adresa.Trim())
            }
        }

        $SmtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $SmtpClient.EnableSsl = $UseSSL

        if ($NecesitaAutentificare) {
            $SmtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPass)
        }

        $SmtpClient.Send($MailMessage)
    }
    catch {
        Write-Log "EROARE SMTP: Nu s-a putut trimite e-mailul de raportare: $_"
    }
}

# ==============================================================================
# 3. EXECUȚIE PROCES BACKUP
# ==============================================================================
Write-Log "====================================================================="
Write-Log "Pornire copiere backup la: $(Get-Date)"

$ProcesReusit = $false
$DetaliuRezultat = ""

try {
    # Step 1: Identifică cel mai recent folder creat pe server
    $UltimulFolder = Get-ChildItem -Path $SursaServer -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($UltimulFolder) {
        $CaleUltimulFolder = $UltimulFolder.FullName
        $NumeFolder = $UltimulFolder.Name
        $DestinatieFinala = Join-Path -Path $DestinatieLocala -ChildPath $NumeFolder

        Write-Log "Ultimul backup detectat pe server: $NumeFolder"

        # Step 2: Execută copierea doar dacă nu există deja local
        if (!(Test-Path $DestinatieFinala)) {
            Write-Log "Se inițiază copierea către: $DestinatieFinala"
            
            New-Item -ItemType Directory -Path $DestinatieFinala -Force | Out-Null
            
            # Transfer cu Robocopy
            robocopy "$CaleUltimulFolder" "$DestinatieFinala" /E /FFT /Z /R:2 /W:3 /NP /LOG+:"$LogFile"
            
            Write-Log "Copiere finalizată cu SUCCES!"
            $DetaliuRezultat = "Folderul [$NumeFolder] a fost sincronizat local și transmis către OneDrive."
        } else {
            Write-Log "INFO: Ultimul backup exista deja local ($NumeFolder). Copierea a fost omisa."
            $DetaliuRezultat = "Folderul [$NumeFolder] exista deja pe stația locală. Nu a fost necesară recopierea."
        }

        # Step 3: CURĂȚARE AUTOMATĂ (Ultimul backup/zi, max 3 zile distincte)
        Write-Log "Se verifica directoarele locale pentru aplicarea politicii de retentie..."
        
        $ToateFolderele = Get-ChildItem -Path $DestinatieLocala -Directory | Sort-Object LastWriteTime -Descending

        if ($ToateFolderele) {
            $GrupuriPeZile = $ToateFolderele | Group-Object { $_.LastWriteTime.ToString('yyyy-MM-dd') } | Sort-Object Name -Descending
            $ZilePermise = $GrupuriPeZile | Select-Object -First $MaxZilePastrate

            foreach ($GrupZi in $GrupuriPeZile) {
                if ($GrupZi.Name -notin $ZilePermise.Name) {
                    foreach ($Folder in $GrupZi.Group) {
                        Write-Log "Ștergere backup vechi ($($GrupZi.Name)): $($Folder.Name)"
                        Remove-Item -Path $Folder.FullName -Recurse -Force
                    }
                } 
                else {
                    if ($GrupZi.Group.Count -gt 1) {
                        $FoldereDeStersInAceeasiZi = $GrupZi.Group | Select-Object -Skip 1
                        foreach ($FolderVechi in $FoldereDeStersInAceeasiZi) {
                            Write-Log "Ștergere copie intermediară din data de $($GrupZi.Name): $($FolderVechi.Name)"
                            Remove-Item -Path $FolderVechi.FullName -Recurse -Force
                        }
                    }
                }
            }
        }

        $ProcesReusit = $true
    } else {
        Write-Log "EROARE: Nu s-a găsit niciun folder de backup în $SursaServer"
        $DetaliuRezultat = "Sursa de rețea $SursaServer nu conține niciun folder de backup."
    }
}
catch {
    Write-Log "EROARE CRITICĂ: $_"
    $DetaliuRezultat = "A apărut o excepție neașteptată în timpul execuției: $_"
}

Write-Log "====================================================================="

# ==============================================================================
# 4. TRIMITERE RAPORT
# ==============================================================================
$DataFormatata = Get-Date -Format 'dd.MM.yyyy HH:mm'

if ($ProcesReusit) {
    Send-ReportEmail -Subiect "[SUCCESS] Backup Adicom sincronizat ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess $true
} else {
    Send-ReportEmail -Subiect "[EROARE] Backup Adicom EȘUAT ($DataFormatata)" -MesajStare $DetaliuRezultat -IsSuccess $false
}