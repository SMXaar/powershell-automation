# ==============================================================================
# SCRIPT NOTIFICARE DARK MODE + E-MAIL - VERIFICARE ADI-COM
# Inspector Informatică - DASIL Slobozia
# ==============================================================================

# 1. PARAMETRI & VERIFICARE VECHIME BACKUP ADI-COM
$CaleCopiaDeSiguranta = "E:\Copia de Siguranta"
$PragZileAvertizare   = 10

# Folder log-uri
$LogDir = "C:\Logs_Backup_Aplicatie"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogFile = "$LogDir\check_notification_$(Get-Date -Format 'yyyy-MM-dd').log"

$Script:LogCurent = @()

function Write-Log {
    param ([string]$Mesaj)
    $TextCuOra = "$(Get-Date -Format 'HH:mm:ss') - $Mesaj"
    $Script:LogCurent += $TextCuOra
    Add-Content -Path $LogFile -Value $TextCuOra -Encoding UTF8
}

# Verificăm dacă directorul rădăcină există
if (!(Test-Path $CaleCopiaDeSiguranta)) {
    Write-Log "EROARE: Calea $CaleCopiaDeSiguranta nu există."
    exit 0
}

# Căutăm cel mai recent subdirector creat/modificat de aplicația Adi-Com
$UltimulBackup = Get-ChildItem -Path $CaleCopiaDeSiguranta -Directory | 
                 Sort-Object LastWriteTime -Descending | 
                 Select-Object -First 1

$ZileTrecute = 0
$DataUltimuluiBackupText = ""

if ($UltimulBackup) {
    $DataUltimuluiBackup = $UltimulBackup.LastWriteTime
    $DataUltimuluiBackupText = $DataUltimuluiBackup.ToString('dd.MM.yyyy HH:mm')
    $ZileTrecute = (New-TimeSpan -Start $DataUltimuluiBackup -End (Get-Date)).Days

    # Dacă ultimul backup este mai recent de 10 zile, SCRIPTUL SE OPREȘTE AICI (nu afișează și nu trimite nimic)
    if ($ZileTrecute -lt $PragZileAvertizare) {
        Write-Log "INFO: Ultimul backup este recent ($ZileTrecute zile - $DataUltimuluiBackupText). Nicio acțiune necesară."
        exit 0
    }
} else {
    Write-Log "ATENȚIE: Nu s-a găsit niciun subdirector de backup în $CaleCopiaDeSiguranta."
    $ZileTrecute = "nedeterminată (folder gol)"
    $DataUltimuluiBackupText = "Niciun backup găsit"
}

Write-Log "ATENȚIE: Backup învechit detectat ($ZileTrecute zile vechime)."

# ==============================================================================
# 2. CONFIGURARE & TRIMITERE E-MAIL (SMTP)
# ==============================================================================
$SmtpServer            = "nume.domeniu.ro"
$SmtpPort              = 587
$Expeditor             = '"DASIL - Automatizari IT" <domeniu@dasil.ro>'
$Destinatar            = "test1@dasil.ro, test2@dasil.ro"
$NecesitaAutentificare = $true
$SmtpUser              = "numey@dasil.ro"
$SmtpPass              = "xxx"
$UseSSL                = $true

function Send-ReportEmail {
    param (
        [string]$Subiect, 
        [string]$MesajStare
    )
    try {
        $TextLogCurent = $Script:LogCurent -join "`r`n"
        $CorpEmail = @"
=== INFORMARE AUTOMATA BACKUP - ADICOM / CANTINA ===
Data/Ora Verificare: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Status: ATENTIE - BACKUP INVECHIT

Detalii:
$MesajStare

==================== LOG EXECUTIE ====================
$TextLogCurent

-- 
Sistem Automatizat de Backup & Disaster Recovery
Inspector Informatică Sorin Marian
"@
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $MailMessage = New-Object System.Net.Mail.MailMessage
        $MailMessage.From = $Expeditor
        $MailMessage.Subject = $Subiect
        $MailMessage.Body = $CorpEmail
        $MailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

        foreach ($Adresa in $Destinatar.Split(',')) {
            if ($Adresa.Trim() -ne "") { $MailMessage.To.Add($Adresa.Trim()) }
        }

        $SmtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $SmtpClient.EnableSsl = $UseSSL
        if ($NecesitaAutentificare) { 
            $SmtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPass) 
        }
        $SmtpClient.Send($MailMessage)
        Write-Log "E-mail de alertă trimis cu succes."
    } catch {
        Write-Log "EROARE SMTP: $_"
    }
}

# Trimitere e-mail de avertizare
$DataFormatata = Get-Date -Format 'dd.MM.yyyy HH:mm'
$MesajEmailDetails = "Ultimul backup efectuat în aplicația Adi-Com datează din: $DataUltimuluiBackupText ($ZileTrecute zile în urmă).`r`nA fost afișată notificarea vizuală pe ecranul utilizatorului."
Send-ReportEmail -Subiect "[ATENTIE] Backup Invechit Adi-Com ($ZileTrecute zile) - $DataFormatata" -MesajStare $MesajEmailDetails

# ==============================================================================
# 3. CONSTRUIRE & AFIȘARE FEREASTRĂ DARK MODE (TEXT DINAMIC)
# ==============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Activare titlu/bordură Dark Mode la nivel de OS (Windows 10/11)
$DwmApiCode = @"
using System;
using System.Runtime.InteropServices;

public class Dwm {
    [DllImport("dwmapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@
Add-Type -TypeDefinition $DwmApiCode -ErrorAction SilentlyContinue

# Configurare Fereastră Alertă
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Notificare Backup Adi-Com - DASIL"
$Form.Size = New-Object System.Drawing.Size(420, 190)
$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual

# Poziționare în colțul din dreapta jos (deasupra Taskbar-ului)
$Screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$Form.Location = New-Object System.Drawing.Point(($Screen.Right - $Form.Width - 10), ($Screen.Bottom - $Form.Height - 10))

$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form.MaximizeBox = $false
$Form.MinimizeBox = $false
$Form.TopMost = $true # Rămâne peste toate celelalte ferestre

# Culori dedicate Dark Mode
$Form.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
$Form.ForeColor = [System.Drawing.Color]::White

# Aplicare titlu întunecat pentru Windows 10/11
$Form.add_HandleCreated({
    $useDarkMode = 1
    [Dwm]::DwmSetWindowAttribute($Form.Handle, 20, [ref]$useDarkMode, 4)
})

# Iconiță de Avertizare
$IconLabel = New-Object System.Windows.Forms.Label
$IconLabel.Location = New-Object System.Drawing.Point(15, 20)
$IconLabel.Size = New-Object System.Drawing.Size(45, 45)
$IconLabel.Image = [System.Drawing.SystemIcons]::Warning.ToBitmap()
$Form.Controls.Add($IconLabel)

# Titlu Alertă
$TitleLabel = New-Object System.Windows.Forms.Label
$TitleLabel.Location = New-Object System.Drawing.Point(65, 15)
$TitleLabel.Size = New-Object System.Drawing.Size(330, 25)
$TitleLabel.Text = "ATENȚIE: BACKUP INVECHIT!"
$TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$TitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 193, 7) # Galben avertizare
$Form.Controls.Add($TitleLabel)

# Mesaj Text cu preluare dinamică a numărului de zile
$MessageLabel = New-Object System.Windows.Forms.Label
$MessageLabel.Location = New-Object System.Drawing.Point(65, 45)
$MessageLabel.Size = New-Object System.Drawing.Size(330, 55)
$MessageLabel.Text = "Nu ați mai efectuat o copie de siguranță de $ZileTrecute zile.`r`nVă rugăm să generați o copie nouă din aplicație!"
$MessageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$MessageLabel.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$Form.Controls.Add($MessageLabel)

# Buton de Închidere "Am înțeles"
$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Location = New-Object System.Drawing.Point(270, 110)
$CloseButton.Size = New-Object System.Drawing.Size(125, 30)
$CloseButton.Text = "Am înțeles"
$CloseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$CloseButton.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$CloseButton.ForeColor = [System.Drawing.Color]::White
$CloseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$CloseButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$CloseButton.Add_Click({ $Form.Close() })
$Form.Controls.Add($CloseButton)

# Timer pentru închidere automată după 10 minute (600.000 ms)
$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 600000 # 10 minute
$Timer.Add_Tick({
    $Timer.Stop()
    $Form.Close()
})
$Timer.Start()

# Afișare fereastră
[void]$Form.ShowDialog()

# Curățare resurse din memorie la închidere
$Timer.Dispose()
$Form.Dispose()