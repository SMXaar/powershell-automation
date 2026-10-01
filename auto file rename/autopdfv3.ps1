# ==========================================================
# CONFIGURARE
# ==========================================================

$PdfFolder  = "D:\scripturi ps1 backup\other\autodevis\test"
$ExcelPath  = "D:\scripturi ps1 backup\other\autodevis\test\anaf.xlsx"
$CnpColumn  = "CNP"
$NameColumn = "Nume"


# ==========================================================
# 1. ASIGURARE MODUL IMPORTEXCEL
# ==========================================================

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[INFO] Instalez modulul ImportExcel..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel


# ==========================================================
# 2. VERIFICARE FOLDER + EXCEL
# ==========================================================

if (-not (Test-Path -Path $PdfFolder)) {
    Write-Host "[EROARE] Folderul nu există: $PdfFolder" -ForegroundColor Red
    Exit
}

if (-not (Test-Path -Path $ExcelPath)) {
    Write-Host "[EROARE CRITICĂ] Nu există fișierul Excel: $ExcelPath" -ForegroundColor Red
    Exit
}


# ==========================================================
# 3. CITIRE EXCEL
# ==========================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Încarc datele din fișierul Excel..." -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Cyan

$ExcelData = Import-Excel -Path $ExcelPath
$CNPTable = @{}

foreach ($row in $ExcelData) {
    if ($row.$CnpColumn) {
        $cnpVal = $row.$CnpColumn.ToString().Trim()
        $numeVal = ""

        if ($row.$NameColumn) {
            $numeVal = $row.$NameColumn.ToString().Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($cnpVal) -and -not [string]::IsNullOrWhiteSpace($numeVal)) {
            $CNPTable[$cnpVal] = $numeVal
        }
    }
}

Write-Host "[SUCCES] S-au încărcat $($CNPTable.Count) înregistrări din Excel." -ForegroundColor Green
Write-Host ""


# ==========================================================
# 4. PRELUARE FIȘIERE PDF / WORD
# ==========================================================

$Files = @(
    Get-ChildItem -Path $PdfFolder -File | Where-Object {
        $_.Extension -in ".pdf", ".docx", ".doc"
    }
)

if ($Files.Count -eq 0) {
    Write-Host "[INFO] Nu există PDF-uri sau documente Word în folder." -ForegroundColor Yellow
    Exit
}

Write-Host "Am găsit $($Files.Count) fișiere PDF/Word." -ForegroundColor Cyan
Write-Host ""


# ==========================================================
# 5. PROCESARE FIȘIERE
# ==========================================================

$word = $null
$startTime = Get-Date

$processed = 0
$success   = 0
$warnings  = 0
$errors    = 0

try {
    foreach ($file in $Files) {
        $processed++
        $doc = $null
        $documentText = $null
        $cnp = $null
        $isCnpInFilename = $false

        Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "[$processed/$($Files.Count)] $($file.Name)" -ForegroundColor White

        try {
            # --------------------------------------------------
            # PAS A: Caută CNP în NUMELE FIȘIERULUI
            # --------------------------------------------------
            $matchFilename = [regex]::Match($file.BaseName, '(?<!\d)[1-8]\d{12}(?!\d)')

            if ($matchFilename.Success) {
                $cnp = $matchFilename.Value
                $isCnpInFilename = $true
                Write-Host "  [INFO] CNP găsit în denumirea fișierului: $cnp" -ForegroundColor Yellow
            }
            else {
                # --------------------------------------------------
                # PAS B: Citire conținut fișier via Microsoft Word COM
                # --------------------------------------------------
                if ($word -eq $null) {
                    Write-Host "[INFO] Pornesc Microsoft Word..." -ForegroundColor Gray
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0
                    Write-Host "[SUCCES] Word pornit." -ForegroundColor Green
                    Write-Host ""
                }

                $doc = $word.Documents.Open($file.FullName, $false, $true, $false)
                $documentText = $doc.Content.Text
                
                $doc.Close([ref]$false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                $doc = $null

                # Extragere TOATE CNP-urile posibile (13 cifre) din document
                $allMatches = [regex]::Matches($documentText, '(?<!\d)[1-8]\d{12}(?!\d)')

                # Căutăm primul CNP extras din document care EXISTĂ efectiv în tabelul Excel
                foreach ($m in $allMatches) {
                    if ($CNPTable.ContainsKey($m.Value)) {
                        $cnp = $m.Value
                        break # Am găsit CNP-ul titularului din Excel, oprim căutarea
                    }
                }

                if ($cnp) {
                    Write-Host "  CNP potrivit din conținut: $cnp" -ForegroundColor Yellow
                }
            }

            # --------------------------------------------------
            # PAS C: Verificăm dacă s-a identificat un CNP valid
            # --------------------------------------------------
            if (-not $cnp) {
                Write-Host "  [ATENȚIE] Niciun CNP din fișier nu există în Excel." -ForegroundColor Red
                $warnings++
                continue
            }

            # --------------------------------------------------
            # PAS D: Căutare CNP în Tabela Excel
            # --------------------------------------------------
            if (-not $CNPTable.ContainsKey($cnp)) {
                Write-Host "  [ATENȚIE] CNP-ul $cnp nu există în Excel." -ForegroundColor Red
                $warnings++
                continue
            }

            $numeCorespunzator = $CNPTable[$cnp]
            Write-Host "  Nume identificat: $numeCorespunzator" -ForegroundColor Cyan

            # --------------------------------------------------
            # PAS E: Pregătire Nume Nou (Adăugare '- taxe' dacă era CNP în denumire)
            # --------------------------------------------------
            $safeName = $numeCorespunzator -replace '[\\/:*?"<>|]', '_'

            if ($isCnpInFilename) {
                $newFileName = "$safeName - taxe$($file.Extension)"
            } else {
                $newFileName = "$safeName$($file.Extension)"
            }

            $newFilePath = Join-Path -Path $PdfFolder -ChildPath $newFileName

            # Skip dacă este deja redenumit
            if ($file.Name -eq $newFileName) {
                Write-Host "  [SKIP] Fișierul este deja redenumit corect." -ForegroundColor DarkYellow
                continue
            }

            # Rezolvare duplicate (_1, _2 etc.)
            if (Test-Path -Path $newFilePath) {
                $i = 1
                do {
                    if ($isCnpInFilename) {
                        $newFileName = "${safeName} - taxe_$i$($file.Extension)"
                    } else {
                        $newFileName = "${safeName}_$i$($file.Extension)"
                    }
                    $newFilePath = Join-Path -Path $PdfFolder -ChildPath $newFileName
                    $i++
                } while (Test-Path -Path $newFilePath)
            }

            # --------------------------------------------------
            # PAS F: Redenumire
            # --------------------------------------------------
            Rename-Item -Path $file.FullName -NewName $newFileName
            Write-Host "  [SUCCES] → $newFileName" -ForegroundColor Green
            $success++

        }
        catch {
            $errors++
            Write-Host "  [EROARE] $($_.Exception.Message)" -ForegroundColor Red
        }
        finally {
            if ($doc -ne $null) {
                try { $doc.Close([ref]$false) } catch {}
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
                $doc = $null
            }
        }
    }
}
finally {
    # Închidere instanță Word la final
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
        $word = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ==========================================================
# 6. RAPORT FINAL
# ==========================================================

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " PROCESARE FINALIZATĂ" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total fișiere      : $($Files.Count)"
Write-Host "Procesate          : $processed"
Write-Host "Redenumite         : $success" -ForegroundColor Green
Write-Host "Atenționări        : $warnings" -ForegroundColor Yellow
Write-Host "Erori              : $errors" -ForegroundColor Red
Write-Host ""
Write-Host "Durată totală      : $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan