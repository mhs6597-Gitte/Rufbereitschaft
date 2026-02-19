# Skript zum Erstellen einer Kopie der Rufbereitschafts-Vorlage mit Jahres- und Monatsangabe
# Die Vorlage sollte den Namen Rufbereitschaft_TWS_YYYY_MM.xlsx haben
# Das Skript kopiert die Vorlage, benennt sie entsprechend um und trägt das Datum des Monatsersten in die Zelle A2 der Reiter "RB-Einsätze" und "Rufbereitschaft" ein.

param(
    [int]$Month
)

# Definiere den Speicherort
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$templateName = "Rufbereitschaft_TWS_YYYY_MM.xlsx"
$templatePath = Join-Path $scriptDir $templateName

# Prüfe, ob die Vorlage existiert
if (-not (Test-Path $templatePath)) {
    Write-Host "Fehler: Vorlage '$templateName' nicht gefunden in '$scriptDir'" -ForegroundColor Red
    exit 1
}

# Wenn Monat nicht als Parameter übergeben wurde, frage den Benutzer
if ($Month -eq 0) {
    do {
        $monthInput = Read-Host "Bitte geben Sie den Monat ein (1-12)"
        [int]$Month = $monthInput
        
        if ($Month -lt 1 -or $Month -gt 12) {
            Write-Host "Ungültige Eingabe! Bitte geben Sie einen Wert zwischen 1 und 12 ein." -ForegroundColor Yellow
            $Month = 0
        }
    } while ($Month -eq 0 -or [string]::IsNullOrWhiteSpace($monthInput))
}
else {
    # Validiere den Parameter
    if ($Month -lt 1 -or $Month -gt 12) {
        Write-Host "Fehler: Monat muss zwischen 1 und 12 liegen." -ForegroundColor Red
        exit 1
    }
}

# Hole das aktuelle Jahr
$year = (Get-Date).Year

# ============================================================================
# Funktion zum Füllen der Rufbereitschaftsliste
# ============================================================================
function Set-RufbereitschaftListe {
    param(
        [object]$Worksheet,
        [int]$Year,
        [int]$Month
    )
    
    # Definiere die Startreihe für Dateneintrag
    $dataStartRow = 5
    $currentWriteRow = $dataStartRow
    
    # Berechne den ersten und letzten Tag des Monats
    $firstDayOfMonth = [datetime]::new($Year, $Month, 1)
    $lastDayOfMonth = $firstDayOfMonth.AddMonths(1).AddDays(-1)
    
    # Ermittle den Wochentag des ersten Tages (0=Sonntag, 1=Montag, ..., 6=Samstag)
    $firstDayOfWeek = [int]$firstDayOfMonth.DayOfWeek
    
    # Berechne das Startdatum der Woche (Montag, auch wenn vor dem 1. des Monats)
    if ($firstDayOfWeek -eq 0) {
        # Sonntag - gehe 6 Tage zurück zum Montag
        $weekStartDate = $firstDayOfMonth.AddDays(-6)
    }
    else {
        # Beliebiger Wochentag - gehe zum Montag der Woche (1-$firstDayOfWeek)
        $weekStartDate = $firstDayOfMonth.AddDays(1 - $firstDayOfWeek)
    }
    
    # Definiere die Muster-Zeilen (Mo=5, Di=6, Mi=7, Do=8, Fr=9, Sa=10, So=11)
    $patternRows = @{
        1 = 5  # Montag
        2 = 6  # Dienstag
        3 = 7  # Mittwoch
        4 = 8  # Donnerstag
        5 = 9  # Freitag
        6 = 10 # Samstag
        0 = 11 # Sonntag
    }
    
    # Durchlaufe alle Tage ab dem Montag der Woche
    $currentDate = $weekStartDate
    while ($currentDate -le $lastDayOfMonth) {
        # Bestimme den Wochentag
        $dayOfWeek = [int]$currentDate.DayOfWeek
        $sourceRow = $patternRows[$dayOfWeek]
        
        # Kopiere die komplette Zeile (mit Formatierung und Inhalten)
        $Worksheet.Rows($sourceRow).Copy() | Out-Null
        $Worksheet.Rows($currentWriteRow).PasteSpecial() | Out-Null
        
        # Überschreibe das Datum in Spalte D (4. Spalte)
        $dateString = $currentDate.ToString("dd.MM.yyyy")
        $Worksheet.Cells($currentWriteRow, 4).Value = $dateString
        
        $currentWriteRow++
        $currentDate = $currentDate.AddDays(1)
    }
    
    # Lösche die Muster-Zeilen (5-11), falls gewünscht
    $Worksheet.Rows("5:11").Delete() | Out-Null
}

# Formatiere den Monat mit führender Null
$monthFormatted = $Month.ToString("00")

# Erstelle den neuen Dateinamen
$newFileName = "Rufbereitschaft_TWS_${year}_${monthFormatted}.xlsx"
$newFilePath = Join-Path $scriptDir $newFileName

# Falls die Zieldatei bereits existiert, füge eine Nummerierung hinzu
$counter = 1
$originalNewFilePath = $newFilePath
while (Test-Path $newFilePath) {
    $counterFormatted = $counter.ToString("00")
    $baseFileName = "Rufbereitschaft_TWS_${year}_${monthFormatted}-${counterFormatted}.xlsx"
    $newFilePath = Join-Path $scriptDir $baseFileName
    $counter++
}

if ($newFilePath -ne $originalNewFilePath) {
    $newFileName = Split-Path -Leaf $newFilePath
    Write-Host "Datei mit diesem Namen existiert bereits. Verwende: '$newFileName'" -ForegroundColor Yellow
}

try {
    # Kopiere die Vorlage
    Copy-Item -Path $templatePath -Destination $newFilePath -Force
    Write-Host "Datei erfolgreich erstellt: '$newFileName'" -ForegroundColor Green
    
    # Öffne Excel und trage das Datum des Monatsersten ein
    $dateOfFirstDay = [datetime]::new($year, $Month, 1)
    $dateString = $dateOfFirstDay.ToString("dd.MM.yyyy")
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($newFilePath)
        
        # Versuche den Reiter "RB-Einsätze" zu finden
        $worksheet = $null
        foreach ($sheet in $workbook.Sheets) {
            if ($sheet.Name -eq "RB-Einsätze") {
                $worksheet = $sheet
                break
            }
        }
        
        if ($null -eq $worksheet) {
            Write-Host "Warnung: Reiter 'RB-Einsätze' nicht gefunden." -ForegroundColor Yellow
        }
        else {
            # Setze den Wert in Zelle A2
            $cell = $worksheet.Range("A2")
            $cell.Value = $dateString
            Write-Host "Datum '$dateString' in Zelle A2 des Reiters 'RB-Einsätze' eingetragen." -ForegroundColor Green
        }
        
        # Versuche den Reiter "Rufbereitschaft" zu finden
        $worksheet = $null
        foreach ($sheet in $workbook.Sheets) {
            if ($sheet.Name -eq "Rufbereitschaft") {
                $worksheet = $sheet
                break
            }
        }
        
        if ($null -eq $worksheet) {
            Write-Host "Warnung: Reiter 'Rufbereitschaft' nicht gefunden." -ForegroundColor Yellow
        }
        else {
            # Setze den Wert in Zelle A2
            $cell = $worksheet.Range("A2")
            $cell.Value = $dateString
            Write-Host "Datum '$dateString' in Zelle A2 des Reiters 'Rufbereitschaft' eingetragen." -ForegroundColor Green
            
            # Fülle die Rufbereitschaftsliste
            Set-RufbereitschaftListe -Worksheet $worksheet -Year $year -Month $Month
            Write-Host "Rufbereitschaftsliste für $month/$year erstellt." -ForegroundColor Green
        }
        
        # Speichere und schließe die Datei
        $workbook.Save()
        $workbook.Close()
        $excel.Quit()
        
        # Gib COM-Objekte frei
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    catch {
        Write-Host "Warnung: Excel-Operationen nicht möglich: $_" -ForegroundColor Yellow
        Write-Host "Die Datei wurde erstellt, aber das Datum konnte nicht automatisch eingetragen werden." -ForegroundColor Yellow
    }
    
    Write-Host "Pfad: $newFilePath" -ForegroundColor Cyan
}
catch {
    Write-Host "Fehler beim Kopieren der Datei: $_" -ForegroundColor Red
    exit 1
}
