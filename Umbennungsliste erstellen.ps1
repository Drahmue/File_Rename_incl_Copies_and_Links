# --- Konfiguration ---
$searchPath  = "S:\Korrespondenz\Post Archiv"
$searchTerm  = "Booking "
$replaceTerm = "Booking_com "

# Dynamischer Pfad für die Ergebnisdatei (im Ordner des Skripts)
$timestamp   = Get-Date -Format "yyyy-MM-dd_HHmm"
$reportPath  = Join-Path -Path $PSScriptRoot -ChildPath "Dateiliste_$timestamp.txt"

# --- Logik ---

# Prüfen, ob der Suchpfad existiert
if (-not (Test-Path -Path $searchPath)) {
    Write-Host "FEHLER: Der Pfad '$searchPath' wurde nicht gefunden." -ForegroundColor Red
    return
}

# Filter für Get-ChildItem
$filter = "$searchTerm*"

# Dateien abrufen und verarbeiten
$results = Get-ChildItem -Path $searchPath -File -Filter $filter | ForEach-Object {
    $originalName = $_.Name
    $extension    = $_.Extension
    
    # Ersetzung im Basisnamen (ohne Endung)
    $newNameBase = $_.BaseName -replace "^$searchTerm", $replaceTerm
    
    # Zusammensetzen: Neuer Name + ursprüngliche Endung
    $newNameWithExt = $newNameBase + $extension
    
    # Format: Originalname mit Endung;Neuer Name mit Endung
    "$originalName;$newNameWithExt"
}

# Ergebnisse speichern
if ($results) {
    $results | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "Erfolg! Die Liste wurde hier erstellt: $reportPath" -ForegroundColor Green
    Write-Host "Verarbeitete Dateien: $($results.Count)" -ForegroundColor Cyan
} else {
    Write-Host "Keine Dateien gefunden, die mit '$searchTerm' beginnen." -ForegroundColor Yellow
}