param(
    [string]$projectPath = "BKB-Tool/BKB-Tool.csproj"
)

Write-Host "Incrementing version in $projectPath" -ForegroundColor Cyan

# Projektdatei einlesen
$csprojContent = Get-Content $projectPath -Raw

# Aktuelle Version extrahieren
if ($csprojContent -match '<Version>([\d\.]+)</Version>') {
    $currentVersion = $matches[1]
    Write-Host "Current version: $currentVersion" -ForegroundColor Yellow
    
    # Version in Teile zerlegen
    $versionParts = $currentVersion.Split('.')
    $major = [int]$versionParts[0]
    $minor = [int]$versionParts[1]
    $build = [int]$versionParts[2]
    
    # Build-Nummer erhöhen
    $build++
    
    $newVersion = "$major.$minor.$build"
    Write-Host "New version: $newVersion" -ForegroundColor Green
    
    # Alle Version-Tags aktualisieren
    $csprojContent = $csprojContent -replace '<Version>[\d\.]+</Version>', "<Version>$newVersion</Version>"
    $csprojContent = $csprojContent -replace '<AssemblyVersion>[\d\.]+</AssemblyVersion>', "<AssemblyVersion>$newVersion</AssemblyVersion>"
    $csprojContent = $csprojContent -replace '<FileVersion>[\d\.]+</FileVersion>', "<FileVersion>$newVersion</FileVersion>"
    
    # Datei speichern
    $csprojContent | Set-Content $projectPath -NoNewline
    
    Write-Host "Version successfully updated!" -ForegroundColor Green
    
    # Version für GitHub Actions ausgeben
    Write-Output "NEW_VERSION=$newVersion" >> $env:GITHUB_OUTPUT
    Write-Output "CURRENT_VERSION=$currentVersion" >> $env:GITHUB_OUTPUT
} else {
    Write-Host "Error: Could not find version tag in project file" -ForegroundColor Red
    exit 1
}