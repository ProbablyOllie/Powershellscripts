# 1. Force-terminate Teams and Outlook to release file locks
$ProcessesToKill = @("Teams", "ms-teams", "msteams", "Outlook")
Write-Host "Closing Teams and Outlook..." -ForegroundColor Yellow

foreach ($proc in $ProcessesToKill) {
    if (Get-Process $proc -ErrorAction SilentlyContinue) {
        Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
    }
}

# Wait for file handles to close
Start-Sleep -Seconds 5

# 2. Clear Primary Teams Cache (New Teams)
$TeamsPath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"
if (Test-Path $TeamsPath) {
    Write-Host "Nuking Teams Cache..." -ForegroundColor Cyan
    Get-ChildItem -Path $TeamsPath -Recurse | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

# 3. Clear Teams Meeting Add-in Cache for Outlook
# This targets the 'Cef' folder where the Add-in stores temporary browser data
$AddInCache = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAddin"
$WebView2Cache = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\EBWebView\CoreSmartScreen" # Common for modern add-ins

Write-Host "Nuking Meeting Add-in Cache..." -ForegroundColor Cyan

$AddInPaths = @(
    "$env:APPDATA\Microsoft\Teams\Service Worker\CacheStorage",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"  # Office Web Extension Cache
)

foreach ($path in $AddInPaths) {
    if (Test-Path $path) {
        Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "Cleanup Complete. Please restart Outlook and Teams." -ForegroundColor Green