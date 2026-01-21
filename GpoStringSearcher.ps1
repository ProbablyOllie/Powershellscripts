# Import the Group Policy module (required if not already loaded)
Import-Module GroupPolicy

# 1. Get the search term from the user
$SearchTerm = Read-Host "Enter the string you want to find (e.g., '10.0.0.1' or 'LoginScript.vbs')"

Write-Host "Searching all GPOs for '$SearchTerm'. This may take a minute..." -ForegroundColor Cyan

# 2. Get all GPOs in the domain
$AllGpos = Get-GPO -All

$Matches = foreach ($Gpo in $AllGpos) {
    # Generate an XML report for the GPO
    [xml]$GpoXml = Get-GPOReport -Guid $Gpo.Id -ReportType Xml
    
    # Convert the XML to a string to perform the search
    $GpoString = $GpoXml.OuterXml
    
    if ($GpoString -like "*$SearchTerm*") {
        Write-Host "[FOUND] Match in: $($Gpo.DisplayName)" -ForegroundColor Green
        [PSCustomObject]@{
            GPOName      = $Gpo.DisplayName
            CreationTime = $Gpo.CreationTime
            Modification = $Gpo.ModificationTime
            Status       = $Gpo.GpoStatus
        }
    }
}

# 3. Output the final list to a table
if ($Matches) {
    Write-Host "`nSearch Complete. Summary of matches:" -ForegroundColor Cyan
    $Matches | Format-Table -AutoSize
}
else {
    Write-Host "No GPOs found containing '$SearchTerm'." -ForegroundColor Red
}