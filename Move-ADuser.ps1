# Import the Active Directory module (if not already loaded)
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found. Please ensure it's installed and available." -ForegroundColor Red
    # You might want to exit here or provide instructions for installation
    exit
}
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Function to get a valid OU path from the user
function Get-ValidOUPath {
    param (
        [string]$PromptMessage = "Please enter the target OU path (e.g., 'OU=NewUsers,DC=contoso,DC=com'): "
    )
    $ouPath = Read-Host $PromptMessage
    while (-not (Get-ADOrganizationalUnit -Identity $ouPath -ErrorAction SilentlyContinue)) {
        Write-Host "Invalid OU path or OU does not exist. Please try again." -ForegroundColor Yellow
        $ouPath = Read-Host $PromptMessage
    }
    return $ouPath
}

# Prompt for the attribute name
$attributeName = Read-Host "Enter the Active Directory attribute name to search by (e.g., 'Department', 'City', 'Description'): "

# Prompt for the attribute value
$attributeValue = Read-Host "Enter the value for the attribute '$attributeName': "

# Prompt for the target OU path
$targetOU = Get-ValidOUPath

Write-Host "`nSearching for users with '$attributeName'='$attributeValue'..." -ForegroundColor Green

# Find users based on the provided attribute and value
try {
    $usersToMove = Get-ADUser -Filter "$attributeName -eq '$attributeValue'" -Properties $attributeName, DistinguishedName -ErrorAction Stop
}
catch {
    Write-Host "Error searching for users: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

if ($usersToMove) {
    Write-Host "Found $($usersToMove.Count) user(s) matching the criteria:" -ForegroundColor Green
    $usersToMove | Select-Object SamAccountName, $attributeName, DistinguishedName | Format-Table -AutoSize

    $confirm = Read-Host "`nDo you want to move these user(s) to '$targetOU'? (Y/N)"
    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        foreach ($user in $usersToMove) {
            Write-Host "Attempting to move user '$($user.SamAccountName)'..." -ForegroundColor Cyan
            try {
                Move-ADObject -Identity $user.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
                Write-Host "Successfully moved '$($user.SamAccountName)' to '$targetOU'." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to move '$($user.SamAccountName)': $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        Write-Host "`nMove operation completed." -ForegroundColor Green
    }
    else {
        Write-Host "Move operation cancelled." -ForegroundColor Yellow
    }
}
else {
    Write-Host "No users found with '$attributeName'='$attributeValue'." -ForegroundColor Yellow
}

Write-Host "`nScript finished."