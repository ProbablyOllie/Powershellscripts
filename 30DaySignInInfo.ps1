# Defines the output path for the CSV
$CsvPath = "C:\temp\ActiveDevicesReport.csv"

# Connects to Microsoft Graph
# We need Device permissions to see the devices and User permissions to resolve the owner's name
Connect-MgGraph -Scopes "Device.Read.All", "User.Read.All"

# Calculate the date 30 days ago in ISO 8601 format (required for Graph filtering)
$DateThreshold = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "Fetching devices active since $DateThreshold..." -ForegroundColor Cyan

# Retrieve devices active in the last 30 days
# We use -All to handle pagination automatically
$Devices = Get-MgDevice -Filter "approximateLastSignInDateTime ge $DateThreshold" -All

$Report = New-Object System.Collections.Generic.List[PSCustomObject]

foreach ($Device in $Devices) {
    # Initialize variables
    $OwnerUpn = "No Owner Assigned"
    $OwnerDisplayName = "N/A"
    
    # Retrieve the Registered Owner of the device
    # Note: A device can have registered owners (users) or be ownerless (e.g., hybrid joined sometimes)
    $Owner = Get-MgDeviceRegisteredOwner -DeviceId $Device.Id -ErrorAction SilentlyContinue
    
    # Check if we got an owner object back
    if ($Owner) {
        # The returned object is a DirectoryObject; we look at AdditionalProperties for user details
        # Some objects might be Service Principals, but usually they are Users
        if ($Owner.AdditionalProperties.ContainsKey('userPrincipalName')) {
            $OwnerUpn = $Owner.AdditionalProperties['userPrincipalName']
            $OwnerDisplayName = $Owner.AdditionalProperties['displayName']
        }
        elseif ($Owner.AdditionalProperties.ContainsKey('displayName')) {
             $OwnerDisplayName = $Owner.AdditionalProperties['displayName']
        }
    }

    # Add to report list
    $Report.Add([PSCustomObject]@{
        DeviceName        = $Device.DisplayName
        LastSignIn        = $Device.ApproximateLastSignInDateTime
        OS                = $Device.OperatingSystem
        OSVersion         = $Device.OperatingSystemVersion
        TrustType         = $Device.TrustType
        UserPrincipalName = $OwnerUpn
        UserDisplayName   = $OwnerDisplayName
        DeviceId          = $Device.Id
    })
}

# Export to CSV
$Report | Export-Csv -Path $CsvPath -NoTypeInformation

Write-Host "Export complete! File saved to: $CsvPath" -ForegroundColor Green