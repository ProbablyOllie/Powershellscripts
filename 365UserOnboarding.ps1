<#
.SYNOPSIS
    GUI Tool to create M365 Users with Country Dropdown.
.DESCRIPTION
    Uses Microsoft Graph PowerShell SDK and Windows Forms.
#>

# 1. Load Windows Forms Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 2. Connect to Microsoft Graph
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Group.ReadWrite.All" -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Could not connect to Microsoft Graph. Please check internet and permissions.`nError: $_", "Connection Error", "OK", "Error")
    Exit
}

# 3. Fetch Data for Dropdowns
try {
    $domains = Get-MgDomain | Where-Object { $_.IsVerified -eq $true } | Select-Object -ExpandProperty Id
    $licenses = Get-MgSubscribedSku | Where-Object { $_.CapabilityStatus -eq "Enabled" }
    $groups = Get-MgGroup -Top 200 | Where-Object { $_.GroupTypes -contains "Unified" -or $_.SecurityEnabled -eq $true } | Sort-Object DisplayName
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error fetching M365 data.`nError: $_", "Data Error", "OK", "Error")
    Exit
}

# --- DEFINE COUNTRY LIST (ISO Codes) ---
# M365 requires 2-letter codes (e.g. GB, US). We map names to codes here.
$countryList = [ordered]@{
    "United Kingdom" = "GB"
    "United States"  = "US"
    "Canada"         = "CA"
    "Australia"      = "AU"
    "Ireland"        = "IE"
    "New Zealand"    = "NZ"
    "Germany"        = "DE"
    "France"         = "FR"
    "Spain"          = "ES"
    "Italy"          = "IT"
    "Netherlands"    = "NL"
    "Belgium"        = "BE"
    "Switzerland"    = "CH"
    "India"          = "IN"
    "South Africa"   = "ZA"
    "Japan"          = "JP"
    "China"          = "CN"
}

# 4. Create the GUI Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "M365 User Creator"
$form.Size = New-Object System.Drawing.Size(500, 680) # Increased height slightly
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

function Create-Label ($text, $top) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point(20, $top)
    $label.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($label)
}

# --- Form Inputs ---

# First Name
Create-Label "First Name:" 20
$txtFirstName = New-Object System.Windows.Forms.TextBox
$txtFirstName.Location = New-Object System.Drawing.Point(180, 20)
$txtFirstName.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($txtFirstName)

# Last Name
Create-Label "Last Name:" 50
$txtLastName = New-Object System.Windows.Forms.TextBox
$txtLastName.Location = New-Object System.Drawing.Point(180, 50)
$txtLastName.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($txtLastName)

# Username (Prefix)
Create-Label "Username (Prefix):" 80
$txtUserPrefix = New-Object System.Windows.Forms.TextBox
$txtUserPrefix.Location = New-Object System.Drawing.Point(180, 80)
$txtUserPrefix.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($txtUserPrefix)

# Domain Dropdown
$cmbDomain = New-Object System.Windows.Forms.ComboBox
$cmbDomain.Location = New-Object System.Drawing.Point(310, 80)
$cmbDomain.Size = New-Object System.Drawing.Size(120, 20)
$cmbDomain.DropDownStyle = "DropDownList" # Prevent typing
foreach ($d in $domains) { $cmbDomain.Items.Add($d) }
if ($cmbDomain.Items.Count -gt 0) { $cmbDomain.SelectedIndex = 0 }
$form.Controls.Add($cmbDomain)

# Password
Create-Label "Password:" 110
$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(180, 110)
$txtPassword.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($txtPassword)

# Location Dropdown (UPDATED)
Create-Label "Usage Location:" 140
$cmbLocation = New-Object System.Windows.Forms.ComboBox
$cmbLocation.Location = New-Object System.Drawing.Point(180, 140)
$cmbLocation.Size = New-Object System.Drawing.Size(250, 20)
$cmbLocation.DropDownStyle = "DropDownList" # Prevent typing invalid names

# Populate Location Dropdown
foreach ($countryName in $countryList.Keys) {
    $cmbLocation.Items.Add($countryName)
}
# Default to United Kingdom (Index 0 based on our ordered list)
$cmbLocation.SelectedIndex = 0 
$form.Controls.Add($cmbLocation)

# Licenses (CheckListBox)
Create-Label "Assign Licenses:" 170
$chkLicenses = New-Object System.Windows.Forms.CheckedListBox
$chkLicenses.Location = New-Object System.Drawing.Point(180, 170)
$chkLicenses.Size = New-Object System.Drawing.Size(250, 80)
foreach ($sku in $licenses) { 
    $chkLicenses.Items.Add($sku.SkuPartNumber) 
}
$form.Controls.Add($chkLicenses)

# Groups (CheckListBox)
Create-Label "Add to Groups:" 260
$chkGroups = New-Object System.Windows.Forms.CheckedListBox
$chkGroups.Location = New-Object System.Drawing.Point(180, 260)
$chkGroups.Size = New-Object System.Drawing.Size(250, 150)
foreach ($grp in $groups) { 
    $chkGroups.Items.Add($grp.DisplayName) 
}
$form.Controls.Add($chkGroups)

# Create Button
$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = "CREATE USER"
$btnCreate.Location = New-Object System.Drawing.Point(180, 430)
$btnCreate.Size = New-Object System.Drawing.Size(250, 40)
$btnCreate.BackColor = "LightGreen"
$form.Controls.Add($btnCreate)

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 490)
$lblStatus.Size = New-Object System.Drawing.Size(440, 60)
$lblStatus.Text = "Ready..."
$form.Controls.Add($lblStatus)

# --- Logic: Create Button Click Event ---
$btnCreate.Add_Click({
    $lblStatus.Text = "Processing..."
    $form.Refresh()

    # 1. Validation
    if ($txtFirstName.Text -eq "" -or $txtLastName.Text -eq "" -or $txtUserPrefix.Text -eq "" -or $txtPassword.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Please fill in all Name, Username and Password fields.", "Missing Info", "OK", "Warning")
        return
    }

    $UPN = "$($txtUserPrefix.Text)@$($cmbDomain.SelectedItem)"
    
    # 2. Get Correct ISO Code from Dropdown
    $selectedCountryName = $cmbLocation.SelectedItem
    $isoCode = $countryList[$selectedCountryName]

    # 3. Build Password Profile
    $PasswordProfile = @{
        Password = $txtPassword.Text
        ForceChangePasswordNextSignIn = $true
    }

    try {
        # 4. Create User
        $newUserParams = @{
            DisplayName = "$($txtFirstName.Text) $($txtLastName.Text)"
            GivenName = $txtFirstName.Text
            Surname = $txtLastName.Text
            UserPrincipalName = $UPN
            MailNickname = $txtUserPrefix.Text
            UsageLocation = $isoCode   # Passes "GB" not "United Kingdom"
            PasswordProfile = $PasswordProfile
            AccountEnabled = $true
        }

        $newUser = New-MgUser @newUserParams -ErrorAction Stop
        $log = "User '$UPN' created in location $isoCode.`n"

        # 5. Assign Licenses
        if ($chkLicenses.CheckedItems.Count -gt 0) {
            $addLicenses = @()
            foreach ($selectedSkuName in $chkLicenses.CheckedItems) {
                $skuObj = $licenses | Where-Object { $_.SkuPartNumber -eq $selectedSkuName }
                if ($skuObj) {
                    $addLicenses += @{ SkuId = $skuObj.SkuId }
                }
            }
            
            if ($addLicenses.Count -gt 0) {
                Set-MgUserLicense -UserId $newUser.Id -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
                $log += "Licenses assigned.`n"
            }
        }

        # 6. Add to Groups
        if ($chkGroups.CheckedItems.Count -gt 0) {
            foreach ($selectedGroupName in $chkGroups.CheckedItems) {
                $grpObj = $groups | Where-Object { $_.DisplayName -eq $selectedGroupName }
                if ($grpObj) {
                    New-MgGroupMember -GroupId $grpObj.Id -DirectoryObjectId $newUser.Id -ErrorAction Stop
                }
            }
            $log += "Added to $($chkGroups.CheckedItems.Count) groups."
        }

        $lblStatus.Text = $log
        [System.Windows.Forms.MessageBox]::Show($log, "Success", "OK", "Information")
        
        # Optional: Reset form fields here if desired

    }
    catch {
        $errorMessage = $_.Exception.Message
        $lblStatus.Text = "Error encountered."
        [System.Windows.Forms.MessageBox]::Show("Failed to create user.`n$errorMessage", "Error", "OK", "Error")
    }
})

# 5. Show the Form
$form.ShowDialog() | Out-Null