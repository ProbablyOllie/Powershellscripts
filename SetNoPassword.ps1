$UserName = Read-Host "Enter the username"

# Set the password to an empty string
$User = [adsi]"WinNT://$env:COMPUTERNAME/$UserName,user"
$User.SetPassword("")
$User.SetInfo()

# Set the password to never expire
Set-LocalUser -Name $UserName -PasswordNeverExpires $true