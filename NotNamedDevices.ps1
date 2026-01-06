
<##created by Ollie Leggett

This script retreives devices that are not part of the naming convention- please see ITG documentation for more details on how to alter this script for your needs##>
$Devices = Get-MgDeviceManagementManagedDevice | Where-Object {$_.DeviceName -notlike "<#ENTER SNIPPIT OF NAMING CONVENTION HERE#>" }
ForEach ($Devices in $Devices) {

    $NewName = "ENTER NAMING CONVENTION PREFIX HERE-($Device.SerialNumber)"

    Update-MgDeviceManagementManagedDevice -ManagedDeviceId $Device.Id -DeviceName $NewName
}