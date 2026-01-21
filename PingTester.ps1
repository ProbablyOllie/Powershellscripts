Add-Type -AssemblyName Microsoft.VisualBasic

# 1. Create the Pop-up Input Box
$Title = "IP Connectivity Test"
$Prompt = "Enter IP addresses or Hostnames (separate multiple with commas):"
$DefaultValue = "8.8.8.8, 1.1.1.1"

$UserInput = [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $DefaultValue)

# 2. Process the input into a list (Array)
if (-not [string]::IsNullOrWhiteSpace($UserInput)) {
    # Split the input by commas and trim any accidental spaces
    $IPList = $UserInput.Split(',').Trim()
    
    $Results = foreach ($IP in $IPList) {
        Write-Host "Testing: $IP..." -ForegroundColor Cyan
        
        if (Test-Connection -ComputerName $IP -Count 1 -Quiet) {
            Write-Host "[SUCCESS] $IP is Online" -ForegroundColor Green
            [PSCustomObject]@{ IP = $IP; Status = "Online" }
        }
        else {
            Write-Host "[FAILED] $IP is Offline" -ForegroundColor Red
            [PSCustomObject]@{ IP = $IP; Status = "Offline" }
        }
    }

    # 3. Output the final summary to the console
    $Results | Format-Table -AutoSize
}
else {
    Write-Host "Operation cancelled or no input provided." -ForegroundColor Yellow
}