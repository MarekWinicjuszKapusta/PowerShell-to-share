### Functions

# Function to generate a timestamp that is added to the log file
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

# Function to generate a log file
$logsDir = "C:\Support\logs\Apps"
if (!(Test-Path -Path $logsDir -PathType Container)) {
    New-Item -ItemType Directory -Path $logsDir | Out-Null
}
$LogFile = "$logsDir\uniflow.log"

Function LogWrite {
    Param ([string]$logstring)
    Add-Content $Logfile -Value "$(Get-TimeStamp) $logstring"
    Write-Host $logstring
}

function Test-ApplicationInstalled {
    param (
        [string]$AppName
    )     
    # Query installed applications using Get-WmiObject and filter by name     
    $installedApps = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $AppName }     
    # Return True if the application is found; otherwise, return False     
    return [bool]($installedApps -ne $null) 
} 
# Usage example: 
#$isInstalled = Test-ApplicationInstalled -AppName "Uniflow SmartClient" 
#Write-Host "Is 'Uniflow SmartClient' installed? $isInstalled"

### Uniflow installation start

Logwrite "Starting Uniflow installation process"

$isInstalled = Test-ApplicationInstalled -AppName "Uniflow SmartClient"
Logwrite "Is uniflow installed? $isInstalled"

If($isInstalled -eq $true){
    Logwrite "Uniflow is already installed"
}
Else{
    Logwrite "Installing Uniflow"
    Set-Location C:\Support\Uniflow
    .\NTware_uniFLOWSmartClient_22.2.1_x64_1.ps1 -deploymenttype "install"
    $isInstalled = Test-ApplicationInstalled -AppName "Uniflow SmartClient"
    Logwrite "Is uniflow installed? $isInstalled"
}
Logwrite "Unregistering UniflowInstallation task"
Unregister-ScheduledTask -TaskName "UniflowInstallation"

Logwrite "End of Uniflow installation process /n"
