# Script for Clean reinstall of Microsoft Teams
# Created by Marek.Kapusta@fujitsu.com
# Last update: 23.01.2023 14:12

$ErrorActionPreference = 'SilentlyContinue'

# Function to generate a timestamp that is added to the log file
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}
    
# Function to generate a log file
$logsDir = "$ENV:SystemDrive\support\logs\apps"
if (!(Test-Path -Path $logsDir -PathType Container)) {
    New-Item -ItemType Directory -Path $logsDir | Out-Null
}
$LogFile = "$logsDir\CleanTeamsRemoval.log"

Function LogWrite {
    Param ([string]$logstring)
    Add-Content $Logfile -Value "$(Get-TimeStamp) $logstring"
}

function RemoveLeftOvers {
    [CmdletBinding()]
    Param(
    [String]$Path
    )
    if(Test-Path $Path){
        LogWrite "Found $path, clearing content of the folder"
        Write-Host "Found $path, clearing content of the folder"
        Get-ChildItem $path | Remove-Item -Recurse -Confirm:$false
        Remove-Item $Path -Recurse -Confirm:$false
        if(Test-Path $Path){
            LogWrite "Error: complications with deleting $path"
            Write-Host "Error: complications with deleting $path, will try again after 10 seconds"
            $attempt = 2
            while((Test-Path $Path) -and ($attempt -le 10)){
                Start-Sleep -Seconds 10
                LogWrite "$attempt attempt to delete $path"
                Write-Host "$attempt attempt to delete $path"
                Get-ChildItem $path | Remove-Item -Recurse -Confirm:$false
                Remove-Item $Path -Recurse -Confirm:$false
                $attempt++
            }
            if(Test-Path $Path){
                LogWrite "Error: unable to delete $path"
                Write-Host "Error: unable to delete $path"
            }
            else {
                LogWrite "$path successfully deleted"
                Write-Host "$path successfully deleted"
            }
        }
        else {
            LogWrite "$path successfully deleted"
            Write-Host "$path successfully deleted"
        }
    }
    else {
        LogWrite "Unable to find $path"
        Write-Host "Unable to find $path"
    }
}

#---------------------------------------------------------------------

LogWrite "** STARTING Teams Cache Removal Script **"
write-host "** STARTING Teams Cache Removal Script **"

# Close Outlook
Stop-Process -Name "Outlook" -ErrorAction SilentlyContinue

# Close Teams
Stop-Process -Name "Teams" -ErrorAction SilentlyContinue

removeLeftOvers -path "C:\Users\$env:UserName\AppData\Roaming\Microsoft\teams"
removeLeftOvers -path "C:\Users\$env:UserName\AppData\Roaming\teams"
removeLeftOvers -path "C:\Users\$env:UserName\AppData\local\Microsoft\teamsMeetingAddin"
removeLeftOvers -path "C:\Users\$env:UserName\AppData\local\Microsoft\teamsPresenceAddin"
removeLeftOvers -path "C:\Users\$env:UserName\AppData\local\Microsoft\SquirrelTemp"
    
LogWrite "** ENDING Teams Cache Removal Script **"
write-host "** ENDING Teams Cache Removal Script **"
