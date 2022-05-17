$ErrorActionPreference = 'SilentlyContinue'

#Function to generate a timestamp that is added to the log file
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)   
}
    
#Function to generate a log file
if ((Test-Path -Path "$ENV:SystemDrive\support\logs\apps" -PathType Container) -ne $true ) {mkdir "$ENV:SystemDrive\support\logs\apps" | Out-Null}
$LogFile = "$ENV:SystemDrive\support\logs\apps\CleanTeamsRemoval.log"

Function LogWrite{
    Param ([string]$logstring)
    Add-content $Logfile -value "$(Get-Timestamp) $logstring"
}

function removeLeftOvers{
    [CmdletBinding()]
    Param(
    [String]$Path
    )
    if(Test-Path $Path){
        LogWrite "found $path, clearing content of the folder"
        write-host "foud $path, clearing content of the folder"
        Get-ChildItem $path | Remove-Item -Recurse -Confirm:$false
        remove-item $Path -Recurse -Confirm:$false
        if(Test-Path $Path){
            LogWrite "error: complications with deleting $path"
            write-host "error: complications with deleting $path, will try again after 10 seconds"
            $attemp = 2
            while((Test-Path $Path)-and($attemp -le 10)){
            Start-Sleep -Seconds 10
            LogWrite "$attemp attempt to delete $path"
            write-host "$attemp attempt to delete $path"
            Get-ChildItem $path | Remove-Item -Recurse -Confirm:$false
            remove-item $Path -Recurse -Confirm:$false
            $attemp++
            }
            if(Test-Path $Path){
                LogWrite "error: unable to delete $path"
                write-host "error: unable to delete $path"
            }
            else{
                LogWrite "$path successfully deleted"
                write-host "$path successfully deleted"
            }
        }
        else{
            LogWrite "$path successfully deleted"
            write-host "$path successfully deleted"
        }
    }
    else{
        LogWrite "unable to find $path"
        write-host "unable to find $path"
    }
}

LogWrite "** STARTING Teams Cache Clean Script **"
Stop-Process -Name "teams" -force
tskill outlook
$username = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]          
removeLeftOvers -path "C:\Users\$username\AppData\Roaming\Microsoft\teams"
removeLeftOvers -path "C:\Users\$username\AppData\Roaming\teams"

LogWrite "** ENDING Teams Cache Clean Script **"
