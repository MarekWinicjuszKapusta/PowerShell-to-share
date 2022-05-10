#Script stolen from Wdomon by Marek Kapusta. Thanks a lot dude, you saved my life.
#Function to generate a timestamp that is added to the log file
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)   
}
    
#Function to generate a log file
if ((Test-Path -Path "$ENV:SystemDrive\support\logs\apps" -PathType Container) -ne $true ) {mkdir "$ENV:SystemDrive\support\logs\apps" | Out-Null}
$LogFile = "$ENV:SystemDrive\support\logs\apps\TeamsRemoval.log"

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
        Get-ChildItem $path | Remove-Item -Recurse -Confirm:$false
        remove-item $Path -Recurse -Confirm:$false
        if(Test-Path $Path){
            LogWrite "error: complications with deleting $path"
            $attemp = 2
            while((Test-Path $Path)-and($attemp -le 10)){
            Start-Sleep -Seconds 10
            LogWrite "$attemp attempt to delete $path"
            Write-Host $Path
            Write-Host $attemp
            Get-ChildItem $path | Remove-Item -Recurse -Confirm:$false
            remove-item $Path -Recurse -Confirm:$false
            $attemp++
            }
            if(Test-Path $Path){
                LogWrite "error: unable to delete $path"
            }
            else{
                LogWrite "$path successfully deleted"
            }
        }
        else{
            LogWrite "$path successfully deleted"
        }
    }
    else{
        LogWrite "unable to find $path"
    }
}

LogWrite "** STARTING Uninstall MS Teams Script **"

# Removal Machine-Wide Installer - This needs to be done before removing the EXE!
# Completly removed part done Wdomon, as it was not working as it should. Previous creator tried to use try and catch, but when not finding wmi_object, powershell was not generating an error, therefore, it was not loggin that object was not found. -Marek
LogWrite "searching for Teams Machine-Wide Installer"
if($teams = Get-WmiObject Win32_Product | Where-Object {$_.name -like "Teams Machine-Wide Installer"}){
ForEach($team in $teams){
    LogWrite "Teams Machine-Wide Installer found, version: $($_.version)"
    $team.uninstall()
    }
    if($teams = Get-WmiObject Win32_Product | Where-Object {$_.name -like "Teams Machine-Wide Installer"}){
    LogWrite "error: Unable to remove Teams Machine-Wide Installer"
    }
    else{
    LogWrite "Successfully removed Teams Machine-Wide Installer"
    }
}
else{
    LogWrite "Teams Machine-Wide Installer not found"
}
# end of Machine-wide Installer removal

    #Variables
    $TeamsUsers = Get-ChildItem -Path "$($ENV:SystemDrive)\Users"

    $TeamsUsers | ForEach-Object {
        if (Test-Path "$($ENV:SystemDrive)\Users\$($_.Name)\AppData\Local\Microsoft\Teams\update.exe") { 
            try {
                LogWrite "Teams folder with update.exe found for user $($_.Name), uninstalling MS Teams..."
                Start-Process -FilePath "$($ENV:SystemDrive)\Users\$($_.Name)\AppData\Local\Microsoft\Teams\Update.exe" -ArgumentList "-uninstall -s" -EA Stop
                #Start-Sleep -Seconds 60            
                removeLeftOvers -path "C:\Users\$($_.Name)\AppData\Roaming\Microsoft\teams"
                removeLeftOvers -path "C:\Users\$($_.Name)\AppData\Roaming\teams"
                removeLeftOvers -path "C:\Users\$($_.Name)\AppData\local\Microsoft\teams"
                removeLeftOvers -path "C:\Users\$($_.Name)\AppData\local\Microsoft\teamsMeetingAddin"
                removeLeftOvers -path "C:\Users\$($_.Name)\AppData\local\Microsoft\teamsPresenceAddin"
                removeLeftOvers -path "C:\Users\$($_.Name)\AppData\local\Microsoft\SquirrelTemp"
            
            }
            Catch { 
            LogWrite "Teams app Uninstall for user $($_.Name) Failed! Error Message:"
            LogWrite $_.Exception.Message
            Out-Null
            }

        }
        else{
            LogWrite "MS Teams for $($_.Name) not found"
        }
    }

    LogWrite "** ENDING Uninstall MS Teams Script **"
