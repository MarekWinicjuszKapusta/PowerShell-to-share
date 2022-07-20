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

LogWrite "** STARTING Clean Teams Removal Script **"
write-host "** STARTING Clean Teams Removal Script **"
#checking if Teams Wide Installer is installed
#if Teams Wide installer is not installed, teams won't reinstall after device is restarted, leaving user without MS Teams
if(test-path "C:\Program Files (x86)\Teams Installer\Teams.exe"){
    $user = (Get-WMIObject -query "SELECT * FROM win32_Process WHERE Name ='explorer.exe'" | Foreach { $owner = $_.GetOwner(); $_ | Add-Member -MemberType "Noteproperty" -name "Owner" -value $("{0}\{1}" -f $owner.Domain, $owner.User) -passthru }).Owner
    $username = $user[0].Split('\')[1]
    tskill outlook
    if (Test-Path "$($ENV:SystemDrive)\Users\$username\AppData\Local\Microsoft\Teams\update.exe") { 
                try {
                    LogWrite "Teams folder with update.exe found for user $username, uninstalling MS Teams..."
                    write-host "Teams folder with update.exe found for user $username, uninstalling MS Teams..."
                    Start-Process -FilePath "$($ENV:SystemDrive)\Users\$username\AppData\Local\Microsoft\Teams\Update.exe" -ArgumentList "-uninstall -s" -EA Stop            
                    removeLeftOvers -path "C:\Users\$username\AppData\Roaming\Microsoft\teams"
                    removeLeftOvers -path "C:\Users\$username\AppData\Roaming\teams"
                    removeLeftOvers -path "C:\Users\$username\AppData\local\Microsoft\teams"
                    removeLeftOvers -path "C:\Users\$username\AppData\local\Microsoft\teamsMeetingAddin"
                    removeLeftOvers -path "C:\Users\$username\AppData\local\Microsoft\teamsPresenceAddin"
                    removeLeftOvers -path "C:\Users\$username\AppData\local\Microsoft\SquirrelTemp"
                }
                Catch { 
                LogWrite "Teams app Uninstall for user $username Failed! Error Message:"
                write-host "Teams app Uninstall for user $username Failed! Error Message:"
                LogWrite $_.Exception.Message
                write-host $_.Exception.Message
                Out-Null
                }
            }
    else{
        LogWrite "MS Teams for $username not found"
        write-host "MS Teams for $username not found"
    }
    
    LogWrite "** ENDING Clean Teams Removal Script **"
    write-host "** ENDING Clean Teams Removal Script **"
}#End of the script with installed Windows wide installer
else{
    LogWrite "** Warning! MS team wide installer is not installed **"
    write-host "** Warning! MS team wide installer is not installed **"
}
