# Script for SAP EPM reinstall
# Created by Marek.Kapusta@fujitsu.com
# Last update: 22.03.2023 15:44

# Function to generate a timestamp that is added to the log file
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

# Function to generate a log file
$logsDir = "$ENV:SystemDrive\support\logs\apps"
if (!(Test-Path -Path $logsDir -PathType Container)) {
    New-Item -ItemType Directory -Path $logsDir | Out-Null
}
$LogFile = "$logsDir\SapEPMReinstall.log"

Function LogWrite {
    Param ([string]$logstring)
    Add-Content $Logfile -Value "$(Get-TimeStamp) $logstring"
    Write-Host $logstring
}

# Function to check if SAP ERP is installed
function Check-SapERP {
    LogWrite "Checking if SAP ERP is installed"
    $SapERP = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq 'EPM add-in for microsoft office'} |  Select-Object DisplayName,DisplayVersion,Publisher,InstallDate
    if($SapERP -ne ""){
        LogWrite "One ERP is installed"
        LogWrite ("version: "+$SapERP.DisplayVersion)
    }
    else {LogWrite "SAP ERP is not installed"
    }
}

$ErrorActionPreference = 'SilentlyContinue'

LogWrite "** STARTING SAP ERP Reinstall Script **"

Check-SapERP

LogWrite "Trasnfering SAP EPM package to C:/support/SapEPM"
Copy-Item \\vfipffil005\HelpDesk\MyIS\fixes\SapEPM -Destination C:\Support\ -Recurse



#End of the script
LogWrite "** End of the SAP ERP Reinstall Script **"
