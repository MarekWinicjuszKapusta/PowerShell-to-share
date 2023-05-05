# Script for network stablity testing
# Created by Marek.Kapusta@fujitsu.com
# Last update: 05.05.2023 12:41
# Script will run 10x1 hour

$testCount = 10

# Loop through the test runs
for ($i = 1; $i -le $testCount; $i++) {
    
    $LogFile = "$(Get-Date -Format "yyyy-MM-dd-HH-mm").txt"
    Start-Transcript -Path "c:/support/logs/apps/$LogFile"
    #3600 = 1 hour
    Ping.exe -n 3600 office.com | ForEach {"{0} - {1}" -f (Get-Date),$_}
    Stop-Transcript

}
