# Ensure the output directory exists
$logPath = "C:\support\logfile"
$logDir = Split-Path $logPath
if (!(Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

# Set the target host and duration
$TargetHost = "onet.pl"
$durationSeconds = 3600

# Initialize counters
$timeoutCount = 0

# Start the ping loop
for ($i = 0; $i -lt $durationSeconds; $i++) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $pingResult = Test-Connection -ComputerName $TargetHost -Count 1 -ErrorAction SilentlyContinue

    if ($pingResult) {
        $reply = "Reply from $($pingResult.Address): time=$($pingResult.ResponseTime)ms"
    } else {
        $reply = "Request timed out."
        $timeoutCount++
    }

    $logLine = "$timestamp - $reply"
    Write-Host $logLine
    Add-Content -Path $logPath -Value $logLine

    Start-Sleep -Seconds 1
}

# Calculate and display summary
$percentageTimeout = [math]::Round(($timeoutCount / $durationSeconds) * 100, 2)
$summary = @"

=== PING SUMMARY ===
Total pings:    $durationSeconds
Timeouts:       $timeoutCount
Success rate:   {0:N2}%
Timeout rate:   $percentageTimeout%
"@ -f (100 - $percentageTimeout)

Write-Host $summary
Add-Content -Path $logPath -Value $summary
