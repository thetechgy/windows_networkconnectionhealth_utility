#region SCRIPT HEADER
<#
    .SYNOPSIS
        Test Network Connection Health
    
        .DESCRIPTION
            This script completes steps related to testing the health of the local device's network connections.
        
    .NOTES
        =======================================
        Created by:  Travis McDade
        Created on:  03/30/2022
        Filename:  Test-NetworkConnectionHealth.ps1
        Version:  1.0
        Updated:  06/03/2022
        Known Issues:  None
        Prerequisites:  
		Related Links: https://www.speedtest.net/apps/cli
                       https://www.cyberdrain.com/monitoring-with-powershell-monitoring-internet-speeds/
         =======================================
#>
#endregion

#region SET VARIABLES
$CompanyName = "Company"
$CompanyHelpdeskURL = "https://helpdesk.company.com"
$LocalPingResultsFileDest = "$($Env:TEMP)\LocalPingResults.txt"
$DownloadURL = "https://install.speedtest.net/app/cli/ookla-speedtest-1.1.1-win64.zip"
$DownloadLocation = "$($Env:ProgramData)\SpeedtestCLI"
##### Absolute monitoring values ####
$MinimumDownloadSpeed = 50 #What is the minimum expected download speed in Mbps
$MinimumUploadSpeed = 5 #What is the minimum expected upload speed in Mbps
$MaxLatency = 60 #What is the maximimum latency until we alert in ms
$MaxJitter = 20 #What is the maximimum jitter until we alert in ms
$MaxPacketLoss = 2 #How much % packetloss until we alert
#### End absolute monitoring values ####
#endregion

#region WELCOME AND DESCRIPTION
Write-Host "Welcome to the $($CompanyName) Network Connection Test Utility!" -BackgroundColor Blue
Write-Host "`r"
Write-Host "The purpose of this utility is to gather information about your network performance to identify any potential issues."
Write-Host "`r"
Write-Host "Press any key to continue running the test. To exit without running the test, simply close this window." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
Write-Host "`r"
#endregion

#region LOCAL NETWORK TEST
Write-Host "LOCAL NETWORK INFO"

#Get active physical network adapters to determine wifi vs ethernet
$LocalNetworkConnectionType = Get-NetAdapter -Physical | Where-Object Status -eq Up | Foreach-Object Name
Write-Host "Local Network Connection Type: $LocalNetworkConnectionType"


#Get default gateway
$LocalDefaultGateway = Get-NetRoute | Where-Object RouteMetric -eq 0 | ForEach-Object NextHop


Write-Host "`r"
Write-Host "Testing local network - please wait..." -ForegroundColor Green
$PingCount = 50
ping -n $PingCount $LocalDefaultGateway | Out-File $LocalPingResultsFileDest
Write-Host "LOCAL NETWORK CONNECTION TEST RESULTS"
$LocalPingResults = Get-Content -Path $LocalPingResultsFileDest -Tail 4
Write-Output $LocalPingResults
#Remove ping temp file
Remove-Item -Path $LocalPingResultsFileDest
Write-Host "`r"
#endregion

#region DOWNLOAD AND EXTRACT SPEEDTEST CLI
try {
    $TestDownloadLocation = Test-Path $DownloadLocation
    if (!$TestDownloadLocation) {
        Write-Host "Downloading and extracting the Speedtest CLI utility - please wait..." -ForegroundColor Green
        $global:progressPreference = 'silentlyContinue'
        New-Item $DownloadLocation -ItemType Directory -Force | Out-Null
        Invoke-WebRequest -Uri $DownloadURL -OutFile "$($DownloadLocation)\speedtest.zip"
        Expand-Archive "$($DownloadLocation)\speedtest.zip" -DestinationPath $DownloadLocation -Force
        Remove-Item "$($DownloadLocation)\speedtest.zip"
        $global:progressPreference = 'Continue'
    } 
}
catch {  
    Write-Host "The download and extraction of Speedtest CLI utility failed. Error: $($_.Exception.Message)" -ForegroundColor Yellow
    exit 1
}
#endregion


#region RUN SPEEDTEST CLI
Write-Host "Testing internet connection - please wait..." -ForegroundColor Green
$SpeedtestResults = & "$($DownloadLocation)\speedtest.exe" --accept-license --format=json | ConvertFrom-Json
#endregion

#Store the result data in a hashtable
[PSCustomObject]$SpeedtestObj = @{
    DownloadSpeed = [math]::Round($SpeedtestResults.download.bandwidth / 1000000 * 8, 2)
    UploadSpeed   = [math]::Round($SpeedtestResults.upload.bandwidth / 1000000 * 8, 2)
    PacketLoss    = [math]::Round($SpeedtestResults.packetLoss)
    ISP           = $SpeedtestResults.isp
    ExternalIP    = $SpeedtestResults.interface.externalIp
    InternalIP    = $SpeedtestResults.interface.internalIp
    UsedServer    = $SpeedtestResults.server.host
    ResultsURL    = $SpeedtestResults.result.url
    Jitter        = [math]::Round($SpeedtestResults.ping.jitter)
    Latency       = [math]::Round($SpeedtestResults.ping.latency)
}
$SpeedtestHealth = @()

#Outputting useful data
Write-Host "INTERNET CONNECTION TEST RESULTS"
Write-Host "Download Speed: $($SpeedtestObj.DownloadSpeed) Mbps"
Write-Host "Upload Speed: $($SpeedtestObj.UploadSpeed) Mbps"
Write-Host "Latency: $($SpeedtestObj.Latency) ms"
Write-Host "Jitter: $($SpeedtestObj.Jitter) ms"
Write-Host "Packet Loss: $($SpeedtestObj.PacketLoss)%"
Write-Host "Public IP Address: $($SpeedtestObj.ExternalIp)"
Write-Host "ISP: $($SpeedtestObj.ISP)"
#Write-Host "Speedtest Server Used: $($SpeedtestObj.UsedServer)"
Write-Host "Speedtest Results URL: $($SpeedtestObj.ResultsURL)"
#Write-Host "Private IP Address: $($SpeedtestObj.InternalIP)"


#Comparing against preset variables.
Write-Host "`r"
if ($SpeedtestObj.DownloadSpeed -lt $MinimumDownloadSpeed) { $SpeedtestHealth += "Warning: Download speed is lower than the recommended $MinimumDownloadSpeed Mbps" }
if ($SpeedtestObj.UploadSpeed -lt $MinimumUploadSpeed) { $SpeedtestHealth += "Warning: Upload speed is lower than the recommended $MinimumUploadSpeed Mbps" }
if ($SpeedtestObj.Latency -gt $MaxLatency) { $SpeedtestHealth += "Warning: Latency is higher than the recommended $MaxLatency ms" }
if ($SpeedtestObj.Jitter -gt $MaxJitter) { $SpeedtestHealth += "Warning: Jitter is higher than the recommended $MaxJitter ms" }
if ($SpeedtestObj.PacketLoss -gt $MaxPacketLoss) { $SpeedtestHealth += "Warning: Packet loss is higher than the recommended $MaxPacketLoss %" }
 
if (!$SpeedtestHealth) {
    $SpeedtestHealth = "Healthy"
    Write-Host "Your connection test results look good - no problems detected!" -BackgroundColor Green
    Write-Host "`r"
}
else {
    Write-Host "Your connection has some issues - see below for details:" -BackgroundColor Red
    Write-Output $SpeedtestHealth
    Write-Host "$($CompanyName) IT recommends unplugging your router for 60 seconds, plugging it back in, restarting your computer and re-running the test." -ForegroundColor Yellow
    Write-Host "If you're still receiving warnings above after completing those steps, please submit a ticket at $($CompanyHelpdeskURL) so that $($CompanyName) IT can assist you further." -ForegroundColor Yellow
    Write-Host "`r"
}