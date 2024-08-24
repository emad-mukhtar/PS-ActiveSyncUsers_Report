# Exchange ActiveSync Report Generator
# Author: Emad Mukhtar (https://github.com/emad-mukhtar)

# Function to test Exchange connection
function Test-ExchangeConnection {
    try {
        Get-ExchangeServer -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Function to ensure valid directory path
function Get-ValidDirectory {
    param ([string]$prompt)
    do {
        $path = Read-Host $prompt
        if (!(Test-Path -Path $path -PathType Container)) {
            $create = Read-Host "The directory $path does not exist. Create it? (Y/N)"
            if ($create -eq "Y" -or $create -eq "y") {
                New-Item -ItemType Directory -Path $path | Out-Null
            }
            else {
                Write-Host "Please provide a valid directory path."
                $path = $null
            }
        }
    } while (!$path)
    return $path
}

# Set the comment as a variable
$Comment = "Thank you for using my script. Check my GitHub page for more: https://github.com/emad-mukhtar"

# Show the comment in the console
Write-Host $Comment -ForegroundColor Cyan

# Prompt for server name and credentials
$ServerName = Read-Host "Enter the MS Exchange server FQDN name"
$Credentials = Get-Credential -Message "Enter your Exchange admin credentials"

# Connect to Exchange server
try {
    Connect-ExchangeServer -Server $ServerName -Credential $Credentials -ErrorAction Stop
    Write-Host "Successfully connected to Exchange server." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange server: $_" -ForegroundColor Red
    exit
}

# Verify Exchange connection
if (-not (Test-ExchangeConnection)) {
    Write-Host "Failed to establish a connection to Exchange." -ForegroundColor Red
    exit
}

# Get valid save location
$SaveLocation = Get-ValidDirectory "Enter the save location for the report (e.g., C:\Reports)"

# Retrieve and filter mobile devices
try {
    $Devices = Get-MobileDevice -ErrorAction Stop
    $EnabledDevices = $Devices | Where-Object {$_.DeviceAccessState -eq 'Allowed'}
    Write-Host "Retrieved $(($EnabledDevices | Measure-Object).Count) enabled devices." -ForegroundColor Green
}
catch {
    Write-Host "Error retrieving mobile devices: $_" -ForegroundColor Red
    exit
}

# Generate report
$Report = @()
foreach ($Device in $EnabledDevices) {
    $Properties = @{
        'User Principal Name' = $Device.UserDisplayName
        'Device ID' = $Device.DeviceId
        'Device Type' = $Device.DeviceType
        'First Sync' = $Device.FirstSyncTime
        'Last Sync' = $Device.LastSuccessSync
        'Device OS' = $Device.DeviceOS
        'Device Model' = $Device.DeviceModel
    }
    $Report += New-Object -TypeName PSObject -Property $Properties
}

# Export report
$Date = Get-Date -Format yyyyMMdd_HHmmss
$ReportPath = Join-Path $SaveLocation "ActiveSyncReport_$Date.csv"
try {
    $Report | Export-Csv -Path $ReportPath -NoTypeInformation -ErrorAction Stop
    Write-Host "Report successfully saved to: $ReportPath" -ForegroundColor Green
}
catch {
    Write-Host "Error saving report: $_" -ForegroundColor Red
}

# Disconnect from Exchange server
Disconnect-ExchangeServer -Confirm:$false
Write-Host "Disconnected from Exchange server." -ForegroundColor Green

Write-Host "Script completed. Thank you for using this tool!" -ForegroundColor Cyan
