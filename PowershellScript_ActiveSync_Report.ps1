# Set the comment as a variable
$Comment = "Thank you for using my script, check my github page for more https://github.com/emad-mukhtar"

# Show the comment in the console
Write-Host $Comment

# Prompt the user to enter the server name
$ServerName = Read-Host "Enter the server FQDN name"

# Prompt the user to enter credentials
$Credentials = Get-Credential

# Connect to the specified Exchange server using the provided credentials
Connect-ExchangeServer -Server $ServerName -Credential $Credentials

# Prompt the user to enter a save location for the report
$SaveLocation = Read-Host "Enter the save location for the report (e.g. C:\Reports)"

# Check if the Save Location directory exists
if (!(Test-Path -Path $SaveLocation -PathType Container)) {
    # Prompt the user to create the directory if it doesn't exist
    $create = Read-Host "The directory $SaveLocation does not exist, would you like to create it? (Y/N)"
    if ($create -eq "Y" -or $create -eq "y") {
        New-Item -ItemType Directory -Path $SaveLocation
    } else {
        Write-Host "The script will now exit"
        Exit
    }
}

# Retrieve all mobile devices configured in the Exchange server
$Devices = Get-MobileDevice

# Filter the retrieved devices to only include those with an "Allowed" device access state
$EnabledDevices = $Devices | Where-Object {$_.DeviceAccessState -eq 'Allowed'}

# Initialize an empty array to store the report
$Report = @()

# Loop through each enabled device
foreach ($Device in $EnabledDevices) {
    # Create an object with the specified properties
    $Properties = @{
    'User Principal Name' = $Device.UserDisplayName
    'Device ID' = $Device.DeviceId
    'Device Type' = $Device.DeviceType
    'First Sync' = $Device.FirstSyncTime
    }

    # Add the object to the report array
    $Report += New-Object -TypeName PSObject -Property $Properties
}

# Get the current date and time in the specified format
$Date = Get-Date -Format yyyyMMdd_HHmmss

# Add the Comment to the top of the exported report
$Report |  Select-Object -Property 'https://github.com/emad-mukhtar',* |
# Export the report array to a CSV file with the specified save location, file name, and date format
Export-Csv -Path "$SaveLocation\ActiveSyncReport $Date.csv" -NoTypeInformation

# Disconnect from the Exchange server
Disconnect-ExchangeServer -Confirm:$false
