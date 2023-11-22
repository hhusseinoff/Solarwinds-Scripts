<#
.SYNOPSIS
    This script initializes and configures a node in the SolarWinds Orion environment.

.DESCRIPTION
    - Takes mandatory parameters like hostname, account name, encryption key, and various Orion settings.
    - Sets up logging directories and captures script execution details in a log file.
    - Checks for the availability of the SwisPowerShell module.
    - Converts an encryption key and retrieves an encrypted password.
    - Establishes a connection to the Orion server using provided credentials.
    - Retrieves the machine's fully qualified domain name.
    - Mutes alerts for the node in the Orion environment.
    - Assigns CITY and DEPARTMENT Custom node parameters according to what is passed at the script input
    - Adds the node to the relevant Orion Group, that is passed as an input param

.NOTES
    Ensure that the SwisPowerShell module is installed and that the necessary permissions are granted for script execution.
    The script also handles various error scenarios and logs them for troubleshooting.
    The script is intended to run under an account that has the appropriate permission on IT-NOC. Right now - ZONAL\Zservice


File Name      : Initialize_SW_V1.ps1
Last Modified  : Oct 5, 2023
Prerequisites  : OrionSDK and SwisPowershell Modules Installed, Connectivity with it-noc.zonal.co.uk via the swisPowershell module.
#>


param(
    [Parameter(Mandatory=$true)]
    [string]$hostname,
    [Parameter(Mandatory=$true)]
    [string]$AccountName,
    [Parameter(Mandatory=$true)]
    [string]$AccEncryptionKey,
    [Parameter(Mandatory=$true)]
    [string]$OrionCITY,
    [Parameter(Mandatory=$true)]
    [string]$OrionDEPARTMENT
    [Parameter(Mandatory=$true)]
    [string]$OrionGroupName
)

# Capture script directory
$ScriptDir = $PSScriptRoot

# Define log folder root
$LogFolderRoot = "C:\Temp"

# Create folders if they don't exist
if(!(Test-Path -Path $LogFolderRoot)) {
    New-Item -ItemType Directory -Force -Path $LogFolderRoot
}

# Define log folder path
$LogFolderPath = "C:\Temp\Solarwinds_Initialize"

# Create folders if they don't exist
if(!(Test-Path -Path $LogFolderPath)) {
    New-Item -ItemType Directory -Force -Path $LogFolderPath
}

# Custom logging function
function debug($message) {
    $logMessage = "$(Get-Date -Format yyyy-MM-dd--HH-mm-ss) $message"
    Write-Host $logMessage
    Add-Content -Path "$LogFolderPath\InitializationLog.log" -Value $logMessage -Force
}

debug "------------------------------------------------------------------------------------------------------------------------------------------------"

# Check if the SwisPowerShell module is available
if (-not (Get-Module -Name SwisPowerShell -ListAvailable)) {
    
    debug "SwisPowerShell module was not found. Exiting with error code 1..."
    
    exit 1

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"
}

debug "Starting the NodeAction script for Orion server: $hostname. Action: $Action"

debug "Converting encryption key from a comma-separated string to an array of strings..."

$AccEncryptionKeyStringArray = $AccEncryptionKey.Split("#")

debug "Creating a byte array object..."

$byteArray = New-Object Byte[] 16

$count = 0

debug "Transcribing the encryption key to the byte array..."

foreach($element in $AccEncryptionKeyStringArray)
{
    $elementInByte = [byte]$element
    $byteArray[$count] = $elementInByte
    $count = $count + 1
}

debug "Loading encrypted password from $ScriptDir\crypto.txt..."

$encryptedPassword = (Get-Content -Path "$ScriptDir\crypto.txt")

debug "Converting encrypted password to a secure string..."

$securePassword = ConvertTo-SecureString -String $encryptedPassword -Key $byteArray

debug "Constructing a PSCredential object..."

$credential = [pscredential]::new($AccountName,$securePassword)

debug "Connecting to the Orion server..."

# Connect to the Orion server
$swis = Connect-Swis -Hostname $hostname -Credential $credential -ErrorAction SilentlyContinue

debug "Validating if the connection was successful..."

$ConnectionValidation = Get-SwisData -SwisConnection $swis -Query 'SELECT TOP 1 NodeID FROM Orion.Nodes' -ErrorAction SilentlyContinue

if($null -eq $ConnectionValidation)
{
    debug "Connection establishment failed."

    debug "Error: $($Error[0]). Exiting with error code 2..."

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 2
}

debug "Successfully connected to the Orion server."

debug "Proceeding to get the machine name..."

# Get the machine's fully qualified domain name
$MachineName = $env:COMPUTERNAME

debug "Machine name: $MachineName."


##Action 1: Mute alerts

debug "ACTION 1: MUTING ALERTS"

debug "Preparing command to retrieve node info..."

$query = "SELECT TOP 1 Uri FROM Orion.Nodes WHERE SysName = '$MachineName'"

debug "Running SWQL query: $query"

$nodeUri = Get-SwisData -SwisConnection $swis -Query $query -ErrorAction SilentlyContinue

if(!($nodeUri))
{
    debug "Failed to get node info. Exiting..."

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 3
}

debug "Successfully retrieved node info."

debug "Converting URI to string.."

$nodeUri = $nodeUri.ToString()

debug "Node URI: $nodeUri"

debug "Performing MUTE on the Node..."

$response = Invoke-SwisVerb -SwisConnection $swis -EntityName 'Orion.AlertSuppression' -Verb 'SuppressAlerts' -Arguments @(,$nodeUri, [DateTime]::UtcNow) -ErrorAction SilentlyContinue

if(!($response))
{
    debug "Failed to MUTE $MachineName"

    debug "Exiting with code 4..."

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 4
}
else
{
    debug "ACTION 1 COMPLETED. NODE MUTED."
}

##Action 2: Add to appropriate group

debug "ACTION 2: ADD TO ORION GROUP $OrionGroupName"

debug "Retrieving the Orion GroupID for $OrionGroupName..."

$OrionGroupID = Get-SwisData -SwisConnection $swis -Query "SELECT ContainerID FROM Orion.Container WHERE Name = 'Tanfield'"

if(!($OrionGroupID))
{
    debug "FAILED to get Orion GroupID for $OrionGroupName..."

    debug "Exiting with code 5"

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 5
}
else
{
    debug "Successfully retrieved GroupID for $OrionGroupName : $OrionGroupID"

    debug "Proceeding to add $nodeUri to group $OrionGroupName..."

    $resultAddToGroup = Invoke-SwisVerb -SwisConnection $swis -EntityName "Orion.Container" -Verb "AddDefinition" -Arguments @(
    # group ID
    $OrionGroupID,

    # group member to add
    ([xml]"
       <MemberDefinitionInfo xmlns='http://schemas.solarwinds.com/2008/Orion'>
         <Name>Up Devices</Name>
         <Definition>$nodeUri</Definition>
       </MemberDefinitionInfo>"
    ).DocumentElement
  ) -ErrorAction SilentlyContinue

  if(!($resultAddToGroup))
  {
    debug "FAILED to add $MachineName to Orion group $OrionGroupName..."

    debug "Exiting with code 6..."

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 6
  }
  else
  {
    debug "Successfully added $MachineName to Orion group $OrionGroupName!"
  }
}

##action 3: edit node custom properties - department, location

debug "ACTION 3: ADD CUSTOM PROPERTIES"

debug "Appending 'CustomProperties' to NodeUri $nodeUri..."

$nodeUriCustomProperties = $nodeUri + "/CustomProperties"

debug "Getting SwisObject for CustomProperties on Uri $nodeUriCustomProperties..."

$NodeCustomPropsObj = Get-SwisObject -SwisConnection $swis -Uri $nodeUriCustomProperties -ErrorAction SilentlyContinue

if(!($NodeCustomPropsObj))
{
    debug "Failed to retrieve Custom Properties Swis Object for $MachineName..."
}
else
{
    debug "Successfully retrieved Custom Proerties Swis Object for $MachineName."

    debug "Setting Orion Custom properties City $OrionCITY and Department $OrionDEPARTMENT for $MachineName... "

    try
    {
        Set-SwisObject -SwisConnection $swis -Uri $nodeUriCustomProperties -Properties @{Department=$OrionDEPARTMENT;City=$OrionCITY} -ErrorAction SilentlyContinue
    }
    catch
    {
        debug "FAILED to add Custom Properties!!!"

        debug "Exiting with code 7"

        debug "------------------------------------------------------------------------------------------------------------------------------------------------"

        exit 7
    }

    debug "Successfully added Custom Properties!"
}