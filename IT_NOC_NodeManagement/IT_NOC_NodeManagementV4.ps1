## Last Edited: June 9th, 2023

## This script connects to the Orion server and performs specified actions on the node it is run on.

## It requires the hostname of the Orion server, the desired action ("MuteAlerts", "UnmuteAlerts", "UnmanageNode", or "ManageNode"),
## and an account name with its associated encryption key.

## The account used must have sufficient permissions on the Orion server (typically the Orion service account).

## Actions include muting and unmuting alerts, managing and unmanaging a node, and repolling a node after it has been remanaged.

## The encryption key for the account password is expected as a comma-separated string.
## It is transcribed into a byte array to convert the encrypted password back into a secure string.

## The encrypted password is loaded from a 'crypto.txt' file located in the same directory as the script.

## Logs are generated for each operation and are saved to a file at C:\Temp\IT_NOC_NodeManagement\NodeActions.log.

## Exit codes:
## 0: The script completed successfully and the action was performed.
## 1: The SwisPowerShell module was not found on the machine where the script is run.
## 2: Connection to the Orion server could not be established.
## 3: The script failed to retrieve node information from the Orion server.
## 4: The script failed to perform the action on the node.

## Notes:
## The script is to be run with an account that has sufficient permissions on the Orion server.
## This is typically the Orion service account.
## The script will create a log file at C:\Temp\IT_NOC_NodeManagement\NodeActions.log.
## The log file will contain detailed information about each operation that the script performs.
## All of the above information and processes are made explicit through the debug log outputs in the script.


param(
    [Parameter(Mandatory=$true)]
    [string]$hostname,
    [Parameter(Mandatory=$true)]
    [ValidateSet("MuteAlerts", "UnmuteAlerts", "UnmanageNode", "ManageNode")]
    [string]$Action,
    [Parameter(Mandatory=$true)]
    [string]$AccountName,
    [Parameter(Mandatory=$true)]
    [string]$AccEncryptionKey
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
$LogFolderPath = "C:\Temp\IT_NOC_NodeManagement"

# Create folders if they don't exist
if(!(Test-Path -Path $LogFolderPath)) {
    New-Item -ItemType Directory -Force -Path $LogFolderPath
}

# Custom logging function
function debug($message) {
    $logMessage = "$(Get-Date -Format yyyy-MM-dd--HH-mm-ss) $message"
    Write-Host $logMessage
    Add-Content -Path "$LogFolderPath\NodeActions.log" -Value $logMessage -Force
}

debug "------------------------------------------------------------------------------------------------------------------------------------------------"

# Check if the SwisPowerShell module is available
if (-not (Get-Module -Name SwisPowerShell -ListAvailable)) {
    
    debug "SwisPowerShell module was not found. Exiting with error code 1..."
    
    exit 1
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

    exit 2
}

debug "Successfully connected to the Orion server."

debug "Proceeding to get the machine name..."

# Get the machine's fully qualified domain name
$MachineName = $env:COMPUTERNAME

debug "Machine name: $MachineName."

switch ($Action) {
    'MuteAlerts' {
        
        debug "Preparing command to retrieve node info..."

        # Query to get Node details
        $query = "SELECT TOP 1 Uri FROM Orion.Nodes WHERE SysName = '$MachineName'"

        debug "Running SWQL query: $query"

        # Run the SWQL query
        $nodeUri = Get-SwisData -SwisConnection $swis -Query $query -ErrorAction SilentlyContinue

        if(!($nodeUri))
        {
            debug "Failed to get node info. Exiting..."
            exit 3
        }

        debug "Successfully retrieved node info."

        debug "Converting URI to string.."

        $nodeUri = $nodeUri.ToString()

        debug "Node URI: $nodeUri"

        debug "Performing $Action on the Node..."
        
        $response = Invoke-SwisVerb -SwisConnection $swis -EntityName 'Orion.AlertSuppression' -Verb 'SuppressAlerts' -Arguments @(,$nodeUri, [DateTime]::UtcNow) -ErrorAction SilentlyContinue
    }
    'UnmuteAlerts' {
        
        debug "Preparing command to retrieve node info..."

        # Query to get Node details
        $query = "SELECT TOP 1 Uri FROM Orion.Nodes WHERE SysName = '$MachineName'"

        debug "Running SWQL query: $query"

        # Run the SWQL query
        $nodeUri = Get-SwisData -SwisConnection $swis -Query $query -ErrorAction SilentlyContinue

        if(!($nodeUri))
        {
            debug "Failed to get node info. Exiting..."
            exit 3
        }

        debug "Successfully retrieved node info."

        debug "Node URI: $nodeUri"

        $nodeUri = @( $nodeUri |% {[string]$_} )

        debug "Performing $Action on the Node..."
        
        $response = Invoke-SwisVerb -SwisConnection $swis -EntityName 'Orion.AlertSuppression' -Verb 'ResumeAlerts' -Arguments @(,$nodeUri) -ErrorAction SilentlyContinue
    }
    'UnmanageNode' {
        
        debug "Preparing command to retrieve node ID..."

        # Query to get Node details
        $query = "SELECT NodeID FROM Orion.Nodes WHERE Caption = '$MachineName'"

        debug "Running SWQL query: $query"

        # Run the SWQL query
        $nodeID = Get-SwisData -SwisConnection $swis -Query $query -ErrorAction SilentlyContinue

        if(!($nodeID))
        {
            debug "Failed to get node ID. Exiting..."
            exit 3
        }

        debug "Successfully retrieved node ID."

        debug "Constructing a descriptor for the NodeID.."

        $nodeDescriptor = "N:$nodeID"

        debug "Constructed Node Descriptor: $nodeDescriptor"

        $UnmanageEnd = Get-Date

        $UnmanageEnd.AddYears(10) | Out-Null

        debug "Performing $Action on the Node..."
        
        $response = Invoke-SwisVerb -SwisConnection $swis -EntityName 'Orion.Nodes' -Verb 'Unmanage' -Arguments @($nodeDescriptor, [DateTime]::UtcNow, $UnmanageEnd, "false") -ErrorAction SilentlyContinue
    }
    'ManageNode' {
        
        debug "Preparing command to retrieve node ID..."

        # Query to get Node details
        $query = "SELECT NodeID FROM Orion.Nodes WHERE Caption = '$MachineName'"

        debug "Running SWQL query: $query"

        # Run the SWQL query
        $nodeID = Get-SwisData -SwisConnection $swis -Query $query -ErrorAction SilentlyContinue

        if(!($nodeID))
        {
            debug "Failed to get node ID. Exiting..."
            exit 3
        }

        debug "Successfully retrieved node ID."

        debug "Constructing a descriptor for the NodeID.."

        $nodeDescriptor = "N:$nodeID"

        debug "Constructed Node Descriptor: $nodeDescriptor"

        $ManageEnd = Get-Date

        $ManageEnd.AddYears(99) | Out-Null

        debug "Performing $Action on the Node..."
        
        $response = Invoke-SwisVerb -SwisConnection $swis -EntityName 'Orion.Nodes' -Verb 'Remanage' -Arguments @($nodeDescriptor, [DateTime]::UtcNow, $ManageEnd, "false")
    }
}

if($Action -eq "ManageNode")
{
    debug "Action initiated as to Remanage a node. Initiating and repolling of the node..."
    
    debug "Preparing command to retrieve node ID..."

    # Query to get Node details
    $query = "SELECT NodeID FROM Orion.Nodes WHERE Caption = '$MachineName'"

    debug "Running SWQL query: $query"

    # Run the SWQL query
    $nodeID = Get-SwisData -SwisConnection $swis -Query $query -ErrorAction SilentlyContinue

    if(!($nodeID))
    {
        debug "Failed to get node ID. Repolling will not be performed immediately. Exiting..."
        exit 0
    }

    debug "Successfully retrieved node ID."

    debug "Constructing a descriptor for the NodeID.."

    $nodeDescriptor = "N:$nodeID"

    debug "Constructed Node Descriptor: $nodeDescriptor"

    $responseManagePost = Invoke-SwisVerb -SwisConnection $swis -EntityName 'Orion.Nodes' -Verb 'PollNow' -Arguments @($nodeDescriptor)

    if($responseManagePost)
    {
        debug "Node Repolling successful."
    }
    else
    {
        debug "Node Repolling FAILED."
    }

}

if ($response)
{
    debug "Node action $Action successful."

    debug "Script execution complete. Exiting with error code 0..."

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 0
}
else
{
    debug "Failed to perform action $Action on the Node. Exiting with error code 4..."

    debug "------------------------------------------------------------------------------------------------------------------------------------------------"

    exit 4
}