#region GLOBAL CONFIG

# BEFORE YOU BEGIN... Please edit Settings.ps1 and ScriptConfig.ps1 as appropriate.
# All initial inspiration and examples from https://www.reddit.com/r/PowerShell/comments/3sehai/connectwise_rest_api_miniguide/

# Notes & ToDos
#   TODO: Is there a basic API call we can do to see if we are authed, so we can run that before running any commands?
#         Currently we try to run the command we need, then if we need to re-auth, we do so, then we run the command again
#         but it would be much better if we could run a single command at the start then just get on with the script/our lives.
#
#   TODO: - Get all CW contacts & AD users in a group and use Compare-Object to see if any are missing.
#         - Add a -DryRun parameter (Perhaps defaults to this and needs a paramater such as -Run to make changes)
#         - Record all changes and allow a revert?
#         - Make it only ever apply to a single Company (And perhaps create some verification)
#         - Add more validation for the 2 config files??
#         - Make it GET a CW user first (Function needed) to see if the one we're trying to create already exists
#

$ScriptConfigPath      = "$PSScriptRoot\ScriptConfig.ps1"
$ConnectWiseConfigPath = "$PSScriptRoot\ConnectWiseConfig.ps1"

# This will source the entire $ScriptConfig variable from our script config file
if (!(Test-Path $ScriptConfigPath)) {
    Throw 'FATAL ERROR: ScriptConfig.ps1 not found.' 
} else {
    . $ScriptConfigPath
}

# This will source the entire $CWConfig variable from our ConnectWise config file
if (!(Test-Path $ConnectWiseConfigPath)) {
    Throw 'FATAL ERROR: ConnectWiseConfig.ps1 not found.'
} else {
    . $ConnectWiseConfigPath
}

# These are strings used in our REST statements
$global:Accept              = 'application/vnd.connectwise.com+json; version=v2015_3'
$global:ContentType         = 'application/json'
$global:BaseURI             = "https://$($CWConfig.ServerBaseURL)/v4_6_Release/apis/3.0"

# NOTE:
# The following variables are used without being passed to functions (This would stop
# this from being modularised without refactoring)
# $CWConfig.CompanyName

#endregion GLOBAL CONFIG

#region FUNCTIONS


# TODO: Add a second parameterset to each function that just accepts an object
#       containing the parameters

function Get-CWKeys
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$ServerBaseURL,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$ImpersonationMember,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$BaseURI
    )

    # Define the base URI for this function
    $FunctionBaseURI    = "$BaseURI/system/members/$ImpersonationMember/tokens"
    
    # Format the auth string required by REST API
    $CWAuthString = '{0}+{1}:{2}' -f $CWConfig.CompanyName, $CWConfig.Username, $CWConfig.Password

    # Convert the user and pass (aka public and private key) to base64 encoding
    $encodedAuth  = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($CWAuthstring)));

    # Create the message header
    $Header = @{
        Authorization    = ('Basic {0}' -f $encodedAuth)
        Accept           = $Accept
        Type             = 'application/json'
        'x-cw-usertype'  = 'integrator'
    };

    # Create the message body
    $Body   = @"
memberIdentifier":"$ImpersonationMember"
"@

    # Execute the request
    $Response = Invoke-RestMethod -Uri $FunctionBaseuri -Method Post -Headers $Header -Body $Body -ContentType $ContentType

    #return the results
    return $Response

}


Function Get-CWContact
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Company,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$ServerBaseURL,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$CompanyName,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$BaseURI,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [psobject]$CWAuth
    )

    # Define the base URI for this function
    $FunctionBaseURI = "$BaseURI/company/contacts"

    if (!($CWAuth)) { 
        Throw 'Invalid CWCredentials defined.'
        $CWAuth
    }

    # Format the auth string required by REST API
    $CWAuthString = "$($CWConfig.CompanyName)+" + $($CWAuth.publickey) + ':' + $($CWAuth.privatekey)

    # Convert the user and pass (aka public and private key) to base64 encoding
    $EncodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($CWAuthstring)))

    # Create the message header
    $Header = @{
        "Authorization"="Basic $encodedAuth"
    }

    # Create the message body
    $Body = @{
        # CONDITIONS: "conditions" = "firstname LIKE `"$`" AND lastname LIKE `"Marley`" AND company/id = 12345"
        "conditions" = "company/identifier LIKE `"$Company`""
    }

    # Execute the request
    $JSONResponse = Invoke-RestMethod -URI $FunctionBaseURI -Headers $Header -ContentType $ContentType -Body $Body -Method Get

    If($JSONResponse)
    {
        Return $JSONResponse
    }

    Else
    {
        Return $False
    }

}

function New-Credential {
  <#
  .SYNOPSIS
    Create a new PSCredential object
  .EXAMPLE
    $MyCred = New-Credential 'MyDomain\Username' 'Password'
  .EXAMPLE
    PS C:\>MyCommand -Credential (New-Credential -Username 'MyDomain\Username' -Password 'Password')
  .PARAMETER Username
    Username.
  .PARAMETER Password
    Password.
  #>

  [CmdletBinding()]
  param
  (
    [Parameter(
      Mandatory=$True,
      ValueFromPipeline=$True,
      ValueFromPipelineByPropertyName=$True,
      Position=0,
      HelpMessage='Enter username')]
      [ValidateNotNullOrEmpty()]
    [string]$Username,

    [Parameter(
      Mandatory=$True,
      ValueFromPipelineByPropertyName=$True,
      Position=1,
      HelpMessage='Enter password')]
    [ValidateNotNullOrEmpty()]
    [string]$Password

  )

  process {
    Write-Verbose "Create credential object for $Username"

    Return New-Object System.Management.Automation.PSCredential `
      ($Username,(ConvertTo-SecureString -String $Password -AsPlainText -Force))

  }

}

#endregion FUNCTIONS

#region SCRIPT BODY

#
#region    Authenticate with ConnectWise
#

# Get our ConnectWise authentication token
#   Create a function that has our command & arguments (DRY!)
$AuthArgs = @{
    ServerBaseUrl       = $CWConfig.ServerBaseUrl
    ImpersonationMember = $CWConfig.ImpersonationMember
    BaseURI             = $BaseURI
}

# Now, if we don't have an auth token, then get one! 
#   This is just for use in the ISE really as we don't need to get a new
#   8 hour authentication token every time we re-run the script.
If (!($CWauth)) {
    Write-Host 'No authentication token exists, requesting one.' -ForegroundColor Gray
    $global:CWAuth = Get-CWKeys @AuthArgs
} else {
    Write-Host 'Authentication token already exists, using that.' -ForegroundColor Gray
}                        

#endregion

<#
#
#region    Find ConnectWise contacts
#

# Now, we find our contacts
#   Create a function that has our command & arguments (DRY!)
$GetContactsArgs = @{
    Company       = $ScriptConfig.CWCompany
    ServerBaseURL = $CWConfig.ServerBaseURL
    CompanyName   = $CWConfig.CompanyName
    CWAuth        = $CWAuth
    BaseURI       = $BaseURI
}

Try {
    $CWContacts = Get-CWContact @GetContactsArgs

} catch {
    # NOTE: Exception is $_
    
    if ($_.Exception.Message -contains '(401) Unauthorized.') {
        Write-Warning 'Getting "Unauthorized" error from ConnectWise API - Has auth token expired?'

        # Let's try authenticating again
        $CWAuth = Get-CWKeys @AuthArgs

        $CWContacts = Get-CWContact @GetContactsArgs
    }
} finally {
    # If we have contacts returns, then...
    if ($CWContacts) {
        Write-Host "$($CWContacts.Count) results returned" -ForegroundColor Green

        # Print them all for the time being
        ForEach ($CWContact in $CWContacts) {
            Write-Host ("- {0} {1}" -f $CWContact.FirstName, $CWContact.LastName)
        }
    } else {
    # No contacts returned, write a warning (Probably a bad 
        Write-Warning 'No contacts returned, please check the CWCompany value in the Script Config file.'
    }
}

#endregion

#>


#
#region    Add a new contact

function New-CWContact {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Company,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$ServerBaseURL,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$CompanyName,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$BaseURI,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [psobject]$CWAuth,

        [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [psobject]$NewUserObject
    )

    # Define the base URI for this function
    $FunctionBaseURI = "$BaseURI/company/contacts"

    if (!($CWAuth)) { 
        Throw 'Invalid CWCredentials defined.'
        $CWAuth
    }

    # Format the auth string required by REST API
    $CWAuthString = "$($CWConfig.CompanyName)+" + $($CWAuth.publickey) + ':' + $($CWAuth.privatekey)

    # Convert the user and pass (aka public and private key) to base64 encoding
    $EncodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($CWAuthstring)))

    $Header = @{
        "Authorization"="Basic $encodedAuth"
    }
    
    # Create the message body
    $Body =  $NewUserObject

    # Execute the request
    $JSONResponse = Invoke-RestMethod -URI $FunctionBaseURI -Headers $Header -ContentType $ContentType -Body $Body -Method Post

    If($JSONResponse)
    {
        Return $JSONResponse
    }

    Else
    {
        Return $False
    }

}

#   Create a function that has our command & arguments (DRY!)
$NewContactArgs = @{
    Company       = $ScriptConfig.CWCompany
    ServerBaseURL = $CWConfig.ServerBaseURL
    CompanyName   = $CWConfig.CompanyName
    CWAuth        = $CWAuth
    BaseURI       = $BaseURI
    NewUserObject = @"
{
    "firstName" : "James",
    "lastName" : "Booth",
    "company" : {
        "identifier" : "MagicRoosterLLC"
    },
    "portalPassword": "portal",
    "disablePortalLoginFlag": false,
    "inactiveFlag": false,
    "unsubscribeFlag": true
}
"@
}

try {
    $CWNewUserResponse = New-CWContact @NewContactArgs

    $CWNewUserResponse
} catch {
    Write-Error $_
}

#endregion

#endregion SCRIPT BODY