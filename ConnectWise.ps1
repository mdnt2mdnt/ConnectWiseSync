#region GLOBAL CONFIG

# All initial inspiration and examples from https://www.reddit.com/r/PowerShell/comments/3sehai/connectwise_rest_api_miniguide/

# Notes & ToDos
#   TODO: Is there a basic API call we can do to see if we are authed, so we can run that before running any commands?
#         Currently we try to run the command we need, then if we need to re-auth, we do so, then we run the command again
#         but it would be much better if we could run a single command at the start then just get on with the script/our lives.

# This will source the entire $CWDetails variable from our config file
. "$PSScriptRoot\Settings.ps1"

    # The config file should have the following...
    # $CWDetails = @{
    #    ServerBaseURL        = 'xxx'
    #    CompanyName          = 'xxx'
    #    Username             = 'xxx'
    #    Password             = 'xxx'
    #    ImpersonationMember  = 'xxx'
    # }

# These are strings used in our REST statements
$global:Accept              = 'application/vnd.connectwise.com+json; version=v2015_3'
$global:ContentType         = 'application/json'
$global:BaseURI             = "https://$($CWDetails.ServerBaseURL)/v4_6_Release/apis/3.0"

#endregion GLOBAL CONFIG

#region FUNCTIONS


# TODO: Add a second parameterset to each function that just accepts an object
#       containing the parameters

function Get-CWKeys
{
    [CmdLetBinding()]
    param
    (
        [Parameter(Mandatory)][string]$ServerBaseURL,
        [Parameter(Mandatory)][string]$ImpersonationMember,
        [Parameter(Mandatory)][string]$CWAuthString,
        [Parameter(Mandatory)][string]$BaseURI
    )

    # Define the base URI for this function
    $FunctionBaseURI    = "$BaseURI/system/members/$ImpersonationMember/tokens"
    
    # Format the auth string required by REST API
    $CWAuthString = '{0}+{1}:{2}' -f $CWDetails.CompanyName, $CWDetails.Username, $CWDetails.Password

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
        [Parameter(Mandatory)][string]$Company,
        [Parameter(Mandatory)][string]$ServerBaseURL,
        [Parameter(Mandatory)][string]$CompanyName,
        [Parameter(Mandatory)][string]$BaseURI,
        [Parameter(Mandatory)][psobject]$CWAuth
    )

    # Define the base URI for this function
    $FunctionBaseURI = "$BaseURI/company/contacts"

    if (!($CWAuth)) { 
        Throw 'Invalid CWCredentials defined.'
        $CWAuth
    }

    # Format the auth string required by REST API
    $CWAuthString = "$($CWDetails.CompanyName)+" + $($CWAuth.publickey) + ':' + $($CWAuth.privatekey)

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

#endregion FUNCTIONS

#region SCRIPT BODY

#
# Authenticate with ConnectWise!
#

# Get our ConnectWise authentication token
#   Create a function that has our command & arguments (DRY!)
Function AuthCommand {
    Write-Host 'Getting a new authentication token'
    Return (Get-CWKeys `
                -ServerBaseUrl $CWDetails.ServerBaseUrl `
                -ImpersonationMember $CWDetails.ImpersonationMember `
                -CWAuthString $CWAuthString `
                -BaseURI $BaseURI)
}

# Now, if we don't have an auth token, then get one! 
#   This is just for use in the ISE really as we don't need to get a new
#   8 hour authentication token every time we re-run the script.
If (!($CWauth)) {
    Write-Host 'No authentication token exists, requesting one.' -ForegroundColor Gray
    $global:CWAuth = AuthCommand
} else {
    Write-Host 'Authentication token already exists, using that.' -ForegroundColor Gray
}                        



#
# Find ConnectWise contacts
#

# Now, we find our contacts
#   Create a function that has our command & arguments (DRY!)
Function GetContacts {
    Return Get-CWContact -Company '*roost*' `
             -ServerBaseURL $CWDetails.ServerBaseURL `
             -CompanyName $CWDetails.CompanyName `
             -CWAuth $CWAuth `
             -BaseURI $BaseURI
}

Try {
    
    $CWContacts = GetContacts

    # If we have contacts returns, then...
    if ($CWContacts) {
        # Print them all for the time being
        $CWContacts

       Write-Host "$($CWContacts.Count) results returned"
    } else {
    # No contacts returned, write a warning (Probably a bad 
        Write-Warning 'No contacts returned, please check the Company Name.'
    }

} catch {
    # NOTE: Exception is $_
    
    if ($_.Exception.Message -contains '(401) Unauthorized.') {
        Write-Warning 'Getting "Unauthorized" error from ConnectWise API - Has auth token expired?'

        # Let's try authenticating again
        $Global:CWAuth = AuthCommand

        $CWContacts = GetContacts
    }
}

#endregion SCRIPT BODY