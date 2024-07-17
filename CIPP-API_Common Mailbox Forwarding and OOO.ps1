#$debug = $true

### Authentication
$CIPPAPIUrl = Read-host -Prompt "Enter CIPP API URL"
$ApplicationId = Read-host -Prompt "Enter app ID"
$ApplicationSecret = Read-Host -Prompt "Enter app secret" -MaskInput
$TenantId = Read-Host -Prompt "Enter tenant ID"

$AuthBody = @{
    client_id     = $ApplicationId
    client_secret = $ApplicationSecret
    scope         = "api://$($ApplicationId)/.default"
    grant_type    = 'client_credentials'
}
$token = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $AuthBody

$AuthHeader = @{ Authorization = "Bearer $($token.access_token)" }

### Begin Script

$commonName = Read-Host -Prompt "Enter mailbox common name. Example: {tenantuser}@contoso.org"
$forwardToMailbox = Read-Host "Enter forward to mailbox. Example: {mytenants}@example.com"
$forwardToDomain = Read-Host "Enter forward to domain. Example: mytenants@{example.com}"
$oooMessage = Read-Host "Enter OOO Message"

#Get all accounts with commonName
$accounts = Invoke-RestMethod -Uri "$CIPPAPIUrl/api/listusers?graphfilter=startswith(userPrincipalName,$commonName)&TenantFilter=AllTenants" -Method GET -Headers $AuthHeader -ContentType "application/json"

foreach ($account in $accounts){

    ### Mailbox Forwarding
    #Get plus alias destination from the tenant name.
    $tenantShortName = $account.Tenant.Substring(0,$account.Tenant.IndexOf("."))

    #Create request body
    $forwardBody = @{
        tenantfilter = $account.Tenant
        userid = $account.userPrincipalName
        disableforwarding = 'false'
        keepcopy = 'true'
        forwardexternal = $forwardToMailbox+'+'+$tenantShortName+'@'+$forwardToDomain
    }

    #Convert to JSON
    $forwardBody = (ConvertTo-Json -InputObject $forwardBody)

    #Execute forwarding
    Invoke-RestMethod -Uri "$CIPPAPIUrl/api/execemailforward" -Method POST -Body $forwardBody -Headers $AuthHeader -ContentType "application/json"

    if($debug){pause}

    ### Mailbox OOO
    #Create request body
    $oooBody = @{
        tenantfilter = $account.Tenant
        user = $account.userPrincipalName
        autoreplystate = 'enabled'
        internalmessage = $oooMessage
    }

    #Convert to JSON
    $oooBody = (ConvertTo-Json -InputObject $oooBody)

    #Execute OOO
    Invoke-RestMethod -Uri "$CIPPAPIUrl/api/execsetooo" -Method POST -Body $oooBody -Headers $AuthHeader -ContentType "application/json"

    if($debug){pause}

}