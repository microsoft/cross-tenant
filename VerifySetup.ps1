#################################################################################
#
# The sample scripts are not supported under any Microsoft standard support
# program or service. The sample scripts are provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including, without
# limitation, any implied warranties of merchantability or of fitness for a particular
# purpose. The entire risk arising out of the use or performance of the sample scripts
# and documentation remains with you. In no event shall Microsoft, its authors, or
# anyone else involved in the creation, production, or delivery of the scripts be liable
# for any damages whatsoever (including, without limitation, damages for loss of business
# profits, business interruption, loss of business information, or other pecuniary loss)
# arising out of the use of or inability to use the sample scripts or documentation,
# even if Microsoft has been advised of the possibility of such damages.
#
#################################################################################

<# .SYNOPSIS
    This script can be used by a tenant that wishes to validate the setup required for cross-tenant mailbox migration.

    This script performs the following checks when run with -Context Target parameter:
        1. Validates the following on the AAD application:
            a. Is registered in the target tenant directory
            b. Is setup with right permissions on MSGraph and Exchange
            c. Is consented by an administrator
        2. Validates the following in KeyVault:
            a. The KeyVault url is correct
            b. Exchange first party application has READ permissions on the secret
        3. Validates the following in OrganizationRelationship:
            a. Has a relationship with Source tenant
            b. The move direction is correct.
        4. Validates the following on Migration Endpoint:
            a. ApplicationId is correct.
            b. ApplicationKeyVaultUrl is correct.
            c. RemoteTenantId is correct.

   .PARAMETER PartnerTenantId
   PartnerTenantId - the tenant id of the partner tenant.
   
   .PARAMETER PartnerTenantDomain
   PartnerTenantDomain - the tenant domain of the partner tenant.

   .PARAMETER ApplicationId
   ApplicationId - the application setup for mailbox migration.

   .PARAMETER ApplicationKeyVaultUrl
   ApplicationKeyVaultUrl - the keyvault url for application secret.

   .EXAMPLE - TargetTenant
   $report = VerifySetup.ps1 -PartnerTenantId <SourceTenantId> -ApplicationId <AADApplicationId> -ApplicationKeyVaultUrl <appKeyVaultUrl> -PartnerTenantDomain <PartnerTenantDomain> -Verbose

   .EXAMPLE - TargetTenant
   $report = VerifySetup.ps1 -PartnerTenantId <SourceTenantId> -ApplicationId <AADApplicationId> -ApplicationKeyVaultUrl <appKeyVaultUrl> -PartnerTenantDomain <PartnerTenantDomain> -SubscriptionId <AzureSubscriptionId> -Verbose

   .EXAMPLE - SourceTenant
   $report = VerifySetup.ps1 -PartnerTenantId <TargetTenantId> -ApplicationId <AADApplicationId>
#>

param
(
[Parameter(Mandatory = $true, HelpMessage='Partner tenant id', ParameterSetName = 'VerifyTarget')]
[Parameter(Mandatory = $true, HelpMessage='Partner tenant id', ParameterSetName = 'VerifySource')]
[ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
[string]$PartnerTenantId,



[Parameter(Mandatory = $true, HelpMessage='AAD ApplicationId', ParameterSetName = 'VerifyTarget')]
[Parameter(Mandatory = $true, HelpMessage='Partner tenant id', ParameterSetName = 'VerifySource')]
[ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
[string]$ApplicationId,



[Parameter(Mandatory = $true, HelpMessage='PartnerTenantDomain', ParameterSetName = 'VerifyTarget')]
[ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
[string]$PartnerTenantDomain,



[Parameter(Mandatory = $true, HelpMessage='App secret key vault url', ParameterSetName = 'VerifyTarget')]
[ValidateScript({
    if ($_ -cmatch "^https://[a-zA-Z_0-9\-]+\.vault\.azure.net(:443){0,1}/certificates/[a-zA-Z_0-9\-]+/[a-zA-Z_0-9]+$")
    {
        $true
    }
    elseif ($_ -cmatch "^https://[a-zA-Z_0-9]+\.vault\.azure.us(:443){0,1}/certificates/[a-zA-Z_0-9]+/[a-zA-Z_0-9]+$") {
        $true
    }
    else
    {
    throw [System.Management.Automation.ValidationMetadataException] "Please make sure key vault url matches format specified here: https://docs.microsoft.com/en-us/azure/key-vault/general/about-keys-secrets-certificates#vault-name-and-object-name"
    }
})]
[string]$ApplicationKeyVaultUrl,

[Parameter(Mandatory = $false, HelpMessage='SubscriptionId for key vault', ParameterSetName = 'VerifyTarget')]
[Parameter(Mandatory = $false, HelpMessage='SubscriptionId for key vault', ParameterSetName = 'VerifySource')]
[ValidateScript({-not [string]::IsNullOrWhiteSpace($_)})]
[string]$SubscriptionId
)

$ErrorActionPreference = 'Stop'

$MS_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
$MS_GRAPH_APP_ROLE = "User.Invite.All"
$EXO_APP_ID = "00000002-0000-0ff1-ce00-000000000000"
$EXO_APP_ROLE = "Mailbox.Migration"

function Main() {
    $report = @{}
    Check-ExchangeOnlinePowershellConnection
    $isTargetTenant = $PSCmdlet.ParameterSetName -eq 'VerifyTarget'
    $azureADAccount = Connect-AzureAD
    Write-Verbose "Connected to AzureAD - $($azureADAccount | Out-String)"
    $azAccount = Connect-AzAccount -Tenant $azureADAccount.Tenant.ToString()
    Write-Verbose "Connected to Az Account - $($azAccount | Out-String)"
    $currentTenantId = $azureADAccount.TenantId.Guid
    if($isTargetTenant -eq $true)
    {
        Check-AzSubscription
    }
    Write-Verbose "Verifying Application; AppId: [$ApplicationId] Current tenant: [$currentTenantId] Partner tenant: [$PartnerTenantId] IsTargetTenant: [$isTargetTenant]"
    $errors, $warnings = Verify-Application $ApplicationId $currentTenantId $PartnerTenantId $isTargetTenant
    $report["Application"] = @{ "Errors" = $errors; "Warnings" = $warnings }
    Write-Host "`r`n"
    Print-Result "Verifying AAD Application" $errors $warnings
    if ($isTargetTenant -eq $true) {
        Write-Verbose "Verifying KeyVault; AppId: [$ApplicationId] ApplicationKeyVaultUrl: [$ApplicationKeyVaultUrl]"
        $errors = Verify-KeyVault $ApplicationId $ApplicationKeyVaultUrl
        Print-Result "Verifying KeyVault" $errors
        $report["KeyVault"] = @{ "Errors" = $errors }
    }
    
    Write-Verbose "Verifying OrganizationRelationship; AppId: [$ApplicationId] Partner tenant: [$PartnerTenantId] IsTargetTenant: [$isTargetTenant]"
    $errors = Verify-OrganizationRelationship $PartnerTenantId $ApplicationId $isTargetTenant
    Print-Result "Verifying OrganizationRelationship" $errors
    $report["OrganizationRelationship"] = @{ "Errors" = $errors }
    
    if ($isTargetTenant -eq $true) {
        Write-Verbose "Verifying MigrationEndpoint; AppId: [$ApplicationId] Partner tenant: [$PartnerTenantDomain] ApplicationKeyVaultUrl: [$ApplicationKeyVaultUrl]"
        $errors = Verify-MigrationEndpoint $PartnerTenantDomain $ApplicationId $ApplicationKeyVaultUrl
        Print-Result "Verifying MigrationEndpoint" $errors 
        $report["MigrationEndpoint"] = @{ "Errors" = $errors }
    }
    
    Write-Verbose ($report | ConvertTo-Json)
    $report
}

function Print-Result([string]$opName, $errors, $warnings) {
    Write-Host "[$opName].............." -NoNewLine
    if (!$errors -and !$warnings) {
        Write-Host "[Passed]" -NoNewLine -ForeGroundColor Green
        Write-Host "`r`n"
        return
    } 

    if ($errors) {
        Write-Host "[Failed]" -ForeGroundColor Red
        Write-Host ($errors -join "`n") -ForeGroundColor Red
    }
        
    if ($warnings) {
        if (!$errors) {
            Write-Host "[Warnings]" -ForeGroundColor Yellow
        }
        
        Write-Host ($warnings -join "`n") -ForeGroundColor Yellow
    }
    
    Write-Host "`r`n`r`n"
}

function Verify-Application ([string]$appId, [string]$currentTenantId, [string]$partnerTenantId, [bool]$isTargetTenant) {
    $warnings = @()
    $errors = @()
    $spn = Get-AzureADServicePrincipal -Filter "AppId eq '$appId'"
    if (!$spn) {
        $errors += "App [$appId] is not registered in [$currentTenantId] tenant"
        return $errors
    }
    
    if ($isTargetTenant -eq $true) {
        if ($spn.AppOwnerTenantId -ne $currentTenantId) {
            $error += "App [$appId] was found in the [$currentTenantId] tenant but is not owned by it. Since this is target tenant, the app used for migration must be owned by target tenant."
        }
    } elseif ($spn.AppOwnerTenantId -ne $partnerTenantId) {
        $error += "App [$appId] was found in the [$currentTenantId] tenant but is not owned by $partnerTenantId. Please use an application owned by target tenant for mailbox migrations."
    }
    
    # Check MSGraph and EXO has incoming app roles assignment from the tenant friending app
    # 1. collect spns of MSGraph and EXO applications
    $msGraphSpn = Get-AzureADServicePrincipal -Filter "AppId eq '$MS_GRAPH_APP_ID'"
    $exoSpn = Get-AzureADServicePrincipal -Filter "AppId eq '$EXO_APP_ID'"
    if (!$msGraphSpn -or !$exoSpn) {
        $errors += "Internal Error: SPNs of MSGraph or EXO not found."
        return $errors
    }
    
    # Get the permission objects from Exo and MsGraph
    $exoMailboxMigrationPermissions = $exoSpn.AppRoles | ? { $_.Value -eq $EXO_APP_ROLE }
    $msGraphDirectoryPermissions = $msGraphSpn.AppRoles | ? { $_.Value -eq $MS_GRAPH_APP_ROLE }
    
    # Get the permission objects of the permissions assigned to the application
    $exoPermissionForApp = Get-AzureADServiceAppRoleAssignment -ObjectId $exoSpn.ObjectId -All $true | ? { $_.PrincipalId -eq $spn.ObjectId }
    $msGraphPermissionForApp = Get-AzureADServiceAppRoleAssignment -ObjectId $msGraphSpn.ObjectId -All $true | ? { $_.PrincipalId -eq $spn.ObjectId }
    
    if (!$exoPermissionForApp -or ($exoPermissionForApp.Id -ne $exoMailboxMigrationPermissions.Id)) {
        $errors += "App [$appId] does not have [$EXO_APP_ROLE] permission on Exchange setup or the permission is not consented by an Administrator"
    }
    
    if (!$msGraphPermissionForApp -or ($msGraphPermissionForApp.Id -ne $msGraphDirectoryPermissions.Id)) {
        $warnings += "App [$appId] does not have [$MS_GRAPH_APP_ROLE] permission on MSGraph setup or the permission is not consented by an Administrator"
    }
    
    $errors, $warnings
}

function Verify-KeyVault([string]$appId, [string]$appKvUrl) {
    $errors = @()
    try {
        $uri = [System.Uri]::new($appKvUrl)
        $kvName = $uri.Host.Split(".")[0]
        $kv = Get-AzKeyVault -VaultName $kvName
        if (!$kv) {
            $errors += "KeyVault: $kvName not found"
            return $errors
        } 
             
        $exoSpn = Get-AzureADServicePrincipal -Filter "AppId eq '$EXO_APP_ID'"
        $exoAccessPolicy = $kv.AccessPolicies | ? { $_.ObjectId -eq $exoSpn.ObjectId }
        if (!$exoAccessPolicy) {
            $errors += "Exchange does not have any permissions on the KeyVault [$kvName]"
            return $errors
        }        
        
        $certStorePermissions = $exoAccessPolicy.PermissionsToCertificates.ToLower()
        $secretStorePermissions = $exoAccessPolicy.PermissionsToSecrets.ToLower()
        "get", "list" | % { if (!$certStorePermissions.Contains($_)) {$errors += "Exchange does not have [$_] permission on KeyVault [$kvName]'s Certificate container"}}
        "get", "list" | % { if (!$secretStorePermissions.Contains($_)) {$errors += "Exchange does not have [$_] permission on KeyVault [$kvName]'s Secrets container"}}
    } catch {
        $errors += $_.Message
    }
    
    $errors
}

function Verify-OrganizationRelationship([string]$partnerTenantId, [string]$appId, [bool]$isTargetTenant) {
    $errors = @()
    $orgRel = Get-OrganizationRelationship | ? { $_.DomainNames -contains $partnerTenantId }
    if (!$orgRel) {
        $errors += "Organization relationship does not exist with [$partnerTenantId]"
        return $errors
    }
    
    if ($isTargetTenant -eq $true) {
        if (!$orgRel.MailboxMoveEnabled) {
            $errors += "MailboxMove is not enabled in Organization relationship with [$partnerTenantId]"
        }
        
        if ($orgRel.MailboxMoveCapability -ne 'Inbound') {
            $errors += "MailboxMoveCapability is invalid in Organization relationship with [$partnerTenantId]. It should be [Inbound] found [$($orgRel.MailboxMoveCapability)]"
        }
        
        if ($errors) {
            return $errors
        }
    } else {
        if (!$orgRel.MailboxMoveEnabled) {
            $errors += "MailboxMove is not enabled in Organization relationship with [$partnerTenantId]"
        }
        
        if ($orgRel.MailboxMoveCapability -ne 'RemoteOutbound') {
            $errors += "MailboxMoveCapability is invalid in Organization relationship with [$partnerTenantId]. It should be [RemoteOutbound] found [$($orgRel.MailboxMoveCapability)]"
        }
        
        if ($orgRel.OAuthApplicationId -ne $appId) {
            $errors += "Mailbox Migration ApplicationId is not whitelisted in the Organization Relationship with [$partnerTenantId]. Expected [$appId] found [$($orgRel.ApplicationId)]"
        }
        
        if (!$orgRel.MailboxMovePublishedScopes) {
            $errors += "Source tenant needs to specify MailboxMovePublishedScopes to allow migration"
        }
    }
    
    return $errors
}

function Verify-MigrationEndpoint([string]$partnerTenantDomain, [string]$appId, [string]$appKvUrl) {
    $errors = @()
    $migEp = Get-MigrationEndpoint | ? { $_.ApplicationId -eq $appId }
    if (!$migEp) {
        $errors += "Migration Endpoint containing [$appId] not found."
        return $errors
    }
    
    if ($migEp.RemoteTenant -ne $partnerTenantDomain) {
        $errors += "RemoteTenant does not match in Migration Endpoint. Expected [$partnerTenantDomain] found [$($migEp.RemoteTenant)]"
    }
    
    if ($migEp.ApplicationId -ne $appId) {
        $errors += "ApplicationId does not match in Migration Endpoint. Expected [$appId] found [$($migEp.ApplicationId)]"
    }
    
    if ($migEp.AppSecretKeyVaultUrl -ne $appKvUrl) {
        $errors += "AppSecretKeyVaultUrl does not match in Migration Endpoint. Expected [$appKvUrl] found [$($migEp.AppSecretKeyVaultUrl)]"
    }
    
    if (!$migEp.IsRemote) {
        $errors += "IsRemote does not match in Migration Endpoint. Expected [true] found [$($migEp.IsRemote)]"
    }
    
    return $errors
}

function Check-ExchangeOnlinePowershellConnection {
    if ($Null -eq (Get-Command New-OrganizationRelationship -ErrorAction SilentlyContinue)) {
        Write-Error "Please connect to the Exchange Online Management module or Exchange Online through basic authentication before running this script!";
    }
}

function Check-AzSubscription {
    if (!$SubscriptionId)
    {
        $subscriptions = Get-AzSubscription

        if ($subscriptions.Count -gt 1) {
            Write-Error "Multipule Azure subscriptions were found for this tenant. Please rerun the script and use the -SubscriptionId parameter with the correct subscription"
        }
        elseif (!$subscriptions) {
            Write-Error "No valid Azure subscriptions were found for this tenant."
        }

        Set-AzContext -Subscription $subscriptions.SubscriptionId
        }
    elseif ($SubscriptionID) 
    {
    Write-Verbose "SubscriptionId - $SubscriptionId was provided."
    Set-AzContext -Subscription $SubscriptionId
    }
}

Main