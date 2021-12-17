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
        2. Validates the following in OrganizationRelationship:
            a. Has a relationship with Source tenant
            b. The move direction is correct.
        3. Validates the following on Migration Endpoint:
            a. ApplicationId is correct.
            b. RemoteTenantId is correct.

   .PARAMETER PartnerTenantId
   PartnerTenantId - the tenant id of the partner tenant.
   
   .PARAMETER PartnerTenantDomain
   PartnerTenantDomain - the tenant domain of the partner tenant.

   .PARAMETER ApplicationId
   ApplicationId - the application setup for mailbox migration.

   .EXAMPLE - TargetTenant
   $report = VerifySetup.ps1 -PartnerTenantId <SourceTenantId> -ApplicationId <AADApplicationId>  -PartnerTenantDomain <PartnerTenantDomain> -Verbose

   .EXAMPLE - TargetTenant
   $report = VerifySetup.ps1 -PartnerTenantId <SourceTenantId> -ApplicationId <AADApplicationId>  -PartnerTenantDomain <PartnerTenantDomain> -Verbose

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
[string]$PartnerTenantDomain
)

$ErrorActionPreference = 'Stop'

$MS_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
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

    Write-Verbose "Verifying Application; AppId: [$ApplicationId] Current tenant: [$currentTenantId] Partner tenant: [$PartnerTenantId] IsTargetTenant: [$isTargetTenant]"
    $errors, $warnings = Verify-Application $ApplicationId $currentTenantId $PartnerTenantId $isTargetTenant
    $report["Application"] = @{ "Errors" = $errors; "Warnings" = $warnings }
    Write-Host "`r`n"
    Print-Result "Verifying AAD Application" $errors $warnings
    
    Write-Verbose "Verifying OrganizationRelationship; AppId: [$ApplicationId] Partner tenant: [$PartnerTenantId] IsTargetTenant: [$isTargetTenant]"
    $errors = Verify-OrganizationRelationship $PartnerTenantId $ApplicationId $isTargetTenant
    Print-Result "Verifying OrganizationRelationship" $errors
    $report["OrganizationRelationship"] = @{ "Errors" = $errors }
    
    if ($isTargetTenant -eq $true) {
        Write-Verbose "Verifying MigrationEndpoint; AppId: [$ApplicationId] Partner tenant: [$PartnerTenantDomain]"
        $errors = Verify-MigrationEndpoint $PartnerTenantDomain $ApplicationId
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
    
    $errors, $warnings
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

function Verify-MigrationEndpoint([string]$partnerTenantDomain, [string]$appId) {
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

Main