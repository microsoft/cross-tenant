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
    This script can be used by a tenant that wishes to move resources out of their tenant.
    For example fabrikam.com would run this script in order for the contoso.com tenant to pull mailboxes from the fabrikam.com tenant.

    This script is intended for the resource tenant in above example fabrikam.com, and it sets up the organization relationship in exchange to authorize the migration.
    Following are key properties in organization relationship used here:
    - ApplicationId of the azure ad application that resource tenant consents to for mailbox migrations.
    - SourceMailboxMovePublishedScopes contains the groups of users that are in scope for migration. Without this no mailboxes can be migrated.


   .PARAMETER SourceMailboxMovePublishedScopes
   SourceMailboxMovePublishedScopes - Identity of the scope used by source tenant admin.

   .PARAMETER ResourceTenantDomain
   ResourceTenantDomain - the resource tenant.

   .PARAMETER TargetTenantDomain
   TargetTenantDomain - The target tenant.

   .PARAMETER TargetTenantId
   TargetTenantId - The target tenant id.

   .EXAMPLE
   SetupCrossTenantRelationshipForResourceTenant.ps1 -ResourceTenantDomain fabrikam.onmicrosoft.com -TargetTenantDomain contoso.onmicrosoft.com -TargetTenantId d925e0c6-d4db-40c6-a864-49db24af0460 -SourceMailboxMovePublishedScopes "SecurityGroupName"
#>

[CmdletBinding(SupportsShouldProcess)]
param
(
    [Parameter(Mandatory = $true, HelpMessage='Setup Options')]
    [string[]]$SourceMailboxMovePublishedScopes,

    [Parameter(Mandatory = $true, HelpMessage='Resource tenant domain')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$ResourceTenantDomain,

    [Parameter(Mandatory = $true, HelpMessage='Target tenant domain')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $TargetTenantDomain,

    [Parameter(Mandatory = $true, HelpMessage='The application id for the azure ad application to be used for mailbox migrations')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $ApplicationId,

    [Parameter(Mandatory = $true, HelpMessage='Target tenant id. This is azure ad directory id or external directory object id in exchange online.')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $TargetTenantId
)

$ErrorActionPreference = 'Stop'


function Main() {
    Check-ExchangeOnlinePowershellConnection
    Run-ExchangeSetupForResourceTenant $TargetTenantDomain $TargetTenantId $ResourceTenantDomain $ApplicationId $SourceMailboxMovePublishedScopes
    Write-Host "Exchange setup complete." -Foreground Green
}

function Check-ExchangeOnlinePowershellConnection {
    if ($Null -eq (Get-Command New-OrganizationRelationship -ErrorAction SilentlyContinue)) {
        Write-Error "Please connect to the Exchange Online Management module or Exchange Online through basic authentication before running this script!";
    }
}

function Run-ExchangeSetupForResourceTenant([string]$targetTenant, [string]$targetTenantId, [string]$resourceTenantDomain, [string]$appId, [string[]]$sourceMailboxMovePublishedScopes) {
    # 1. Verify migration scope.
    # 2. Create organization relationship
    $orgRel = Get-OrganizationRelationship | ? { $_.DomainNames -contains $targetTenantId }

    if ($orgRel) {
        Write-Verbose "Organization relationship already exists with $targetTenantId. Updating it."
        $capabilities = @($orgRel.MailboxMoveCapability.Split(",").Trim())
        if (-not $orgRel.MailboxMoveCapability.Contains("RemoteOutbound")) {
            Write-Verbose "Adding RemoteOutbound capability to the organization relationship. Existing capabilities: $capabilities"
            $capabilities += "RemoteOutbound"
        }

        $orgRel | Set-OrganizationRelationship -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability $capabilities -OAuthApplicationId $appId -MailboxMovePublishedScopes $sourceMailboxMovePublishedScopes
    } else {
        $randomSuffix = [Random]::new().Next(0, 10000)
        $orgRelName = "$($targetTenant.Split('.')[0])_$($resourceTenantDomain.Split('.')[0])_$randomSuffix"
        $maxLength = [System.Math]::Min(64, $orgRelName.Length)
        $orgRelName = $orgRelName.SubString(0, $maxLength)

        Write-Verbose "Creating organization relationship: $orgRelName in $resourceTenantDomain"
        New-OrganizationRelationship `
            -DomainNames $targetTenantId `
            -Enabled:$true `
            -MailboxMoveEnabled:$true `
            -MailboxMoveCapability RemoteOutbound `
            -Name $orgRelName `
            -OAuthApplicationId $appId `
            -MailboxMovePublishedScopes $sourceMailboxMovePublishedScopes
    }
}

function PreValidation() {
    Write-Host `n
    Write-Host "Welcome to the Cross-tenant mailbox migration preview! Before running this script, please be sure to review the details provided on docs.microsoft.com at the following URL: `nhttps://docs.microsoft.com/en-us/microsoft-365/enterprise/cross-tenant-mailbox-migration"
    Write-Host "`nIt is also recommended before running this script to review the script in a script editor or Notepad prior to running."`n
    Write-Host "For general feedback and / or questions, please contact crosstenantmigrationpreview@service.microsoft.com.`nThis is not a support alias and should not be used if you are currently experiencing an issue and need immediate assistance."`n
    $title = "Confirm: Configure Cross-Tenant mailbox migration preview."
    $message = "`nIf you are ready to begin configuring your tenants, select 'Y'.`nIf you need to review any additional details and proceed at a later time, select 'N'.`n`nDo you wish to proceed?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $choice=$host.ui.PromptForChoice($title, $message, $options, 1)
    if ($choice -ne 0) {
        Exit}
    Start-Sleep 2
    Write-Host "`nWe are verifying that you are using the latest version of the script."`n
    Write-Host "This requires that we download the latest version of the script from GitHub to compare with your local copy."
    Write-Host "This file will be stored on your local computer temporarily, as well as overwrite your existing script file if it is out of date."
    $title = "Confirm: Allow for download from GitHub."
    $message = "`nIf you are ready to begin this step, select 'Y'. `nIf you would prefer to manually download the scripts to make sure you have the latest version, select 'N'"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $choice=$host.ui.PromptForChoice($title, $message, $options, 1)
    if ($choice -ne 0) {
        Exit}
    else {Verification}

}

function Verification {
    Write-Host "`nBeginning verification steps."`n
    Write-Host "Verifying ability to create a new organization relationship in the tenant."
    try {
        New-OrganizationRelationship -DomainNames contoso.onmicrosoft.com -Name Contoso -WhatIf -ErrorAction Stop
    }
    catch {
        Write-Output "You need to run the command Enable-OrganizationCustomization before continuing with execution of the script."
        Exit
    }
    Write-Host "`nVerifying that your script is up to date with the latest changes."
    Write-Host "`nBeginning download of SetupCrossTenantRelationshipForResourceTenant.ps1 and creation of temporary files."
    if ((Test-Path -Path $ScriptDir\XTenantTemp) -eq $true) {
        Remove-Item -Path $ScriptDir\XTenantTemp\ -Recurse -Force | Out-Null
    }
    New-Item -Path . -Name XTenantTemp -ItemType Directory | Out-Null
    Invoke-WebRequest -Uri https://github.com/microsoft/cross-tenant/releases/download/Preview/SetupCrossTenantRelationshipForResourceTenant.ps1 -Outfile $ScriptDir\XTenantTemp\SetupCrossTenantRelationshipForResourceTenant.ps1
    if ((Get-FileHash $ScriptDir\SetupCrossTenantRelationshipForResourceTenant.ps1).hash -eq (Get-FileHash $ScriptDir\XTenantTemp\SetupCrossTenantRelationshipForResourceTenant.ps1).hash) {
        Write-Host "`nYou are using the latest version of the script. Removing temporary files and proceeding with setup."
        Start-Sleep 1
        Remove-Item -Path $ScriptDir\XTenantTemp\ -Recurse -Force | Out-Null
    }
    elseif ((Get-FileHash $ScriptDir\SetupCrossTenantRelationshipForResourceTenant.ps1).hash -ne (Get-FileHash $ScriptDir\XTenantTemp\SetupCrossTenantRelationshipForResourceTenant.ps1).hash) {
        Write-Host "`nYou are not using the latest version of the script."`n
        Start-Sleep 1
        Write-Host "`nReplacing the local copy of SetupCrossTenantRelationshipForResourceTenant.ps1 and cleaning up temporary files..."
        Start-Sleep 1
        Copy-Item $ScriptDir\XTenantTemp\SetupCrossTenantRelationshipForResourceTenant.ps1 -Destination $ScriptDir | Out-Null
        Start-Sleep 1
        Remove-Item -Path $ScriptDir\XTenantTemp\ -Recurse -Force | Out-Null
        Write-Host "Update completed. You will need to run the script again."
        Start-Sleep 1
        Exit
    }
}

PreValidation
Main
