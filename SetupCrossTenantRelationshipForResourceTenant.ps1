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
    This script can be used by a tenant that wishes to pull resources out of another tenant.
    For example contoso.com would run this script in order to pull mailboxes from fabrikam.com tenant.

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
   SetupCrossTenantRelationshipForResourceTenant.ps1 -ResourceTenantDomain contoso.onmicrosoft.com -TargetTenantDomain fabrikam.onmicrosoft.com -TargetTenantId d925e0c6-d4db-40c6-a864-49db24af0460 -SourceMailboxMovePublishedScopes "SecurityGroupName"
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
function UserPrompt() {
    Write-Host "Welcome to the Cross-tenant mailbox migration preview! Before running this script, please be sure to review the details provided on docs.microsoft.com at https://docs.microsoft.com/en-us/microsoft-365/enterprise/cross-tenant-mailbox-migration."
    Write-Host "It is also recommended before running this script to review the script in a script editor or Notepad prior to running."`n
    Write-Host "For general feedback and / or questions, please contact mailto:crosstenantmigrationpreview@service.microsoft.com. This is not a support alias and should not be used if you are currently experiencing an issue and need immediate assistance."`n
    $title = "If you are ready to begin configuring your tenants, select 'Y'. If you need to review any additional details and proceed at a later time, select 'N'."
    $message = "Do you wish to proceed?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $choice=$host.ui.PromptForChoice($title, $message, $options, 1)
    if ($choice -eq 0) {
        Main}
    else {Exit}
}

UserPrompt
