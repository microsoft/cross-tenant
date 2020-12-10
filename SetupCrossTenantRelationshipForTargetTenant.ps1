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

    This script is intended for the target tenant and would setup the following using the SubscriptionId specified or the default subscription:
        1. Create a resource group or use the one specified as parameter
        2. Create a key vault in the above resource group specified as a parameter
        3. Setup above key vault's access policy to grant exchange access to secrets and certificates.
        4. Request a self-signed certificate to be put in the key vault.
        5. Retrieve the public part of certificate from key vault
        6. Create an AAD application and setup its permissions for MSGraph and exchange
        7. Set the secret for above application as the certificate in 4.
        8. Wait for the tenant admin to consent to the application permissions
        9. Once confirmed, send an email using initiation manager to the tenant admin of resource tenant.
        10. Create a migration endpoint in exchange with the ApplicationId, Pointer to application secret in KeyVault and RemoteTenant
        11. Create an organization relationship with resource tenant authorizing migration.

   .PARAMETER SubscriptionId
   SubscriptionId - the subscription to use for key vault.

   .PARAMETER ResourceTenantAdminEmail
   ResourceTenantAdminEmail - the resource tenant admin email.

   .PARAMETER ResourceGroup
   ResourceGroup - the resource group name.

   .PARAMETER KeyVaultName
   KeyVaultName - the key vault name.

   .PARAMETER KeyVaultLocation
   KeyVaultLocation - the location of the key vault

   .PARAMETER CertificateName
   CertificateName - the name of certificate in key vault

   .PARAMETER CertificateSubject
   CertificateSubject - the subject of certificate in key vault

   .PARAMETER AzureAppPermissions
   AzureAppPermissions - fine grained control over the permissions to be given to the application.

   .PARAMETER UseAppAndCertGeneratedForSendingInvitation
   UseAppAndCertGeneratedForSendingInvitation - download the private key of generated certificate from key vault to be used for sending invitation.

   .PARAMETER ResourceTenantDomain
   ResourceTenantDomain - the resource tenant.

   .PARAMETER TargetTenantDomain
   TargetTenantDomain - The target tenant.

   .PARAMETER ResourceTenantId
   ResourceTenantId - The resource tenant id.

   .EXAMPLE
   SetupCrossTenantRelationshipForTargetTenant.ps1 -ResourceTenantDomain fabrikam.onmicrosoft.com -TargetTenantDomain contoso.onmicrosoft.com -ResourceTenantAdminEmail admin@contoso.onmicrosoft.com -ResourceGroup "TESTPSRG" -KeyVaultName "TestPSKV" -CertificateSubject "CN=TESTCERTSUBJ" -AzureAppPermissions Exchange, MSGraph -UseAppAndCertGeneratedForSendingInvitation -KeyVaultAuditStorageAccountName "KeyVaultLogsStorageAcnt" -ExistingApplicationId d7404497-1e2f-4b58-bdd5-93e82dad91a4

   .EXAMPLE
   SetupCrossTenantRelationshipForTargetTenant.ps1 -ResourceTenantDomain fabrikam.onmicrosoft.com -TargetTenantDomain contoso.onmicrosoft.com -ResourceTenantId <ContosoTenantId>
#>

[CmdletBinding(SupportsShouldProcess)]
param
(
    [Parameter(Mandatory = $true, HelpMessage='SubscriptionId for key vault', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $true, HelpMessage='SubscriptionId for key vault', ParameterSetName = 'TargetSetupAzure')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $true, HelpMessage='Resource tenant admin email', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $true, HelpMessage='Resource tenant admin email', ParameterSetName = 'TargetSetupAzure')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$ResourceTenantAdminEmail,

    [Parameter(Mandatory = $true, HelpMessage='Resource group for key vault', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $true, HelpMessage='Resource group for key vault', ParameterSetName = 'TargetSetupAzure')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$ResourceGroup,

    [Parameter(Mandatory = $true, HelpMessage='KeyVault name', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $true, HelpMessage='KeyVault name', ParameterSetName = 'TargetSetupAzure')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$KeyVaultName,

    [Parameter(HelpMessage='KeyVault location', ParameterSetName = 'TargetSetupAll')]
    [Parameter(HelpMessage='KeyVault location', ParameterSetName = 'TargetSetupAzure')]
    [string]$KeyVaultLocation = "West US",

    [Parameter(Mandatory = $false, HelpMessage='Resource group for storage account used for key vault audit logs', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $false, HelpMessage='Resource group for storage account used for key vault audit logs', ParameterSetName = 'TargetSetupAzure')]
    [string]$KeyVaultAuditStorageResourceGroup,

    [Parameter(Mandatory = $false, HelpMessage='Storage account name for storing key vault audit logs', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $false, HelpMessage='Storage account name for storing key vault audit logs', ParameterSetName = 'TargetSetupAzure')]
    [string]$KeyVaultAuditStorageAccountName,

    [Parameter(HelpMessage='Certificate name to use', ParameterSetName = 'TargetSetupAll')]
    [Parameter(HelpMessage='Certificate name to use', ParameterSetName = 'TargetSetupAzure')]
    [string]$CertificateName,

    [Parameter(HelpMessage='Certificate subject to use', ParameterSetName = 'TargetSetupAll')]
    [Parameter(HelpMessage='Certificate subject to use', ParameterSetName = 'TargetSetupAzure')]
    [ValidateScript({$_.StartsWith("CN=") })]
    [string]$CertificateSubject,

    [Parameter(HelpMessage='Application permissions', ParameterSetName = 'TargetSetupAll')]
    [Parameter(HelpMessage='Application permissions', ParameterSetName = 'TargetSetupAzure')]
    $AzureAppPermissions = 'All',

    [Parameter(HelpMessage='Use the certificate generated for azure application when sending invitation', ParameterSetName = 'TargetSetupAll')]
    [Parameter(HelpMessage='Use the certificate generated for azure application when sending invitation', ParameterSetName = 'TargetSetupAzure')]
    [Switch]$UseAppAndCertGeneratedForSendingInvitation,

    [Parameter(Mandatory = $true, HelpMessage='Resource tenant domain')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    [string]$ResourceTenantDomain,

    [Parameter(Mandatory = $true, HelpMessage='Target tenant domain')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $TargetTenantDomain,

    [Parameter(Mandatory = $true, HelpMessage='Target tenant id. This is azure ad directory id or external directory object id in exchange online.', ParameterSetName = 'TargetSetupAll')]
    [Parameter(Mandatory = $true, HelpMessage='Target tenant id. This is azure ad directory id or external directory object id in exchange online.', ParameterSetName = 'TargetSetupExchange')]
    [ValidateScript({ -not [string]::IsNullOrWhiteSpace($_) })]
    $ResourceTenantId,

    [Parameter(HelpMessage='Existing Application Id. If existing application Id is present and can be found, new application will not be created.', ParameterSetName = 'TargetSetupAll')]
    [Parameter(HelpMessage='Existing Application Id. If existing application Id is present and can be found, new application will not be created.', ParameterSetName = 'TargetSetupAzure')]
    [guid]$ExistingApplicationId  = [guid]::Empty
)

$ErrorActionPreference = 'Stop'

$MS_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
$MS_GRAPH_APP_ROLE = "Directory.ReadWrite.All"
$EXO_APP_ID = "00000002-0000-0ff1-ce00-000000000000"
$EXO_APP_ROLE = "Mailbox.Migration"
$REPLY_URL = "https://office.com"
$FIRSTPARTY_POWERSHELL_CLIENTID = "a0c73c16-a7e3-4564-9a95-2bdf47383716"
$FIRSTPARTY_POWERSHELL_CLIENT_REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob' -as [Uri]

function Main() {
    Check-ExchangeOnlinePowershellConnection

    $AzureAppPermissions = ([ApplicationPermissions]$AzureAppPermissions)
    if ($PSCmdlet.ParameterSetName -eq 'TargetSetupAll' -or $PSCmdlet.ParameterSetName -eq 'TargetSetupAzure') {
        Import-AzureModules
        if (-not $AzureAppPermissions.HasFlag([ApplicationPermissions]::MSGraph) -and $UseAppAndCertGeneratedForSendingInvitation) {
            Write-Error "Cannot use application for sending invitation as it does not have permissions on MSGraph"
        }

        $azureADAccount = Connect-AzureAD
        Write-Verbose "Connected to AzureAD - $($azureADAccount | Out-String)"
        $azAccount = Connect-AzAccount
        Write-Verbose "Connected to Az Account - $($azAccount | Out-String)"

        Write-Host "Setting up key vault in the $TargetTenantDomain tenant"
        $subscriptions = Get-AzSubscription
        Write-Verbose "SubscriptionId - $SubscriptionId was provided. Searching for it in $($subscriptions | Out-String)"
        $subscription = $subscriptions | ? { $_.SubscriptionId -eq $SubscriptionId}
        if (-not $subscription) {
            Write-Error "Subscription with id $SubscriptionId was not found."
        }

        Write-Verbose "Found subscription - $($SubscriptionId | Out-String)"
        Set-AzContext -Subscription $SubscriptionId

        ## Grab the EXO & MSGraph APP SPN ##
        $spns = @()
        $msGraphSpn=Get-AzureADServicePrincipal -Filter "AppId eq '$MS_GRAPH_APP_ID'"
        $exoAppSpn = Get-AzureADServicePrincipal -Filter "AppId eq '$EXO_APP_ID'"
        $spns += $msGraphSpn
        $spns += $exoAppSpn
        Write-Verbose "Found exchange service principal in $TargetTenantDomain - $($exoAppSpn | Out-String)"

        $certificatePublicKey, $certificatePrivateKey = Create-KeyVaultAndGenerateCertificate `
                                                            $TargetTenantDomain `
                                                            $ResourceTenantDomain `
                                                            $ResourceGroup `
                                                            $KeyVaultName `
                                                            $KeyVaultLocation `
                                                            $CertificateName `
                                                            $CertificateSubject `
                                                            $exoAppSpn.ObjectId `
                                                            $UseAppAndCertGeneratedForSendingInvitation `
                                                            $KeyVaultAuditStorageResourceGroup `
                                                            $KeyVaultAuditStorageAccountName `
                                                            $ExistingApplicationId

        Write-Verbose "Creating an application in $TargetTenantDomain"
        if (-not $AzureAppPermissions.HasFlag([ApplicationPermissions]::MSGraph)) {
            Write-Warning "MSGraph permission was not specified, however, an app needs at least one permission on ADGraph in order for admin to consent to it via the consent url. This app may only be consented from the azure portal."
        }

        $appOwnerTenantId, $appCreated = Create-Application $TargetTenantDomain $ResourceTenantDomain ($certificatePublicKey.Certificate) $spns ([ApplicationPermissions]$AzureAppPermissions) $ExistingApplicationId
        $global:AppId = $appCreated.AppId
        $appReplyUrl = $appCreated.ReplyUrls[0]
        $global:CertificateId = $certificatePublicKey.Id
        Write-Host "Application details to be registered in organization relationship: ApplicationId: [ $AppId ]. KeyVault secret Id: [ $CertificateId ]. These values are available in variables `$AppId and `$CertificateId respectively" -Foreground Green
        Write-Verbose "Sending the consent URI for this app to $ResourceTenantAdminEmail."
        Read-Host "Please consent to the app for $TargetTenantDomain before sending invitation to $ResourceTenantAdminEmail"
        Send-AdminConsentUri $TargetTenantDomain $ResourceTenantDomain $ResourceTenantAdminEmail $AppId $certificatePrivateKey $appReplyUrl $appCreated.DisplayName
    }

    if ($PSCmdlet.ParameterSetName -eq 'TargetSetupAll' -or $PSCmdlet.ParameterSetName -eq 'TargetSetupExchange') {
        $AppId = Ensure-VariableIsPopulated "AppId" "Please enter the application id for the azure ad application to be used for mailbox migrations"
        $CertificateId = Ensure-VariableIsPopulated "CertificateId" "Please enter the key vault url for the migration app's secret"
        Run-ExchangeSetupForTargetTenant $TargetTenantDomain $ResourceTenantDomain $ResourceTenantId $AppId $CertificateId
        Write-Host "Exchange setup complete. Migration endpoint details are available in `$MigrationEndpoint variable" -Foreground Green
    }
}

function Check-ExchangeOnlinePowershellConnection {
    if ($Null -eq (Get-Command New-OrganizationRelationship -ErrorAction SilentlyContinue)) {
        Write-Error "Please connect to the Exchange Online Management module or Exchange Online through basic authentication before running this script!";
    }
}

function Import-AzureModules() {
    $desiredAzureModules = @{
        "AzureAD"      = [Version]"2.0.2.4";
        "Az.Monitor"   = [Version]"1.2.0";
        "Az.KeyVault"  = [Version]"1.2.0";
        "Az.Accounts"  = [Version]"1.5.2";
        "Az.Resources" = [Version]"1.3.1";
    }

    $moduleMissingErrors = @()
    $desiredAzureModules.Keys | % {
        $desiredVersion = [Version]($desiredAzureModules[$_])
        $desiredAzModule = (Get-Module $_ -ListAvailable -Verbose:$false | ? { $_.Version -ge $desiredVersion})
        if (-not $desiredAzModule) {
            $moduleMissingErrors += "Powershell module: [$_] minimum version [$($desiredAzureModules[$_])] is required for running this script. Please install this module using: Install-Module $_ -AllowClobber"
        }
    }

    if ($moduleMissingErrors) {
        Write-Error "Missing modules - `r`n$([string]::Join("`r`n", $moduleMissingErrors))"
    }

    Import-Module AzureAD -Verbose:$false | Out-Null
    $desiredAzureModules.Keys | Import-Module -verbose:$false | Out-Null
}

function Ensure-VariableIsPopulated([string]$variableName, [string]$message) {
    $val = Get-Variable $variableName -ErrorAction Ignore
    if (-not $val) {
        $enteredVal = Read-Host $message
        if (-not $enteredVal) {
            Write-Error "Entered value was not valid"
        }

        $enteredVal
    }

    $val.Value
}

function Create-KeyVaultAndGenerateCertificate([string]$targetTenant, `
                                               [string]$resourceTenantDomain, `
                                               [string]$resourceGrpName, `
                                               [string]$kvName, `
                                               [string]$kvLocation, `
                                               [string]$certName, `
                                               [string]$certSubj, `
                                               [string]$exoAppObjectId, `
                                               $retrieveCertPrivateKey, `
                                               [string]$auditStorageAcntRG, `
                                               [string]$auditStorageAcntName, `
                                               [guid]$existingApplicationId) {
    if ([string]::IsNullOrWhiteSpace($certName)) {
        $randomPrefix = [Random]::new().Next(0, 10000)
        $certName = $randomPrefix.ToString() + "TenantFriendingAppSecret"
    }

    $resGrp = $null
    try {
        $resGrp = Get-AzResourceGroup -Name $resourceGrpName
        if ($resGrp) {
            Write-Verbose "Resource group $resourceGrpName already exists."
        }
    } catch {
        Write-Verbose "Resource group $resourceGrpName not found, this will be created."
    }

    if (-not $resGrp) {
        Write-Verbose "Creating resource group - $resourceGrpName"
        $resGrp = New-AzResourceGroup -Name $resourceGrpName -Location $kvLocation
        Write-Host "Resource Group $resourceGrpName successfully created" -Foreground Green
    }

    $kv = $null
    try {
        $kv = Get-AzKeyVault -VaultName $kvName -ResourceGroupName $resourceGrpName
    } catch {
        Write-Verbose "KeyVault $kvName not found, this will be created."
    }

    if ($kv) {
        Write-Verbose "Keyvault $kvName already exists."
    } else {
        Write-Verbose "Creating KeyVault $kvName"
        $kv = New-AzKeyVault -Name $kvName -Location $kvLocation -ResourceGroupName $resourceGrpName
        Write-Host "KeyVault $kvName successfully created" -Foreground Green
    }

    if ($auditStorageAcntRG -and $auditStorageAcntName) {
        Write-Verbose "Setting up auditing for key vault $kvName"
        $storageAcnt = Get-AzStorageAccount -ResourceGroupName $auditStorageAcntRG -Name $auditStorageAcntName
        Set-AzDiagnosticSetting -ResourceId $kv.ResourceId -StorageAccountId $storageAcnt.Id -Enabled $true -Category AuditEvent | Out-Null
        Write-Host "Auditing setup successfully for $kvName" -Foreground Green
    }

    Write-Verbose "Setting up access for key vault $kvName"
    Set-AzKeyVaultAccessPolicy -ResourceId $kv.ResourceId -ObjectId $exoAppObjectId -PermissionsToSecrets get,list -PermissionsToCertificates get,list | Out-Null
    Write-Host "Exchange app given access to KeyVault $kvName" -Foreground Green
    try {
        $cert = Get-AzKeyVaultCertificate -VaultName $kvName -Name $certName
        if ($cert.Certificate) {
            Write-Verbose "Certificate $certName already exists in $kvName"
            if ($retrieveCertPrivateKey -eq $true) {
                Write-Verbose "Retrieving certificate private key"
                $certPrivateKey = Get-AzKeyVaultSecret -VaultName $kvName -Name $certName
            }

            return $cert, $certPrivateKey
        }
    } catch {
        Write-Verbose "Certificate not found, a new request will be generated."
    }

    if ( [string]::IsNullOrWhiteSpace($certSubj)) {
        $certSubj = "CN=" + $targetTenant + "_" + $resourceTenantDomain + "_" + ([Random]::new().Next(0, 10000)).ToString()
        Write-Verbose "Cert subject not provided, generated subject - $certSubj"
    }

    $policy = New-AzKeyVaultCertificatePolicy -SubjectName $certSubj -IssuerName Self -ValidityInMonths 12
    $certReq = Add-AzKeyVaultCertificate -VaultName $kvName -Name $certName -CertificatePolicy $policy
    Write-Host "Self signed certificate requested in key vault - $kvName. Certificate name - $certName" -Foreground Green
    $tries = 5
    $certPrivateKey = $null
    while ($tries -gt 0) {
        try {
            Write-Verbose "Looking for certificate $certName. Attempt - $(6 - $tries)"
            $cert = Get-AzKeyVaultCertificate -VaultName $kvName -Name $certName
            if ($cert.Certificate) {
                Write-Verbose "Certificate found - $($cert | Out-String)"
                if ($retrieveCertPrivateKey -eq $true) {
                    $certPrivateKey = Get-AzKeyVaultSecret -VaultName $kvName -Name $certName
                    if ($certPrivateKey) {
                        Write-Verbose "Certificate private key also found"
                        break;
                    } else {
                        if ($tries -lt 0) {
                            Write-Error "Certificate private key not found after retries."
                        }

                        Write-Verbose "Certificate public key is present, however, its private key is not available, waiting 5 secs and looking again."
                    }
                }
            } else {
                if ($tries -lt 0) {
                    Write-Error "Certificate not found after retries."
                }

                Write-Verbose "Certificate not found, waiting 5 secs and looking again."
                sleep 5
            }
        } catch {
            if ($tries -lt 0) {
                Write-Error "Certificate not found after retries."
            }

            sleep 5
        }

        $tries--
    }

    Write-Verbose "Returning cert - $($cert.Certificate | Out-String)"
    Write-Host "Certificate $certName successfully created" -Foreground Green
    $cert, $certPrivateKey
}

function Create-Application([string]$targetTenantDomain, [string]$resourceTenantDomain, $certificate, $spns, $azAppPermissions, [guid]$ExistingApplicationId) {
    if ([guid]::Empty -ne $ExistingApplicationId) {
        $existingApp = Get-AzureADApplication -Filter "AppId eq '$ExistingApplicationId'"
        if ($Null -ne $existingApp) {
            Write-Warning "Existing application '$ExistingApplicationId' found. Skipping new application creation."
            return (Get-AzureADTenantDetail).ObjectId, $existingApp
        }
    }

    #### Collect all the permissions first ####
    $appPermissions = @()
    $msGraphSpn = $null

    if ($azAppPermissions.HasFlag([ApplicationPermissions]::MSGraph)) {
        ## Calculate permission on MSGraph ##
        $msGraphSpn = $spns | ? { $_.AppId -eq $MS_GRAPH_APP_ID }
        if (-not $msGraphSpn) {
            Write-Error "Tenant does not have access to MSGraph"
        }

        $msGraphAppPermission = $msGraphSpn.AppRoles | ? { $_.Value -eq $MS_GRAPH_APP_ROLE }
        $reqGraph = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
        $reqGraph.ResourceAppId = $msGraphSpn.AppId
        $reqGraph.ResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $msGraphAppPermission.Id,"Role"
        $appPermissions += $reqGraph
    }

    if ($azAppPermissions.HasFlag([ApplicationPermissions]::Exchange)) {
        ## Calculate permission on EXO ##
        $exoAppSpn = $spns | ? { $_.AppId -eq $EXO_APP_ID }
        if (-not $exoAppSpn) {
            Write-Error "Tenant does not have Exchange enabled"
        }

        $exoAppPermission = $exoAppSpn.AppRoles | ? { $_.Value -eq $EXO_APP_ROLE }
        $reqExo = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
        $reqExo.ResourceAppId = $exoAppSpn.AppId
        $reqExo.ResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $exoAppPermission.Id,"Role"
        $appPermissions += $reqExo
    }

    #### Create the app with all the permissions ####
    $appOwnerTenantId = (Get-AzureADTenantDetail).ObjectId
    $randomSuffix = [Random]::new().Next(0, 10000)
    $appName = "$($targetTenantDomain.Split('.')[0])_Friends_$($resourceTenantDomain.Split('.')[0])_$randomSuffix"
    $appCreationParameters = @{
        "AvailableToOtherTenants" = $true;
        "DisplayName" = $appName;
        "Homepage" = $REPLY_URL;
        "ReplyUrls" = $REPLY_URL;
        "RequiredResourceAccess" = $appPermissions
    }

    $appCreated = New-AzureADApplication @appCreationParameters

    $base64CertHash = [System.Convert]::ToBase64String($certificate.GetCertHash())
    $base64CertVal = [System.Convert]::ToBase64String($certificate.GetRawCertData())
    $appCertPwd = New-AzureADApplicationKeyCredential -ObjectId $appCreated.ObjectId -CustomKeyIdentifier $base64CertHash -Value $base64CertVal -StartDate ([DateTime]::Now) -EndDate ([DateTime]::Now).AddDays(363) -Type AsymmetricX509Cert -Usage Verify
    $spn = New-AzureADServicePrincipal -AppId $appCreated.AppId -AccountEnabled $true -DisplayName $appCreated.DisplayName
    $permissions = ""
    if ($azAppPermissions.HasFlag([ApplicationPermissions]::MSGraph)) {
        $permissions += "MSGraph - $MS_GRAPH_APP_ROLE. "
    }

    if ($azAppPermissions.HasFlag([ApplicationPermissions]::Exchange)) {
        $permissions += "Exchange - $EXO_APP_ROLE"
    }

    Write-Host "Application $appName created successfully in $targetTenantDomain tenant with following permissions. $permissions" -Foreground Green
    Write-Host "Admin consent URI for $targetTenantDomain tenant admin is -" -Foreground Yellow
    Write-Host ("https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}" -f $targetTenantDomain, $appCreated.AppId, $appCreated.ReplyUrls[0])

    Write-Host "Admin consent URI for $resourceTenantDomain tenant admin is -" -Foreground Yellow
    Write-Host ("https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}" -f $resourceTenantDomain, $appCreated.AppId, $appCreated.ReplyUrls[0])

    return $appOwnerTenantId, $appCreated
}

function Get-AppOnlyToken([string]$authContextTenant, [string]$appId, [string]$resourceUri, $appSecretCert) {
    $authority = "https://login.microsoftonline.com/$authContextTenant/oauth2/token"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $ssPtr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($appSecretCert.SecretValue)
    
    try {
    $secretValueText = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ssPtr)
    }finally {
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ssPtr)
    }
    
    $certBytes = [System.Convert]::FromBase64String($secretValueText)
    $clientCreds = new-object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate -ArgumentList $appId, ([System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes))
    Write-Verbose "Acquiring token resourceAppIdURI $resourceUri appSecret $appSecretCert"
    return $authContext.AcquireTokenAsync($resourceUri, $clientCreds).Result
}

function Get-AccessTokenWithUserPrompt([string]$authContextTenant, [string]$resourceUri) {
    $authority = "https://login.microsoftonline.com/common/oauth2/token"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    Write-Verbose "Acquiring token resourceAppIdURI $resourceUri"
    $platformParams = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList ([Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Always)
    return $authContext.AcquireTokenAsync($resourceUri, $FIRSTPARTY_POWERSHELL_CLIENTID, $FIRSTPARTY_POWERSHELL_CLIENT_REDIRECT_URI, $platformParams).GetAwaiter().GetResult()
    }

function Send-AdminConsentUri([string]$invitingTenant, [string]$resourceTenantDomain, [string]$resourceTenantDomainAdminEmail, [string]$appId, $appSecretCert, [string]$appReplyUrl, [string]$appName) {
    $authRes = $null
    $msGraphResourceUri = "https://graph.microsoft.com"
    Write-Verbose "Preparing invitation. Waiting for 10 secs before requesting token for the consented application to give time for replication."
    sleep 10
    if ($appSecretCert) {
        $authRes = Get-AppOnlyToken $invitingTenant $appId $msGraphResourceUri $appSecretCert
    } else {
        $authRes = Get-AccessTokenWithUserPrompt $invitingTenant $msGraphResourceUri $appId $appReplyUrl
    }

    if (-not $authRes) {
        Write-Error "Could not retrieve a token for invitation manager api call"
    }

    $invitationBody = @{
        invitedUserEmailAddress = $resourceTenantDomainAdminEmail
        inviteRedirectUrl = ("https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}" -f $resourceTenantDomain, $appId, $appReplyUrl)
        sendInvitationMessage = $true
        invitedUserMessageInfo = @{
            customizedMessageBody = "Organization [$invitingTenant] wishes to pull mailboxes from your organization using [$appName] application. `
            If you recognize this application please click below to provide your consent. `
            To authorize this application to be used for office 365 mailbox migration, please add its application id [$appId] to your organization relationship with [$invitingTenant] in the OAuthApplicationId property."
        }
    }

    $invitationBodyJson = $invitationBody | ConvertTo-Json
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $authRes.CreateAuthorizationHeader())
    Write-Verbose "Sending invitation"

    $resp = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/invitations" -Body $invitationBodyJson -ContentType 'application/json' -Headers $headers

    if ($resp -and $resp.invitedUserEmailAddress) {
        Write-Host "Successfully sent invitation to $($resp.invitedUserEmailAddress)" -Foreground Green
    }
}

function Run-ExchangeSetupForTargetTenant([string]$targetTenant, [string]$resourceTenantDomain, [string]$resourceTenantId, [string]$appId, [string]$appSecretKeyVaultUrl) {
    # 1. Create/Update organization relationship.
    # 2. Create migration endpoint.

    Write-Host "Setting up exchange components on target tenant: $targetTenant"
    if (-not (Get-Command Get-OrganizationRelationship -ErrorAction Ignore)) {
        Write-Error "We could not find exchange powershell cmdlet. Please re-establish the session and rerun this script."
    }

    $orgRel = Get-OrganizationRelationship | ? { $_.DomainNames -contains $resourceTenantId }
    if ($orgRel) {
        Write-Verbose "Organization relationship already exists with $resourceTenantId. Updating it."
        $capabilities = @($orgRel.MailboxMoveCapability.Split(",").Trim())
        if (-not $orgRel.MailboxMoveCapability.Contains("Inbound")) {
            Write-Verbose "Adding Inbound capability to the organization relationship. Existing capabilities: $capabilities"
            $capabilities += "Inbound"
        }

        $orgRel | Set-OrganizationRelationship -Enabled:$true -MailboxMoveEnabled:$true -MailboxMoveCapability $capabilities
        $orgRelName = $orgRel.Name
    } else {
        $randomSuffix = [Random]::new().Next(0, 10000)
        $orgRelName = "$($targetTenant.Split('.')[0])_$($resourceTenantDomain.Split('.')[0])_$randomSuffix"
        $maxLength = [System.Math]::Min(64, $orgRelName.Length)
        $orgRelName = $orgRelName.SubString(0, $maxLength)

        Write-Verbose "Creating organization relationship: $orgRelName in $targetTenant. DomainName: $resourceTenantId"
        New-OrganizationRelationship `
            -DomainNames $resourceTenantId `
            -Enabled:$true `
            -MailboxMoveEnabled:$true `
            -MailboxMoveCapability Inbound `
            -Name $orgRelName
    }

    Write-Verbose "Creating migration endpoint $orgRelName with remote tenant: $resourceTenantDomain, appId: $appId, appSecret: $appSecretKeyVaultUrl"
    $global:MigrationEndpoint = New-MigrationEndpoint `
                                    -Name $orgRelName `
                                    -RemoteTenant $resourceTenantDomain `
                                    -RemoteServer "outlook.office.com" `
                                    -ApplicationId $appId `
                                    -AppSecretKeyVaultUrl $appSecretKeyVaultUrl `
                                    -ExchangeRemoteMove:$true

    $MigrationEndpoint
    Write-Host "MigrationEndpoint created in $targetTenant for source $resourceTenantDomain" -Foreground Green
}


$enumExists = $null
try {
    $enumExists = [ApplicationPermissions] | Get-Member
} catch { }

if (-not $enumExists) {
    Add-Type -TypeDefinition @"
       using System;

       [Flags]
       public enum ApplicationPermissions
       {
          Exchange = 1,
          MSGraph = 2,
          All = Exchange | MSGraph
       }
"@
}

function UserPrompt() {
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
    else {Main}
}

UserPrompt

<#
set-OrganizationRelationship -Identity <tenant>\<id> -OAuthApplicationId 484a8384-979a-4cc9-8791-8e6bb34f76d4
Set-OrganizationRelationship  -Identity <id> -OAuthApplicationId 484a8384-979a-4cc9-8791-8e6bb34f76d4
Set-MigrationEndpoint -Identity 75f7afc6-417a-4fbe-801b-654f6b8f38e3 -Organization <org> -ApplicationId 484a8384-979a-4cc9-8791-8e6bb34f76d4 -AppSecretKeyVaultUrl <kvUrl> -SkipVerification
New-MoveRequest <id> -Remote -RemoteTenant <remoteOrg> -TargetDeliveryDomain <targetOrg> -SourceEndpoint 75f7afc6-417a-4fbe-801b-654f6b8f38e3  -whatif
#>
<#
function Verify-ApplicationLocalTenant ([bool]$localTenant, [string]$appId, [string]$targetTenant, [string]$appReplyUrl, [string]$friendTenant) {
    if ($localTenant -eq $false) {
        Write-Host "Log into $friendTenant"
        Connect-AzureAD
    }

    $consentDomain = ""
    if ($localTenant -eq $true) {
        $consentDomain = $targetTenant
    } else {
        $consentDomain = $friendTenant
    }

    $spn = Get-AzureADServicePrincipal -All $true | ? { $_.AppId -eq $appId }
    if (!$spn) {
        Write-Error "SPN of the app was not created in $consentDomain tenant"
        return
    }

    # Check MSGraph and EXO has incoming app roles assignment from the tenant friending app
    # 1. collect spns of MSGraph and EXO applications
    $spns = Get-AzureADServicePrincipal -All $true | ? { $_.AppId -in @($MS_GRAPH_APP_ID, $EXO_APP_ID) }
    if (!$spns) {
        Write-Error "Internal Error: SPNs of MSGraph or EXO not found."
        return
    }

    $spnExists = $true
    $spns | % {
        # Get SPN of an App
        # https://graph.microsoft.com/beta/tgttenant.onmicrosoft.com/servicePrincipals?$filter=appId eq '851174ff-ddd3-4bfe-b5fe-c7e5af95143c'
        # Get application roles assigned from SPN
        # https://graph.microsoft.com/beta/tgttenant.onmicrosoft.com/servicePrincipals/f05a1a01-a082-46b5-bd81-1bc66e13e408/appRoleAssignedTo
        # If admin consented then there is an app role assignment from App -> MSGraph/EXO
        $appRoleAssignments = Get-AzureADServiceAppRoleAssignment -ObjectId $_.ObjectId -All $true | ? { $_.PrincipalId -eq $spn.ObjectId }
        if (!$appRoleAssignments -and $spnExists -eq $true) {
            $spnExists = $false
            Write-Error "The app: $appId is not consented by tenant admin of $consentDomain. Please consent using the following link:"
            "https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}" -f $consentDomain, $appId, $appReplyUrl
            return
        }
    }

    if ($spnExists) {
        Write-Host "Application $appId is setup correctly in $consentDomain tenant" -Foreground Green
    }
}

function Remove-AppRoleAssignment ([string]$appId, [string]$appIdToRemovePermissionOn) {
    # App: $appId has permission on $appIdToRemovePermissionOn
    # First getting spn of appId
    $spn = Get-AzureADServicePrincipal -All $true | ? { $_.AppId -eq $appId }
    if (!$spn) {
        Write-Error "SPN of the app was not created in $consentDomain tenant"
        return
    }

    # Get spn of app which $appId has permission on, this would be either MSGraph or EXO application
    $spnIdToRemovePermissionOn = Get-AzureADServicePrincipal -All $true | ? { $_.AppId -eq $appIdToRemovePermissionOn }
    if (!$spnIdToRemovePermissionOn) {
        Write-Error "Internal Error: SPNs of MSGraph or EXO not found."
        return
    }

    $appRoleAssignments = Get-AzureADServiceAppRoleAssignment -ObjectId $spnIdToRemovePermissionOn.ObjectId -All $true | ? { $_.PrincipalId -eq $spn.ObjectId }
    if (!$appRoleAssignments) {
        Write-Error "The app: $appId does not have any permission on $appIdToRemovePermissionOn"
        return
    }

    # Remove the app role.
    Remove-AzureADServiceAppRoleAssignment -ObjectId $appRoleAssignments.PrincipalId -AppRoleAssignmentId $appRoleAssignments.ObjectId
}

function Get-AdministrativeUnits ($authRes) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $authRes.CreateAuthorizationHeader())
    Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/beta/administrativeUnits" -ContentType 'application/json' -Headers $headers
}

function Create-AdministrativeUnit ($authRes) {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", $authRes.CreateAuthorizationHeader())
    $AuCreationBody = @{
        displayName = "Mergers AU"
        description = "Admin unit for M&A"
    }

    $AuCreationBodyJson = $AuCreationBody | ConvertTo-Json
    Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/beta/administrativeUnits" -ContentType 'application/json' -Headers $headers -Body $AuCreationBodyJson
}#>
