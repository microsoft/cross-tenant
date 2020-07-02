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
# Author: Thomas Rudolf (Senior Premier Field Engineer)
# Date: 2020-07-02
#
#################################################################################

param(
	[Parameter(Mandatory=$true)][string] $target,
	[string]$filename = "C:\temp\source.csv"
)

if(!(Get-Command Get-Mailbox -ErrorAction SilentlyContinue))
{
	Write-Host "Exchange PowerShell module not loaded" -foregroundcolor red
	break
}

$source = Import-CSV $filename
if(!$source)
{
	Write-Host "Import of CSV file failed" -foregroundcolor red
	break
}

#Is this really necessary? Customer often syncing contacts for collaboration
if((Get-Contact $target -ErrorAction SilentlyContinue))
{
	Remove-MailContact $target -Confirm:$false
	Write-Host "Note: Existing Contact removed" -foregroundcolor yellow
}
$target2 = Get-MailUser $target -ErrorAction SilentlyContinue
if($target2 -eq $null)
{
	Write-Host "MailUser not existing!" -foregroundcolor red
	break
}
if($target2.SKUAssigned -ne $null)
{
	Write-Host "Don't assign a license before running this script!" -foregroundcolor red
	break
}
if($target2.ExternalEmailAddress -notlike "*.onmicrosoft.com")
{
	Write-Host "Requires MailUser with ExternalAddress sourceMBX@sourcetenant.onmicrosoft.com!" -foregroundcolor red
	break
}
if(!($target2.EmailAddresses -like "*.onmicrosoft.com*"))
{
	Write-Host "Requires MailUser with target e-mail address targetMBX@targettenant.onmicrosoft.com!" -foregroundcolor red
	break
}
$target2.EmailAddresses += "X500:"+$source.LegacyExchangeDN
Set-MailUser $target -ExchangeGuid $source.ExchangeGuid -ArchiveGuid $source.ArchiveGuid -EmailAddresses $target2.EmailAddresses
Write-Host "Note: Please assign an Exchange license" -foregroundcolor yellow

	#temporary workaround (moverequest instead of migrationbatch) until build 15.20.3125
	New-MoveRequest $target -TargetDeliveryDomain tratongroup.onmicrosoft.com -Remote -RemoteTenant manonlineservices.onmicrosoft.com
	#Feature? add parameters SuspendWhenReadyToComplete or CompleteAfter
#CSV with $target
#$target | Export-CSV c:\temp\source2.csv -NoTypeInformation
#New-MigrationBatch -Name ("T2Tbatch-"+(get-date -Format yyyyMMddTHHmmssffffZ)) -AutoStart -AllowUnknownColumnsInCsv:$true -CSVData ([System.IO.File]::ReadAllBytes("c:\temp\source2.csv")) -TargetDeliveryDomain tratongroup.onmicrosoft.com -SourceEndpoint tratongroup_manonlineservices_6671

#Get-MoveRequest -Flags "CrossTenant"
