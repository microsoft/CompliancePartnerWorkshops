######  Summary #########################################
# This Script will leverage Security and Compliance     #
# Powershell, Sharepoint PnP Powershell, and ExO V2     #
# to prepare the M365 development sandbox for the       #
# workshop.  The account running the actions must be    #
# a global administrator of the tenant                  #
#########################################################

#prepare the envrionment
if (get-installedmodule -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue) {
    Write-Host "Exchange Online Management Installed, Continuing with Script Execution"
}
else {
    $title    = 'Exchange Online Powershell is Not Installed'
    $question = 'Do you want to install it now?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Write-Host 'Your choice is Yes, installing module'
        Write-Host "This will take several minutes with no visible progress, please be patient" -foregroundcolor Yellow -backgroundcolor Magenta
        Install-Module ExchangeOnlineManagement -SkipPublisherCheck -Force -Confirm:$false 
    } else {
        Write-Host 'Please install the module manually to continue https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps'
        Exit
}
}

if (get-installedmodule -Name PnP.Powershell -ErrorAction SilentlyContinue) {
    Write-Host "SharePoint PnP Module Installed, Continuing with Script Execution"
}
else {
    $title    = 'Sharepoint PnP Module is Not Installed'
    $question = 'Do you want to install it now?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Write-Host 'Your choice is Yes, installing module'
        Write-Host "This will take several minutes with no visible progress, please be patient" -foregroundcolor Yellow -backgroundcolor Magenta
        #do this to force remove the old PNP module just in case it is still there
        Uninstall-Module SharePointPnPPowerShellOnline -Force -AllVersions -ErrorAction silentlycontinue
        Install-Module pnp.powershell -SkipPublisherCheck -Force -Confirm:$false -WarningAction SilentlyContinue 
    } 
    else {
        Write-Host 'Please install the module manually to continue https://pnp.github.io/powershell/articles/installation.html'
        Exit
    }
}

Write-Host "Connecting to Security & Compliance Center. Please logon in the new window" -ForegroundColor Yellow
Connect-IPPSSession

#Capture the org name without requiring the partner to input it and possibly misstype
$orgname = get-user "nestor wilke" | Select-Object -expandproperty userprincipalname
$pos = $orgname.indexof("@")
$orgname = $orgname.substring($pos+1)
$orgname = $orgname.TrimEnd(".onmicrosoft.com")
$sposite = "https://$orgname.sharepoint.com/sites/Mark8ProjectTeam"

Write-host "Connecting to SharePoint Online, Please logon in the new window" -ForegroundColor yellow
Connect-PnPOnline -url $SPOSite -Interactive

#set the compliance portal permissions

### we are defining ALL of the users that exist in the dev lab,
### it's only 16 so it's easy enough to make the array and then
### look for the one additional value to find the account that is the global
### if the partner or customer makes more accounts that's fine as we will add 
### them to the permissions as well.  The difference is expected to be very short
### so we will iterate through them one at a time vs adding all users to each group

$OOTBusers = @('Patti Fernandez','Nestor Wilke','Lidia Holloway','Lynne Robbins','Pradeep Gupta','Lee Gu','Joni Sherman','Adele Vance','Miriam Graham','Megan Bowen','Grady Archie','Diego Siciliani','Isaiah Langer','Henrietta Mueller','Alex Wilber','Johanna Lorenz','Discovery Search Mailbox')
$tenantusers = Get-User | Select-Object -ExpandProperty name
$adminusers = $tenantusers | Where-Object {$_ -notin $OOTBusers}

foreach ($adminuser in $adminusers){
    $adminemail = (get-user $adminuser).userprincipalname
    Add-RoleGroupMember -Identity "ComplianceAdministrator" -Member $adminemail
    Add-RoleGroupMember -Identity "ediscoveryManager" -Member $adminemail
    Add-RoleGroupMember -Identity "ComplianceManagerAdministrators" -Member $adminemail
    Add-RoleGroupMember -Identity "InsiderRiskmanagement" -Member $adminemail
    Add-eDiscoveryCaseAdmin -User $adminemail    
}

#upload files to sharepoint
$temppath = $env:LOCALAPPDATA+"\mark8"+[string](Get-Date -UFormat %Y%m%d%S)
$tempfile = $temppath +".zip"

(New-Object Net.WebClient).DownloadFile("https://github.com/microsoft/CompliancePartnerWorkshops/raw/main/mark8.zip", "$tempfile")
Expand-Archive -LiteralPath $tempfile -DestinationPath $temppath -Force

Write-Host "Uploading sample files to $sposite" -ForegroundColor yellow
$docpath = get-item $temppath
$documents = Get-ChildItem -Path $docpath.FullName
$pnpcount = $null
foreach ($document in $documents) 
    {
        Write-Host "Uploading file $document" -ForegroundColor Cyan
        $pnpresult = Add-PnPfile -path $document.FullName -folder "shared documents"
        $pnpcount +=1

    }

Write-Host "Uploaded $pnpcount files to $sposite" -ForegroundColor Green

Remove-Item $temppath -Recurse
Remove-Item $tempfile

#create our DLP Policies

#create SITs
Write-Host "Creating Sensitive Information Types" -ForegroundColor Yellow
$rulefile = $env:LOCALAPPDATA+"\rulepack"+[string](Get-Date -UFormat %Y%m%d%S)+".xml"
(New-Object Net.WebClient).DownloadFile("https://github.com/microsoft/CompliancePartnerWorkshops/raw/main/rulepack.xml", "$rulefile")
New-DlpSensitiveInformationTypeRulePackage -FileData ([System.IO.File]::ReadAllBytes($rulefile))

Remove-item $rulefile

#create DLP Policies
Write-Host "Creating DLP Policies" -ForegroundColor Yellow
New-DlpCompliancePolicy -Name "Project Obsidian DLP Policy" `
                        -Mode Enable `
                        -ExchangeLocation All `
                        -SharePointLocation All `
                        -OneDriveLocation All `
                        -TeamsLocation All `
                        -Comment "Policy created during workshop lab on $(Get-Date -Format D)"

New-DlpComplianceRule   -Name "Project Obsidian DLP policy rule" `
                        -Policy "Project Obsidian DLP Policy" `
                        -ContentContainsSensitiveInformation @{Name="Project Obsidian";minCount="1"} `
                        -GenerateAlert True `
                        -AlertProperties @{AggregationType="None"} `
                        -ReportSeverityLevel High

New-DlpCompliancePolicy -Name "Project Olivine DLP Policy" `
                        -Mode Enable `
                        -ExchangeLocation All `
                        -SharePointLocation All `
                        -OneDriveLocation All `
                        -TeamsLocation All `
                        -Comment "Policy created during workshop lab on $(Get-Date -Format D)"

New-DlpComplianceRule   -Name "Project Olivine DLP policy rule" `
                        -Policy "Project Olivine DLP Policy" `
                        -ContentContainsSensitiveInformation @{Name="Project Olivine";minCount="1"} `
                        -GenerateAlert True `
                        -AlertProperties @{AggregationType="None"} `
                        -ReportSeverityLevel High
                        
### we check for the admin audit log as the last piece of this.  
### this way we keep the exchangeonline and security and compliance
### cmdlets from stomping on each other
Write-Host "Checking Unified Audit Configuration" -ForegroundColor Green
Write-host "Connecting to Exchange Online to complete check, please logon in new window" -ForegroundColor Yellow
Connect-ExchangeOnline -ShowBanner:$false

if ((Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled -match "False"){
    If ((Get-OrganizationConfig).IsDehydrated -match "True"){
        Write-Host "Organization Customization is not enabled. Enabling, please wait" -ForegroundColor Magenta
        Enable-OrganizationCustomization
    }
    Write-Host "Enabling the Unified Audit Log" -ForegroundColor Green  
    Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
    }

    Else{
    write-host "Unified Audit Log is Enabled, Continuing with Script" -ForegroundColor Yellow
}

Write-Host "Disconnecting Sessions for Cleanup" -ForegroundColor Green
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue -InformationAction Ignore
Disconnect-PnPOnline -InformationAction Ignore

### need to update these with variables when fixing up the other code ##
#start-process "https://compliance.microsoft.com/"
#start-process "https://6glf4w.sharepoint.com/sites/Mark8ProjectTeam/Shared%20Documents/"
