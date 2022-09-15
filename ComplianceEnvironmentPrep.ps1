######  Summary #########################################
# This Script will leverage Security and Compliance     #
# Powershell, Sharepoint PnP Powershell, and ExO V2     #
# to prepare the M365 development sandbox for the       #
# workshop.  The account running the actions must be    #
# a global administrator of the tenant                  #
#########################################################

#prepare the envrionment
param ([switch]$debug,$reportpath=$env:LOCALAPPDATA)

If ($debug){
    $logpath = Join-path ($env:LOCALAPPDATA) ("CEPLOG_" + [string](Get-Date -UFormat %Y%m%d) + ".log")
    Start-Transcript -Path $logpath
}
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
$orgname = $orgname.split(".")
$orgname = $orgname[0]
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
    if(Get-RoleGroupmember -Identity "ComplianceAdministrator" | Where-Object {$_.name -in $adminuser}){
        Write-Host "$adminuser is already a member of Compliance Administrators" -ForegroundColor Yellow
    }
    else{
        Write-Host "Adding $adminuser to Compliance Administrators" -ForegroundColor Green
        Add-RoleGroupMember -Identity "ComplianceAdministrator" -Member $adminemail
    }

    if(Get-RoleGroupmember -Identity "ediscoveryManager" | Where-Object {$_.name -in $adminuser}){
        Write-Host "$adminuser is already a member of eDiscovery Managers" -ForegroundColor Yellow
    }
    else{
        Write-Host "Adding $adminuser to eDiscovery Managers" -ForegroundColor Green
        Add-RoleGroupMember -Identity "ediscoveryManager" -Member $adminemail
    }

    if(Get-RoleGroupmember -Identity "ComplianceManagerAdministrators" | Where-Object {$_.name -in $adminuser}){
        Write-Host "$adminuser is already a member of Compliance Manager Administrators" -ForegroundColor Yellow
    }
    else{
        Write-Host "Adding $adminuser to Compliance Manager Administrators" -ForegroundColor Green
        Add-RoleGroupMember -Identity "ComplianceManagerAdministrators" -Member $adminemail
    }
    
    if(Get-RoleGroupmember -Identity "InsiderRiskmanagement" | Where-Object {$_.name -in $adminuser}){
        Write-Host "$adminuser is already a member of Insider Risk Management Administrators" -ForegroundColor Yellow
    }
    else{
        Write-Host "Adding $adminuser to Insider Risk Management Administrators" -ForegroundColor Green
        Add-RoleGroupMember -Identity "InsiderRiskmanagement" -Member $adminemail
    }

    if(get-ediscoverycaseadmin | Where-Object {$_.name -in $adminuser}){
        Write-Host "$adminuser is already a member of eDiscovery Case Admins" -ForegroundColor Yellow
    }
    else{
        Write-Host "Adding $adminuser to eDiscovery Case Admins" -ForegroundColor Green
        Add-eDiscoveryCaseAdmin -User $adminemail
    }
}

#upload files to sharepoint
$temppath = $env:LOCALAPPDATA+"\mark8"+[string](Get-Date -UFormat %Y%m%d%S)
$tempfile = $temppath +".zip"

#make sure the SPO Site is actually accessible
Write-Host "Checking Connectivity to SharePoint Online" -ForegroundColor Yellow

try {get-pnpsite -ErrorAction stop > $null
}
catch {
    Write-Host "Unable to connect to site $sposite.  Please check the URL and try again later" -ForegroundColor Red
    Write-Host "Disconnecting Sessions for Cleanup" -ForegroundColor Green
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue -InformationAction Ignore
    Disconnect-PnPOnline -InformationAction Ignore
    Exit
}
Write-Host "Successfuly connected to $sposite" -ForegroundColor Green

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
#we are looking for the existence of this specific rule package as a catch before creating it
if((Get-DlpSensitiveInformationTypeRulePackage).RuleCollectionName -match "MRUS Rule Package"){
    Write-Host "Sensitive Information Types already exist, Continuing with Script" -ForegroundColor Yellow
}

else{
Write-Host "Creating Sensitive Information Types" -ForegroundColor Yellow
$rulefile = $env:LOCALAPPDATA+"\rulepack"+[string](Get-Date -UFormat %Y%m%d%S)+".xml"
(New-Object Net.WebClient).DownloadFile("https://github.com/microsoft/CompliancePartnerWorkshops/raw/main/rulepack.xml", "$rulefile")
New-DlpSensitiveInformationTypeRulePackage -FileData ([System.IO.File]::ReadAllBytes($rulefile)) > $null

Remove-item $rulefile
}

$sitrules = Get-DlpSensitiveInformationType | Where-Object {$_.Publisher -match "Microsoft Partner Accelerators"}
Foreach($sit in $sitrules){
    Write-Host "$($sit.name) verified successfully" -ForegroundColor Green
}

#create DLP Policies
if((Get-DlpCompliancePolicy).name -match "Project Obsidian DLP Policy"){
    Write-Host "Project Obsidian DLP Policy Already Exists" -ForegroundColor Yellow
}
else{
Write-Host "Creating DLP Policies" -ForegroundColor Yellow
New-DlpCompliancePolicy -Name "Project Obsidian DLP Policy" `
                        -Mode Enable `
                        -ExchangeLocation All `
                        -SharePointLocation All `
                        -OneDriveLocation All `
                        -TeamsLocation All `
                        -Comment "Policy created for RiskInvestigator on $(Get-Date -Format D)" `
                        > $null

New-DlpComplianceRule   -Name "Project Obsidian DLP policy rule" `
                        -Policy "Project Obsidian DLP Policy" `
                        -ContentContainsSensitiveInformation @{Name="Project Obsidian";minCount="1"} `
                        -GenerateAlert True `
                        -AlertProperties @{AggregationType="None"} `
                        -ReportSeverityLevel High `
                        > $null
}
if((Get-DlpCompliancePolicy).name -match "Project Olivine DLP Policy"){
    Write-Host "Project Olivine DLP Policy Already Exists" -ForegroundColor Yellow
}
else{
New-DlpCompliancePolicy -Name "Project Olivine DLP Policy" `
                        -Mode Enable `
                        -ExchangeLocation All `
                        -SharePointLocation All `
                        -OneDriveLocation All `
                        -TeamsLocation All `
                        -Comment "Policy created for RiskInvestigator on $(Get-Date -Format D)" `
                        > $null

New-DlpComplianceRule   -Name "Project Olivine DLP policy rule" `
                        -Policy "Project Olivine DLP Policy" `
                        -ContentContainsSensitiveInformation @{Name="Project Olivine";minCount="1"} `
                        -GenerateAlert True `
                        -AlertProperties @{AggregationType="None"} `
                        -ReportSeverityLevel High `
                        > $null
} 
$dlpcompliancepolicies = get-DlpCompliancePolicy | Where-Object {$_.Comment -match "RiskInvestigator*"}
Foreach($dlppolicy in $dlpcompliancepolicies){
    Write-Host "$($dlppolicy.name) verified successfully" -ForegroundColor Green
}

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
    
    If ((Get-OrganizationConfig).IsDehydrated -match "True"){
        $i=1
        Do{
            Write-Host "Confirming Organization Configuration is Complete, Check ($i of 5)" -ForegroundColor Yellow
            $orgstatus = (Get-OrganizationConfig).isDehydrated
            start-sleep 30
            $i++
        } While ($orgstatus -contains 'True' -and $i -lt 6) 
    }

    If ((Get-OrganizationConfig).IsDehydrated -match "False"){
        Write-Host "Enabling the Unified Audit Log" -ForegroundColor Green  
        Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
    }
    Else {
        Write-Host "Unable to automatically enable unified audit log. `nPlease enable manually in Microsoft Purview Compliance Portal" -ForegroundColor Red
        Write-Host "Disconnecting Sessions for Cleanup" -ForegroundColor Green
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue -InformationAction Ignore
        Disconnect-PnPOnline -InformationAction Ignore
        Exit
    }
    }

Else{
    write-host "Unified Audit Log is Enabled, Continuing with Script" -ForegroundColor Yellow
    }

Write-Host "Disconnecting Sessions for Cleanup" -ForegroundColor Green
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue -InformationAction Ignore
Disconnect-PnPOnline -InformationAction Ignore

start-process "https://compliance.microsoft.com/"
start-process "https://$orgname.sharepoint.com/sites/Mark8ProjectTeam/Shared%20Documents/"

If ($debug){
    $logpath = Join-path ($env:LOCALAPPDATA) ("CEPLOG_" + [string](Get-Date -UFormat %Y%m%d) + ".log")
    Stop-Transcript -ErrorAction SilentlyContinue
}
