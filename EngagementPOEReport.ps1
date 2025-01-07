# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

############################
#Data Security Engagement - Engagment POE Report
#Author: Jim Banach
#Version 3.5 - January 2025
##################################

#project variables
param (
    [ValidateSet("All","POEReport")]
    [string]$reporttype='POEReport',
    [string]$reportpath=$env:LOCALAPPDATA
)
if ($null -eq $env:LOCALAPPDATA) {
    Write-Host "This script requires the LOCALAPPDATA environment variable to be set."
    # Ask the user for the path to a writable folder that can be used to store the output of the script
    $env:LOCALAPPDATA = Read-Host -Prompt "Please enter the path to a folder where the script can store its output and restart the script"
    $reportpath=$env:LOCALAPPDATA
}
$outputfile=(Join-path ($reportpath) ("DLPReport_"+$reporttype+"_" + [string](Get-Date -UFormat %Y%m%d%S) + ".html"))
$poedate = $poedate = (Get-Date).AddDays(-100)

### Force date time to international for script only###
$currentThread = [System.Threading.Thread]::CurrentThread
$culture = [CultureInfo]::InvariantCulture.Clone()
$currentThread.CurrentCulture = $culture
$currentThread.CurrentUICulture = $culture

##CSS for HTML Output##
$header = @"
<style>
    h1 {
        font-family: Arial, Helvetica, sans-serif;
        color: #107C10;
        font-size: 32px;
    }

    h2 {
        font-family: Arial, Helvetica, sans-serif;
        color: #107C10;
        font-size: 26px;
    }

    h3 {
        font-family: Arial, Helvetica, sans-serif;
        color: #737373;
        font-size: 20px;
    }

    h4 {
        font-family: Arial, Helvetica, sans-serif;
        color: #737373;
        font-size: 16px;
    }

    table {
		font-size: 12px;
		border: 1px; 
		font-family: Arial, Helvetica, sans-serif;
	} 

    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}

    th {
        background: #107C10;
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

    hr {
        width:40%;
        margin-left:0;
        height:5px;
        border-width:0;
        color:gray;
        background-color:gray
    }
     
    #CreationDate {
        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;
    }
</style>
"@

##define functions##
function get-dlpolicysummary{
    #our global variables
    $dlptable = @()
    $policies = Get-DlpCompliancePolicy | Where-Object {$_.WhenCreated -gt $poedate}
    
    foreach($policy in $policies){
    $dlphashtable=@{}
        
    if ($policy.ExchangeLocation.Count -ne 0){
        $exoenabled = "Yes"
    }
    else{
        $exoenabled = $null
    }
    
    if ($policy.SharePointLocation.Count -gt 0){
        $spoenabled = "Yes"
    }
    else{
        $spoenabled = $null
    }
    
    if ($policy.OneDriveLocation.Count -gt 0){
        $od4benabled = "Yes"
    }
    else{
        $od4benabled = $null
    }
    
    if ($policy.TeamsLocation.Count -gt 0){
        $teamsenabled = "Yes"
    }
    else{
        $teamsenabled = $null
    }
    
    if ($policy.EndpointDlpLocation.Count -gt 0){
        $endpointenabled = "Yes"
    }
    else{
        $endpointenabled = $null
    }
    
    if ($policy.ThirdPartyAppDlpLocation.Count -gt 0){
        $DFCAEnabled = "Yes"
    }
    else{
        $DFCAEnabled = $null
    }
    
    if ($policy.OnPremisesScannerDlpLocation.Count -gt 0){
        $OnPremEnabled ="Yes"
    }
    else{
        $OnPremEnabled = $null
    }
        
    #put all the values in a hash table
    $dlphashtable =[Ordered]@{
        PolicyName = $Policy.Name
        CreationDate = ($policy.WhenCreated).ToString("MMM-dd-yyyy HH:mm:ss")
        PolicyMode = $policy.Mode
        ExchangeOnline = $exoenabled
        SharePointOnline = $spoenabled
        OneDrive = $od4benabled
        Teams = $teamsenabled
        EndPoints = $endpointenabled
        DefenderforCA = $DFCAEnabled
        OnPremises = $OnPremEnabled
    }
    #create the new object
    $dlppolicyobject = [PSCustomObject]$dlphashtable
    $dlptable += $dlppolicyobject
     
    }
    
    return $dlptable
}

#prepare the envrionment
if (get-installedmodule -Name ExchangeOnlineManagement -MinimumVersion 3.2.0 -ErrorAction SilentlyContinue) {
    Write-Host "Exchange Online Management v3.2.0 or better is Installed, Continuing with Script Execution"
}
else {
    $title    = 'Exchange Online Powershell v3.2.0 or Better is Not Installed'
    $question = 'Do you want to install it now?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Write-Host 'Your choice is Yes, installing module'
        Write-Host "This will take several minutes with no visible progress, please be patient" -foregroundcolor Yellow -backgroundcolor Magenta
        Uninstall-Module ExchangeOnlineManagement -Force -AllVersions -ErrorAction silentlycontinue
        Install-Module ExchangeOnlineManagement -SkipPublisherCheck -Force -Confirm:$false 
    } 
    else {
        Write-Host 'Please install the module manually to continue https://aka.ms/exov3-module'
        Exit
    }
}

Write-Host 'Connecting to Security & Compliance Center. Please logon in the new window' -ForegroundColor DarkYellow
Connect-IPPSSession -ShowBanner:$false

Write-Host 'Connecting to the Exchange Online, Please logon in the new window' -ForegroundColor DarkYellow
Connect-ExchangeOnline -ShowBanner:$false

#check to make sure the account has access to the availble cmdlets
Write-Host 'Checking Permissions' -ForegroundColor DarkYellow
if ((get-command get-dlpcompliancepolicy) -and (get-command Get-OrganizationConfig) -and (get-command get-accepteddomain) -and (get-command get-compliancesearch) ){
    Write-Host "`r`n`r`nConnected to Microsoft 365, Continuing with Script`r`n`r`n" -ForegroundColor Yellow
}
Else{ 
    Write-Host "`r`n`nAt least one needed cmdlet is missing, check account permissions described in the delivery guide and try again." -ForegroundColor Yellow
    Write-host "`r`nThe following cmdlets are required: `n Get-DlpCompliancePolicy `n Get-OrganizationConfig `n Get-AcceptedDomain `n Get-ComplianceSearch" -ForegroundColor Cyan
    Write-host "`r`nPlease assign the account the following privileges (or equivalent): `n Compliance Administrator (Purview) `n Compliance Management (ExchangeOnline)" -ForegroundColor Cyan
    Write-Host "`r`nDisconnecting Services" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue  -InformationAction Ignore
    Exit
}

#######################
#script activities
#each section below performs one part of the script
#section 1: collects all of the DLP Policies 
#section 2: we capture the current content searches,  we first pull the name of all the searches, then for each
#           we capture the data again to identify which workloads are being scanned by a given content search. 
#section 3: we are creating a unified summary table for the top fo the report. This uses the chart created in 
#           section 1 and the content search details from section 2
#section 4: this is where we are constructing the report itself.  It involves merging data from the prior two
#           sections and then using convertto-html to place it all into a report that can be provided to the
#           customer and submitted as part of the final Proof of Execution (POE) for the engagement
#######################

### section 1 DLP Policies
$dlppolicysummary = get-dlpolicysummary

### section 2, gather additional report details
$comsearch = Get-ComplianceSearch -resultsize unlimited | Where-Object {$_.CreatedTime -gt $poedate}
$searchoutput = foreach($s in $comsearch){Get-ComplianceSearch $s.name | Select-Object Name,ContentMatchQuery,@{Name='CreatedTime';Expression={$_.CreatedTime.ToString("MMM-dd-yyyy HH:mm:ss")}},@{Name='LastModifiedTime';Expression={$_.LastModifiedTime.ToString("MMM-dd-yyyy HH:mm:ss")}},@{Name='JobStartTime';Expression={$_.JobStartTime.ToString("MMM-dd-yyyy HH:mm:ss")}},CreatedBy,Status,*Location}

##new section to begin capturing IRM policy skip over if tenant does not have the IRM cmdlets loaded
if (Get-Command -Name Get-InsiderRiskPolicy -ErrorAction SilentlyContinue) { 
    $irmsearch = Get-InsiderRiskPolicy
    $irmpol = foreach($s in $irmsearch){$s | Select-Object Name,InsiderRiskScenario,createdby,@{Name='CreatedTime';Expression={$_.WhenCreated.ToString("MMM-dd-yyyy HH:mm:ss")}},@{Name='LastModifiedTime';Expression={$_.WhenChanged.ToString("MMM-dd-yyyy HH:mm:ss")}},Enabled}
} 

### section ,3 construct a unified summary table
$dlpsummarychart = $dlppolicysummary
$POEChart = [array]@()

foreach ($item in $dlpsummarychart){
  
    #create the new output hashtable
    $itemtable=[ordered]@{
        DLPPolicyName = $item.PolicyName
        CreationDate = $item.CreationDate
        PolicyMode = $item.PolicyMode
        SITSUsed = $coveredsits.CountofSits
        ExchangeOnline = $item.ExchangeOnline
        OneDrive = $item.OneDrive 
        SharePoint = $item.SharePointOnline
        Teams = $item.Teams
        Endpoints = $item.EndPoints
        DefenderforCA = $item.DefenderforCA 
        OnPremises = $item.OnPremises 
    }

    $summarychart = [PSCustomObject]$itemtable
    $POEChart += $summarychart
}

###section 4, build our html file
$tenantdetails = Get-OrganizationConfig
$scriptrunner = Get-ConnectionInformation | Select-Object userprincipalname,tenantid | Get-Unique
$domaindetails = Get-AcceptedDomain | Where-Object {$_.initialdomain -like "True"}

$reportstamp = "<p id='CreationDate'><b>Report Creation Date:</b> $((Get-date).ToString("MMM-dd-yyyy HH:mm:ss"))<br>
<b>Tenant Name:</b> $($tenantdetails.DisplayName)<br>
<b>Tenant ID:</b> $($scriptrunner.TenantID)<br>
<b>Tenant Domain:</b> $($domaindetails.Name)<br>
<b>Executed by</b>: $($scriptrunner.userprincipalname)</p>"

$reportintro = "<h1> Data Security Engagement: POE Report Details</h1>
<p><b>The following report shows a snapshot of the current status of Content Search and DLP Policy Configuration within the Microsoft 365 environment.</b> </p>
<p>Follow the guidance in the POE document for how to use these results as part of the POE Submission process.</p>"

if($reporttype -match'POEReport'){
    $poehtml = ($poechart | Where-Object {$_.Exchangeonline -like "Yes*"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,ExchangeOnline | ConvertTo-Html -PreContent "<h3>Exchange Module</h3>$reportstamp <b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($searchoutput | Where-Object {$_.exchangelocation -notlike ""} | Select-Object Name,ContentMatchQuery,ExchangeLocation,CreatedTime,LastModifiedTime,JobStartTime,CreatedBy,Status | Sort-Object -Property CreatedTime) | ConvertTo-Html -PreContent "</p><b>Content Search Results:</b>"
    $poehtml += ($searchoutput | Where-Object {$_.SharePointLocation -notlike ""} | Select-Object Name,ContentMatchQuery,SharePointLocation,CreatedTime,LastModifiedTime,JobStartTime,CreatedBy,Status | Sort-Object -Property CreatedTime) | ConvertTo-Html -PreContent "<h3>SharePoint Module</h3>$reportstamp <b>Content Search Results:</b>"
    $poehtml += ($poechart | Where-Object {$_.Teams -like "Yes*"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Teams | Sort-Object -Property CreationDate | ConvertTo-Html -PreContent "<h3>Teams Module</h3>$reportstamp <b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.Endpoints -like "Yes*" -and ($_.DLPPolicyName -notlike "DSPM*" -and $_.DLPPolicyName -notlike "Microsoft AI*")} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Endpoints | Sort-Object -Property CreationDate -Descending | ConvertTo-Html -PreContent "<h3>Endpoints Module</h3>$reportstamp<b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.Endpoints -like "Yes*" -and ($_.DLPPolicyName -like "DSPM*" -or $_.DLPPolicyName -like "Microsoft AI*")} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Endpoints | Sort-Object -Property CreationDate -Descending | ConvertTo-Html -PreContent "<h3>Data Security for AI Module</h3>$reportstamp<b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    if ($irmpol.count -gt 0){$poehtml += ($irmpol | Where-Object {$_.Name -like "DSPM*" -or $_.Name -like "Microsoft AI*"} | Select-Object Name,InsiderRiskScenario,createdby,CreatedTime,LastModifiedTime,Enabled | Sort-Object -Property WhenCreated) | ConvertTo-Html -PreContent "</p><b>IRM Policies:</b>"}

    Convertto-html -Head $header -Body "$reportintro $poehtml" | Out-File $outputfile 
}

#display report in browser
Write-Host "`nReport file available at:" $outputfile -ForegroundColor Yellow -BackgroundColor Blue
Write-host "`n`r"
# Use the appropriate command to open the file in the default browser
if ($IsMacOS -eq $true){ 
	& open $outputfile
}
elseif ($IsLinux -eq $true){
    & xdg-open $outputfile
}
else{
    Start-Process $outputfile
}

#cleanup
Write-Host "Disconnecting Services" -ForegroundColor Yellow
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue  -InformationAction Ignore

<#
.SYNOPSIS
Creates a report of the configured DLP Policies in the Tenant

.DESCRIPTION
The Data Security Engagement POE Report is a PowerShell script based assessment that leverages the Microsoft Security and Compliance PowerShell and Microsoft Graph PowerShell to gather information about the current configuration of Data Loss Prevention (DLP) Policies and Content Searches within the tenant to support the submission of the Proof of Execution document for the data security engagmenet. 

.PARAMETER ReportType
Specifies which products to display in the service summary:
PoEReport (Default)- Builds the reports for submission of the Proof of Execution

.PARAMETER ReportPath
Specifics the location to save the report and temporary files.  
Default location is the local appdata folder for the logged on user on Windows PCs if not specified.
MacOS and Linux clients will always prompt for the path

.EXAMPLE
PS> .\EngagementPOEReport.ps1
Provides the default report output for all DLP Policy information in the customers tenant. 

.EXAMPLE
PS> .\EngagementPOEReport.ps1 -reportpath c:\temp
Saves the report output to the folder c:\temp

.LINK
Find the most recent version of the script here:
https://github.com/microsoft/CompliancePartnerWorkshops

#>