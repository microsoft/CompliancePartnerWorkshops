# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

############################
#Data Security Engagement - Engagment POE Report
#Author: Jim Banach
#Version 3.4 - November 2024
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
$poedate = $poedate = (Get-Date).AddDays(-120)

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
Connect-IPPSSession 

Write-Host 'Connecting to the Exchange Online, Please logon in the new window' -ForegroundColor DarkYellow
Connect-ExchangeOnline 

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

##new section to begin capturing IRM policy
if (Get-Command -Name Get-InsiderRiskPolicy -ErrorAction SilentlyContinue) { 
    $irmpolcount = 1
    $irmsearch = Get-InsiderRiskPolicy
    $irmpol = foreach($s in $irmsearch){$s | Select-Object Name,InsiderRiskScenario,createdby,@{Name='CreatedTime';Expression={$_.WhenCreated.ToString("MMM-dd-yyyy HH:mm:ss")}},@{Name='LastModifiedTime';Expression={$_.WhenChanged.ToString("MMM-dd-yyyy HH:mm:ss")}},Enabled}
} 
else { 
    $irmpolcount= 0
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
    $poehtml += ($poechart | Where-Object {$_.Endpoints -like "Yes*" -and ($_.DLPPolicyName -notlike "DSPM*" -and $_.DLPPolicyName -notlike "AI*")} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Endpoints | Sort-Object -Property CreationDate -Descending | ConvertTo-Html -PreContent "<h3>Endpoints Module</h3>$reportstamp<b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.Endpoints -like "Yes*" -and ($_.DLPPolicyName -like "DSPM*" -or $_.DLPPolicyName -like "AI*")} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Endpoints | Sort-Object -Property CreationDate -Descending | ConvertTo-Html -PreContent "<h3>Data Security for AI Module</h3>$reportstamp<b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    if ($irmpolcount = 1){$poehtml += ($irmpol | Where-Object {$_.Name -like "DSPM*" -or $_.Name -like "AI*"} | Select-Object Name,InsiderRiskScenario,createdby,CreatedTime,LastModifiedTime,Enabled | Sort-Object -Property WhenCreated) | ConvertTo-Html -PreContent "</p><b>IRM Policies:</b>"}

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
# SIG # Begin signature block
# MIIoRQYJKoZIhvcNAQcCoIIoNjCCKDICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBTwoeAWyJTtb+
# L5wB+kJoi4bTquucOcBf5vXg+FGNbqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGiUwghohAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJzndtInyGy7jCZpLK4jcLp0
# afEMsQ9Vio3SS+hVgUFmMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQAxoqggtfi35r+TOc8RCwasDGop5HR8ERyuGz+S1JFf/UqWS/+DkBLt
# 5KzhZSQGbx1uKcGSoESNFWruWjlYTJk6i+jgGsOkywZ8bQ8Y6Qx4a4mKF0HkqAcX
# od1ZBrh6AhsI8yVgzh8KBsKZR5Q++uU0ovhhxOUm8cRd9R9ey0CHoZbvXN9VPn3C
# b/cMhO5NyCcLPpYHHsqCJPTIvYH/+LX1mYXH3Bp30UkpIvdYCOEySyDY/BvqwjfX
# joJxz4TnRPWmXNvJ9NOYz2JjBgKIgVFijKGcLWJYopuJNcHE+EtML7YfsN6/yoIS
# zD9//YxW1BgutOm+AhJAAd5N3XBdZNvQoYIXrTCCF6kGCisGAQQBgjcDAwExgheZ
# MIIXlQYJKoZIhvcNAQcCoIIXhjCCF4ICAQMxDzANBglghkgBZQMEAgEFADCCAVoG
# CyqGSIb3DQEJEAEEoIIBSQSCAUUwggFBAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIHCPZf1QZYU7A0VzM4mkEzpMJwP4ZSJRB07VN7Ea+KXFAgZm61tQ
# maEYEzIwMjQxMTE4MTMwMDA5LjkxN1owBIACAfSggdmkgdYwgdMxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1Mg
# RVNOOjUyMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIR+zCCBygwggUQoAMCAQICEzMAAAIAC9eqfxsqF1YAAQAAAgAw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjQwNzI1MTgzMTIxWhcNMjUxMDIyMTgzMTIxWjCB0zELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQg
# T3BlcmF0aW9ucyBMaW1pdGVkMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046NTIx
# QS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCvVdpp0qQ/ZOS6ehMX
# nvf+0Xsokz0OiK/dxy/bqaqdSGa/YMzn2KQRPhadF+jluIUgdzouqsh6oKBP19P8
# aSzlUo73RlBbZq+H88weeXSRl9f8jnN1Wihcdt1RSQ+Jl81ezNazCYv1SVajybPK
# /0un1MC3D+hc5hMDF1hKo4HlTPPVDcDy4Kk2W6ZA0MkYNjpxKfQyVi6ReSUCsKGr
# qX4piuXqv9ac6pdKScAGmCBKwnfegveieYOI31hQClnCOc2H0zqQNqd5LPvz9i0P
# /akanH38tcQuhrMRQXGHKDgp2ahYY1jB1Hv+J3zWB44RHr2Xl0m/vVL+Yf4vFvov
# r3afy4SYBXDp9W8T5zzCOBhluVkI08DKcKcN25Et2TWOzAKqOo1zdf9YjMDsYazg
# dRLhgisBTHwfYD3i8M2IDwBZrtn8dLBMLIiB5VuV1dgzYG3EwqreSd5GhPbs1Dtj
# ufxlNoCN7sGV+O7zeykY/9BZg1nXBjNhUZHI6l0DxabGrlXx/mvgdob3M9zKQ7Im
# lFnL5XdEaKCEWawIlcBwzOI7voeKfAIiMiacIUoYn8hsuMfont8lepE9uD4fqtgt
# nxcGmZcEUfg9NqRlVjEH8/4RBsvks3DRkgz565/EKWNXg76ceQz7OBLHz7TsFVk8
# EqyWih5uyaEhwE7Tvv2R6FyJgwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFIAlJNq+
# CSMqqB7nFA5effl9d3c6MB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1Gely
# MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
# bDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBD
# QSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQAo8Ib2eNG0
# ipD5+4QKDHNYyxA4jce9buxX0+YRYII5aVO4YIjM2LeJt0tlLRvYgMeDTIuu11W4
# GLcXFV16whe7NjKh8h79qVMF1XgOGKMtNe0Hs2A6ejsbXaI7e3qPLWE9Kq7MYuvL
# /aYRkHAixKhLYP4f7ccInE8PKHMwWo/6mWW08AIH9A3Bnur4cbJm/e/x636tBiDy
# wXc9O5Z4ODd1H/OTU1rAn918UINiVY14IEIu809AFx4xhcVEUqFxJTCzuYV0gMOF
# mnGrIgoPZYPAXI0gYR7Of6d3iRdG6l40TH55KklfKVEP7V3jmFvo/M4gXsGRw+1G
# 0VbbBeCSMuq1NZaUGS/OXa419gncI9lVoPIwNppeA74foOKuwnggb2KQh33jX6ZY
# N6OSPlpif1A3pE5+j8c0eDW2KbCkWhSK+oAW7qKtZkXDlX7IuvwUtzudsxraUVKL
# HO73rN2cOw8ibPRzpK1tjKEpKUze4NGL1RbJ1IqqcRu0noyT5i7G/OmuS5ZAlhZ+
# k++6D7BOeKjKRXBzTJFVyx3jEzOeedG1kMYxJQSX20xWd5thGyHBkjIlOwGAtmYc
# zurZMUr9f33jhKQirJjbYBy4t7Qaqg18BIIhxm3Ntn/M/iVPb83SkNufZ98DONmS
# Ej5Cuqv3zeZBlbBl2vdOTUxgSUNOHPYPQzCCB3EwggVZoAMCAQICEzMAAAAVxedr
# ngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4
# MzIyNVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qls
# TnXIyjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLA
# EBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrE
# qv1yaa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyF
# Vk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1o
# O5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg
# 3viSkR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2
# TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07B
# MzlMjgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJ
# NmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6
# r1AFemzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+
# auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3
# FQIEFgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl
# 0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUH
# AgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
# b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMA
# dQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAW
# gBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8v
# Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRf
# MjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL
# /Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu
# 6WZnOlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5t
# ggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfg
# QJY4rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8s
# CXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCr
# dTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZ
# c9d/HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2
# tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8C
# wYKiexcdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9
# JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDB
# cQZqELQdVTNYs6FwZvKhggNWMIICPgIBATCCAQGhgdmkgdYwgdMxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1Mg
# RVNOOjUyMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQCMk58tlveK+KkvexIuVYVsutaOZKCB
# gzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEB
# CwUAAgUA6uWYhTAiGA8yMDI0MTExODEwNDIxM1oYDzIwMjQxMTE5MTA0MjEzWjB0
# MDoGCisGAQQBhFkKBAExLDAqMAoCBQDq5ZiFAgEAMAcCAQACAgkcMAcCAQACAhK1
# MAoCBQDq5uoFAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAI
# AgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBACGw+1hy8mif
# +K3HGJ/04tMWHCvYtWHVHmIHEhmkH+WwibrVC5rRDjGl+g9rKi/DoNeOd8KwXwhk
# snyE6KrZfLaHxFSXKK4V6pqyG0d0pO0MgulRUsARMMkflLIPhkTzUblOyaeJDX+P
# qQAkAkEPoKv6Y5wyPwm17xWsA5b1erfIGPXOQ+rKXMyvMYYv+9NT1NiwGphk0QWy
# kXoWKb9KQfzx1NvCgSvb+bo5cPvDRRpiHqhT5+qTvDJzEHNYi3eUSWzf1uuwPEWa
# Mn3iTIGGqQwYA1nbRYxn/mslUEFGWsJJ1u+w9hbu8k3w1ckqinlvzE6e8E7hkjsA
# AGX6sZTuQVAxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAgAL16p/GyoXVgABAAACADANBglghkgBZQMEAgEFAKCCAUowGgYJ
# KoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDNTMtSSMIn
# RTNAWZInlEyvjqHasYbuzd3WnvSLECI7PjCB+gYLKoZIhvcNAQkQAi8xgeowgecw
# geQwgb0EINTI7ew1ndu6sE0MZQJXg18zaAfhpa5G50iT/0oCT9knMIGYMIGApH4w
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIAC9eqfxsqF1YAAQAA
# AgAwIgQgGDKsjOxdZSY/n23lSjxEmsha4EPmK05AQ5ReOFJsm7IwDQYJKoZIhvcN
# AQELBQAEggIAbGgeBkoXgVnNmrcGp1nX5KUPqzVuHVi9ES0Eq7/xzvAk/ju40UGV
# 42pYbjWYwEdbaO4Zqwhfy4HKHsWk8JTOzlQBY76bX/SyU/RgkPBYUL2OesUfJ4MP
# 08ncz9Gfl5IkaKr1GMNQu8VwqlJ9WP7OrQGYV190JpI+PlmuOz/Ga1KSgtgq0kIE
# fI0DLBlvlyorOtWald/6p6trZtLmIgPuEI10nz39jmcHKJcnlV9CmbDblThnzdDA
# f2L5moHDonTjpRCZ4DuUFFGGxe1LT1PQH2imZYhBn8DEscLvvyIzsiEj13waYlAh
# JcAGjbf5hy9YlQ6ipT1kUBoLMLaPCIgyzGwPaOBUoJvlbITNHed5wB1kkpv0fg4A
# pTRwR4aS9tVRHYtpccKnBL7HcHsQ5Ix50ignR5B72w79CvBrIbuxEs5PWpQtgOng
# E/AOJv94JhpKvxsM1wruLCSxeHns/YmjofEcKEGef84d2UvCXVevRnQSifKacV6y
# fKC7A1Dm0i4KjO5BM44VFhBs3mE7XDaettOepVyXE43WMxsP3DDD5tP5fH0yiytF
# lsA6WvQbDBb2lUe5Vwn613ZMUs+roZT+rrAmJxnfErvyl6ZHqWfyq3FwxUz6LqBY
# 3zdO/RIuJ6/wfYGw3Hvmrjn1U0MH48T2iD4mMiM8vmTumsLfL8p6his=
# SIG # End signature block