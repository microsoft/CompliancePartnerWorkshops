# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

############################
#Data Security Engagement - Engagment POE Report
#Author: Jim Banach
#Version 3.2 - April 2024
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

##define our functions##
function get-dlpolicysummary{
    #our global variables
    $dlptable = @()
    $policies = get-dlpcompliancepolicy
    
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

### section 2, gather the content search configuration
$search = Get-ComplianceSearch
$searchoutput = foreach($s in $search){Get-ComplianceSearch $s.name | Select-Object Name,ContentMatchQuery,@{Name='CreatedTime';Expression={$_.CreatedTime.ToString("MMM-dd-yyyy HH:mm:ss")}},@{Name='LastModifiedTime';Expression={$_.LastModifiedTime.ToString("MMM-dd-yyyy HH:mm:ss")}},@{Name='JobStartTime';Expression={$_.JobStartTime.ToString("MMM-dd-yyyy HH:mm:ss")}},CreatedBy,Status,*Location}

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
    $poehtml += ($searchoutput | Where-Object {$_.exchangelocation -notlike ""} | Select-Object Name,ContentMatchQuery,ExchangeLocation,CreatedTime,LastModifiedTime,JobStartTime,CreatedBy,Status | Sort-Object -Property CreatedTime) |ConvertTo-Html -PreContent "</p><b>Content Search Results:</b>"
    $poehtml += ($searchoutput | Where-Object {$_.SharePointLocation -notlike ""} | Select-Object Name,ContentMatchQuery,SharePointLocation,CreatedTime,LastModifiedTime,JobStartTime,CreatedBy,Status | Sort-Object -Property CreatedTime) |ConvertTo-Html -PreContent "<h3>SharePoint Module</h3>$reportstamp <b>Content Search Results:</b>"
    $poehtml += ($poechart | Where-Object {$_.Teams -like "Yes*"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Teams | Sort-Object -Property CreationDate | ConvertTo-Html -PreContent "<h3>Teams Module</h3>$reportstamp <b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.Endpoints -like "Yes*"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,Endpoints | Sort-Object -Property CreationDate -Descending | ConvertTo-Html -PreContent "<h3>Endpoints Module</h3>$reportstamp<b>DLP Policies:</b>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")

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
# MIInwQYJKoZIhvcNAQcCoIInsjCCJ64CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD00/1nd1ERlkjc
# ZSHgv2tALuZxO1gcrbCbEtRh+L75VqCCDXYwggX0MIID3KADAgECAhMzAAADrzBA
# DkyjTQVBAAAAAAOvMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwOTAwWhcNMjQxMTE0MTkwOTAwWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOS8s1ra6f0YGtg0OhEaQa/t3Q+q1MEHhWJhqQVuO5amYXQpy8MDPNoJYk+FWA
# hePP5LxwcSge5aen+f5Q6WNPd6EDxGzotvVpNi5ve0H97S3F7C/axDfKxyNh21MG
# 0W8Sb0vxi/vorcLHOL9i+t2D6yvvDzLlEefUCbQV/zGCBjXGlYJcUj6RAzXyeNAN
# xSpKXAGd7Fh+ocGHPPphcD9LQTOJgG7Y7aYztHqBLJiQQ4eAgZNU4ac6+8LnEGAL
# go1ydC5BJEuJQjYKbNTy959HrKSu7LO3Ws0w8jw6pYdC1IMpdTkk2puTgY2PDNzB
# tLM4evG7FYer3WX+8t1UMYNTAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQURxxxNPIEPGSO8kqz+bgCAQWGXsEw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMTgyNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAISxFt/zR2frTFPB45Yd
# mhZpB2nNJoOoi+qlgcTlnO4QwlYN1w/vYwbDy/oFJolD5r6FMJd0RGcgEM8q9TgQ
# 2OC7gQEmhweVJ7yuKJlQBH7P7Pg5RiqgV3cSonJ+OM4kFHbP3gPLiyzssSQdRuPY
# 1mIWoGg9i7Y4ZC8ST7WhpSyc0pns2XsUe1XsIjaUcGu7zd7gg97eCUiLRdVklPmp
# XobH9CEAWakRUGNICYN2AgjhRTC4j3KJfqMkU04R6Toyh4/Toswm1uoDcGr5laYn
# TfcX3u5WnJqJLhuPe8Uj9kGAOcyo0O1mNwDa+LhFEzB6CB32+wfJMumfr6degvLT
# e8x55urQLeTjimBQgS49BSUkhFN7ois3cZyNpnrMca5AZaC7pLI72vuqSsSlLalG
# OcZmPHZGYJqZ0BacN274OZ80Q8B11iNokns9Od348bMb5Z4fihxaBWebl8kWEi2O
# PvQImOAeq3nt7UWJBzJYLAGEpfasaA3ZQgIcEXdD+uwo6ymMzDY6UamFOfYqYWXk
# ntxDGu7ngD2ugKUuccYKJJRiiz+LAUcj90BVcSHRLQop9N8zoALr/1sJuwPrVAtx
# HNEgSW+AKBqIxYWM4Ev32l6agSUAezLMbq5f3d8x9qzT031jMDT+sUAoCw0M5wVt
# CUQcqINPuYjbS1WgJyZIiEkBMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGaEwghmdAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEICRplREwtTXiMDG3mmzFkaSX
# 8HzK08Vp+r2Ef7OcFpqjMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQC4C+5Dnte9qDOjQD3k/b3zj5mZngFHb9KmBu1Rhf7Atfq+ZX5IfGGu
# 6ic4o/NenySPYb4HtceayLZViiutpdeSQCJKxxcr4apgadbv7RANrZsAdUOfRdnz
# i8mSILhJCmhTxWj6z7+rBY493dAmasG6D2gvlFrTdoHbuyjbKdJtnzJgyePF1o8D
# 1bDzYELxzbxHVzz0/FSo1+0LmVzFukWzjdTKUZkIaIekWJRGlt2W/gx5n08SsEax
# dehkyINCr5uG7zXeEfoMW8yHb8r4ZiBd8spXbwGjOvexw8pa3HvIOspHeg9YpB/m
# lTRTh7qt+7eUU9g4hSwQIyJ0MPIzFsI/oYIXKTCCFyUGCisGAQQBgjcDAwExghcV
# MIIXEQYJKoZIhvcNAQcCoIIXAjCCFv4CAQMxDzANBglghkgBZQMEAgEFADCCAVkG
# CyqGSIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEII3B4joG+92PZFBc03u5P609LDdm0ZWBBB2ZGdqS8hSdAgZmMLtP
# ZhIYEzIwMjQwNDMwMjIyMDA3LjA0MlowBIACAfSggdikgdUwgdIxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBF
# U046RDA4Mi00QkZELUVFQkExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2WgghF4MIIHJzCCBQ+gAwIBAgITMwAAAdzB4IzCX1hejgABAAAB3DAN
# BgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
# MzEwMTIxOTA3MDZaFw0yNTAxMTAxOTA3MDZaMIHSMQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBP
# cGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOkQwODIt
# NEJGRC1FRUJBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAi8izIDWyOD2RIonN6WtR
# YXlKGphYvzdqafdITknIhU9QLsXqpNwumGEdn2J1/bV/RFoatTwQfJ0Xw3E8xHYp
# U2IC0IY8lryRXUIa+fdt4YHabaW2aolqcbvWYDLCuQoBNieLAos9AsnTQSRfDlNL
# B+Yldt2BAsWUfJ8DkqD6lSwlfOq6aQi8SvQNc++m0AaqR0UsrCjgFOUSCe/N5N9e
# 6TNfy9C1MAt9Um5NSBFTvOg/9EVa3dZqBqFnpSWgjQULxeUFANUNfkl4wSzHuOAk
# N0ScrjhjyAe4RZEOr5Ib1ejQYg6OK5NYPm6/e+USYgDJH/utIW9wufACox2pzL+K
# pA8yUM5x3QBueI/yJrUFARSd9lPdTHIr2ssH9JGIo/IcOWDyhbBfKK/f5sYHp2Z0
# zrW6vqdS18N/nWU9wqErhWjzek4TX+eJaVWcQdBX00nn8NtRKpbZGpNRrY7Yq6+z
# JEYwSCMYkDXb9KqtGqW8TZ+I3lmZlW2pI9ZohqzHtrQYH591PD6B5GfoyjZLr79t
# kTBL/QgnmBwoaKc1t/JDXGu9Zc+1fMo5+OSHvmJG5ei6sZU9GqSbPlRjP5HnJswl
# aP6Z9warPaFdXyJmcJkMGuudmK+cSsIyHkWV+Dzj3qlPSmGNRMfYYKEci8ThINKT
# aHBY/+4cH2ASzyn/097+a30CAwEAAaOCAUkwggFFMB0GA1UdDgQWBBToc9IF3Q58
# Rfe41ax2RKtpQZ7d2zAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBf
# BgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
# bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0El
# MjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUF
# BwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEA2etvwTCvx5f8
# fWwq3eufBMPHgCqAduQw1Cj6RQbAIg1dLfLUZRx2qwr9HWDpN/u03HWrQ2kqTUlO
# 6lQl8d0TEq2S6EcD7zaVPvIhKn9jvh2onTdEJPhD7yihBdMzPGJ7B8StUu3xZ595
# udxJPSLrKkq/zukJiTEzbhtupsz9X4zlUGmkJSztH5wROLP/MQDUBtkv++Je0eav
# IDQIZ34+31z5p2xh+bup7lQydLR/9gmYQQyQSoZcLPIsr52H5SwWLR3iWR1wT5mr
# kk2Mgd6xfXDO0ZUC29fQNgNl03ZZnWST6E4xuVRX8vyfVhbOE//ldCdiXTcB9cSu
# f7URq3KWJ/N3cKEnXG4YbvphtaCJFecO8KLAOq9Ql69VFjWrLjLi+VUppKG1t1+A
# /IZ54n9hxIE405zQM1NZuMxsvnSp4gQLSUdKkvatFg1W7eGwfMbyfm7kJBqM/DH0
# /Omxkh4VM0fJUXqS6MjhWj0287/MXw63jggyPgztRf1lrhDAZ/kHvXHns6NpfneD
# FPi/Oge8QFcX2oKYdGBcEttGiYl8OfrRqXO/t2kJVAi5DTrafIhkqexfHO4oVvRO
# NdbDo4WkbVuyNek6jkMweTKyuJvEeivhjPl1mNXIcA3IqjRtKsCVV6KFxobkXvhJ
# lPwW3IcBboiAtznD/cP5HWhsOEpnbVYwggdxMIIFWaADAgECAhMzAAAAFcXna54C
# m0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZp
# Y2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMy
# MjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51
# yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY
# 6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9
# cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN
# 7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDua
# Rr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74
# kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2
# K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5
# TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZk
# i1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9Q
# BXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3Pmri
# Lq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUC
# BBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9y
# eS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/yp
# b+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulm
# ZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM
# 9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECW
# OKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4
# FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3Uw
# xTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPX
# fx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVX
# VAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGC
# onsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU
# 5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEG
# ahC0HVUzWLOhcGbyoYIC1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OkQwODItNEJGRC1FRUJBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQAcOf9zP7fJGQhQIl9Jsvd2OdASpqCBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUA
# AgUA6dvigjAiGA8yMDI0MDUwMTA1MzQ1OFoYDzIwMjQwNTAyMDUzNDU4WjB0MDoG
# CisGAQQBhFkKBAExLDAqMAoCBQDp2+KCAgEAMAcCAQACAgdEMAcCAQACAhGyMAoC
# BQDp3TQCAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEFBQADgYEAkgoBwmZRAJ736VCJ
# 4nnZxJWbdrwoYjpnqVaM3HCB1/UEa8HiEbSibUPIEBeykm41vkRkkppoVlC9gp3S
# 5g+4WNDjzwHPG1QDdRYRFZyrrp0t4jxGe1PnmzHZ7hYoAHDDyO3dOW+R9bEbkqoi
# owA7y/JMWQWGZrlksgYtyhggBUUxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdzB4IzCX1hejgABAAAB3DANBglghkgBZQME
# AgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJ
# BDEiBCDTlm2z8FyeTgKyy0gRhCfbY4QrWW55ZOPoMdblxhflzDCB+gYLKoZIhvcN
# AQkQAi8xgeowgecwgeQwgb0EIFOnF4pq2UQ/jLypnOO5YvQ67QirEQsOFfZMvKXE
# gg03MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHc
# weCMwl9YXo4AAQAAAdwwIgQgh56kPdDaCygWDTa5YE+qkSRp05Dx8VVx+wKH9SEU
# KEcwDQYJKoZIhvcNAQELBQAEggIAD2V1qZSSaPSRa5HTlPF4N8XfQSU4EN4NHhSe
# YDE/xmpiPpYG6tSmPiwJswHUH1A5z+2SyCGGjETR2YSB/zp+zbsgWi7j2hJ3BqdK
# y1VeqPXIcQasbk/y1h2ndc9wWHxJfBnE8NWeQwkHqi3y6A1zbdJd6P9fVl2ZG6qq
# GMh3LoZDtj84bv8C9+pHGB5vD00G+0fAtLh2dvdcPemIcYx8gEpwGNsokXtrDV0c
# abYdOqQ1x1z3J+HU8F84xAlkwQCdcYY/vhEx4C9Ctoo8lymNoAfCouv0cTO9tOug
# VWKT4q3yGnRe+Yv6Yqo+W4By9VwMp9jLclrhaQih8F6BZlWCeIbxnOhD6jvGWvQn
# r24idzl4gVDZMCJNroYV3Zlu+kIB6CbKKL5p3AQNwWoxeePxJru8pAopAGs9M12r
# LTjpZWLlXITKE4AX57rH1nyJbcrHGRMHOxevgUm2gINssGAuXs2dxEAWqkuYg77q
# v4vLG5QD+DWTU+Y0Ymx2lsXjORtoMbuQd8AO3vHxXy4LGjONaHsfvBdbEAAVY9S0
# OhMry01fe23YCQU9oQucfXIHPWPjsy26BCvs0fjNhPt1k9n/pvGU3/G4Z5T3d7mm
# qddZb+km1X7iJ01tkh5AH6X4Vy0t/m4unhegkXGn2JG5WCDDXJlgkWGsjwrwgO1Y
# +2cVHqA=
# SIG # End signature block
