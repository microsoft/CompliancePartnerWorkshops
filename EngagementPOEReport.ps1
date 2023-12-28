# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

############################
#Data Security Engagement - Engagment POE Report
#Author: Jim Banach
#Version 3.0 - January 2024
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
    $poehtml += ($searchoutput | Where-Object {$_.exchangelocation -like "all"} | Select-Object Name,ContentMatchQuery,CreatedTime,LastModifiedTime,JobStartTime,CreatedBy,Status | Sort-Object -Property CreationDate -Descending) |ConvertTo-Html -PreContent "</p><b>Content Search Results:</b>"
    $poehtml += ($searchoutput | Where-Object {$_.SharePointLocation -like "all"} | Select-Object Name,ContentMatchQuery,CreatedTime,LastModifiedTime,JobStartTime,CreatedBy,Status) |ConvertTo-Html -PreContent "<h3>SharePoint Module</h3>$reportstamp <b>Content Search Results:</b>"
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
# MIIoLAYJKoZIhvcNAQcCoIIoHTCCKBkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAVwgrMRFxvyDMz
# DaDgJEKbunM4fWUZ+V6GKUI9m8UM26CCDXYwggX0MIID3KADAgECAhMzAAADrzBA
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGgwwghoIAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINpesCt2HttYkaEiqjdXFBPJ
# pidfLMtx8TMPPEk+JvggMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQAWlwkw3ZUnV/VsKslkG6OUl/8ajxGc8lP3+SbtfHllcTTjZP0k0f9G
# rCEMovZV4uXEFFRbP31gbuOAn+zh25lNMG7m6RHhQPXj4ttc4fb7/x6BIFNlVKvB
# MT0srGHxpINKXFbn/9Cvq+E0cpJohUjCD259B04SZCWVeLdaoK4o5WwNUwRrMXc1
# WSTmkh4ONVF7yfbF8z0rwg2MfhRMj08KqrbrIEqmLsHte8/d+HVJimeXS1zkTz6a
# eVrBJFxKt8wNuBO/oM5ggEzuC13g4/f02+h9ewl1W8FHY7iNBeFB1ogbzDwwdT5B
# w6nS9CV1rdFtyvb+6uSSYAtJCjLE1yNYoYIXlDCCF5AGCisGAQQBgjcDAwExgheA
# MIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIDnN1rjxJzQoonpJiJ8EpO91MxUJt3i4HDIEF86sjfb2AgZlexEH
# WsIYEzIwMjMxMjI3MjAwMDEwLjQzNlowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo4OTAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEeowggcgMIIFCKADAgECAhMzAAAB0x0ymhc7QDBzAAEAAAHTMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIzMDUyNTE5
# MTIyNFoXDTI0MDIwMTE5MTIyNFowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo4OTAwLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALSOq5M3iDXFuFcJzwxX5067xSpzcpttpa2Lm92w
# BYzUPh9VKL7g1aAa0/8FVFitWPahWeczLR5rOJ1A4ni5SxwExs8dozFo2mBtEb0U
# RBEWdwBSm1acj5U+Xnc8Pow8vTLPxwcLZkPfB4XjD64wMAacvfoGSbSys41e+cz1
# 42+cbl2OikSqIeh1ZJq5HJ7i5+0FHaxPAYdWbEq7QZLh87zs2BsnhUbMgJHJlfD3
# 5G+9cwb+OEzXUfwBYrMqmfSgwabUxIx428tRZvfUdJl6TH80ES1e+Z2jvk5XTfQ0
# eAheKHFgR5KBQjF9sjk6aAyr9UMJCnav9/L/k1VrcqMJCg2qaYQzqisAnZcqNiEQ
# nOinidYJwn3vRTqtekE8rhcY0oEWGEtrvhMz/KxMUisRc4kbV9S5d9x1ZvQTHQUB
# 5NOvqCaYKqt4k16M0d98b9UR4Xss29Sq5gVGd2IJSGDLrbitbqm1ydBOJF8TRAv+
# AsXjWQDa9kxjNxzXoSJhdBAFoXdcC0x26HV2lepM89AQ7cyzn/kH8q2OFKykxw9S
# 9G9vfkhY36r4v7MTCKmGacIYVO7I4ypzlATSu4Y3czHRW/rH+Fw6ZpfGsdAak0oj
# k+fv1iTz0ByWpTaZcfPVkdan4oFzcPpU/svfYmXDGEnHdqxrTznG/Rc8PnwxFbVZ
# oa9pAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQU0scghrgUAPj3jPfmG/MKabTjXmIw
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAEBiWFihRD7hppDngwU18ToTLy/ita/4
# u0NFKMwzZf2Di5qcD1xTtWK12kg9X/MTq/gASF79WeDZQBHmqPZJXezP58Oo3pUt
# ZRmwpHRBHYlhcqcU9FWPXp7NnI/vN3kfwiy+xwRyid5f5pEcXTEYYzi0MutLzi+P
# pGbRuChYtdacxNnmQ/ijCcaabQuyYie67QYqsNmeR5NWZ+TyBNPLx3XLc/YhhzZQ
# jiIlhcK5JooK4V47TCrKxym+EZBKejVcAUrehrJu4PWZKhDFP2rvv4sAYZBuJKga
# WBONBBrJixBo9wbVDhA3A40aqQBIJlNvMmWeaQeCRaUpItO6U5qKVYhjiFLURn7D
# 6xfQEn0twzXjaHnU6Vcsyg8unMcBvrHbaKloAnkp/e7IVo4pbDiGe7TNaz48o93X
# 3ad14raiBZ9oV1+cS+RYMMfZ2gv5kDlAF3xeeCz+Z3cGueWXYGRn+CJkT98rKiWu
# JHdpMBYLEUJcoiX8KW7ZtueP2p9VgukBVARw9oJ9MB/s5kGVeaW4RO+rVj9I2HEL
# ownVAsKeRdIj/+JdimZEpPvzdApGCaj/jO2Pe4v1nvFtsbEhKD4/QdNFfXnLhNF4
# Fs7ZEU3IKPzyA45GT6zBPWRopdR8YHjOODle6XFJvLe4s3FB5sTpMTdwArT5+djl
# SkdoR2XDh7uKMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1WjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjKNVf2AX9sSuDivbk+
# F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU
# 88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJprx2rrPY2vjUmZNqY
# O7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVWTe/dvI2k45GPsjksUZzp
# cGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3rMvrg0Xn
# Rm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De+JKRHh09/SDPc31BmkZ1
# zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZN
# N3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyOArxCaC4Q6oRRRuLR
# vWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTY
# uVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6bMURHXLvjflSxIUX
# k8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB
# 2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQWBBQqp1L+ZMSavoKR
# PEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0g
# BFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQM
# MAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQE
# AwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQ
# W9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNv
# bS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBa
# BggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MA0GCSqG
# SIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOX
# PTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6U03dmLq2HnjYNi6c
# qYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/z
# jj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis9/kpicO8F7BUhUKz
# /AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTpkbKpW99Jo3QMvOyR
# gNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdU
# bZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz138eW0QBjloZkWsNn6Qo
# 3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4K
# u+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7Fx0ViY1w/ue10Cga
# iQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9
# vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQtB1VM1izoXBm8qGC
# A00wggI1AgEBMIH5oYHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODkwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AFLHbdwxw0HUhDCz8tiRFdrsjkmwoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDpNrJwMCIYDzIwMjMxMjI3MTQy
# NTUyWhgPMjAyMzEyMjgxNDI1NTJaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOk2
# snACAQAwBwIBAAICAekwBwIBAAICE7kwCgIFAOk4A/ACAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkq
# hkiG9w0BAQsFAAOCAQEAn89BuxNOtH+QbXmdZdJa3bjAik/mW1eCOjXSf4ynnqds
# X4lq30uWqE0bBAUuNOfY/sgztahAOZp+q7uyJV2D1gYwUYahEeaYkhEtPR1vjXJC
# MqAdcjUhVWMvVMU12uVlrKctaLZ7i2yAJXkthLlhqGVkSga7Lnsqs7Sv5FRhRp3f
# 1Vna8ZFajITdJarhYpw/Nc9zNGVcrBqUQCjHjU1gn/tJSwWbPDleUuLupIvnq/JP
# EQ5TeL9qBTNxpY0PNtx9a9fTL/vRpYY30w492ktOD/alxl5QI+ZdcqISga3LFLDj
# D8PmiSgnmxjcGssH5I5S87FApvJwkI4fKn2EOitaijGCBA0wggQJAgEBMIGTMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB0x0ymhc7QDBzAAEAAAHT
# MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# LwYJKoZIhvcNAQkEMSIEICf4AUzwhNrOQ54XvOJDVrOHe8eUSjpeq43qKcDFw9bP
# MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgkmb06sTg7k9YDUpoVrO2v24/
# 3qtCASf62Aa1jfE6qvUwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAdMdMpoXO0AwcwABAAAB0zAiBCAfIVvvY3brrfTqwzqrQBWFDpaZ
# 2l1zx8xrPzybP2NyZzANBgkqhkiG9w0BAQsFAASCAgBTddSvn6ILyyskjbG/sLC5
# iY3WMYommJZONJxi7PssoH2LkUXaM8qhurocTuTbApn5vFB8XW2ZHoxnJYpRMnhD
# Okf9zvCxQ38GXUZm64U3pYXTOAvLSVQIhaR6y1zGL2z/pqccSHV13pVbuMZatsvf
# Wv8xNDG+aRJF+7Oj+nq9ouN7PigYF5o9xRrN2nDkPgkXxQCiJXhApUyCMfhWZ4Cl
# 9Ih8+JSLmeFeLGDHiyXnx5i6Cw3uV1C1ZrR1rKivU5wEslBljEjZVBlGAP+4fznC
# ho6gJtTkXzInB1UiH8SSUVDsGYL06GEkErJN5BYgnphHOlFX3iDr+5p2WtMmWOD7
# VZn3AvLD8r9MGfwHubUmS/rwUB5TmrY4VkvPMtLkJiFjleGjDlvFj5+rat4XnqRr
# XlTuSfV+dVjwTceUwF1HsnVqz2jN+jw824Ee9jv6u2//VX2ZQB7lSaIPAXYxTVe5
# m0N4DnLFK3+zxTqVXQ5rgggqcnXrEpK9WEdYhyxTSLV12f07fWzOYh089ycvG5Lx
# qItlADGhbc7cZY/DbQhSjmFuH7jNnHHrkQpbsPgw/pTPF+Gh0x7eeS1Q58uTRE5n
# MGaMrKpw7zzt92IkdDxFHDkkVHVXcH9Au1qc4gZmuizxDX6BZjEFXFYs041JgTSh
# BxtSDTCI8N1OrxESLiTOnw==
# SIG # End signature block
