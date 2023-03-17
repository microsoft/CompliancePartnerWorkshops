# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

#####################################
# Compliance Workshops M365 License Assessment Script
# Author: Jim Banach
# Version 1.5 - March, 2023
######################################


#project variables
param ($ReportType='Simple',$ReportPath=$env:LOCALAPPDATA,[switch]$LargeTenant=$false,[switch]$UseCustomList=$false,[string]$ListPath)
if ($null -eq $env:LOCALAPPDATA) {
    Write-Host "This script requires the LOCALAPPDATA environment variable to be set."
    # Ask the user for the path to a writable folder that can be used to store the output of the script
    $env:LOCALAPPDATA = Read-Host -Prompt "Please enter the path to a folder where the script can store its output and restart the script"
    $reportpath=$env:LOCALAPPDATA
}
$global:logfile = Join-path ($env:LOCALAPPDATA)("Local")
$Plans = @()
$outputfile=(Join-path ($reportpath) ("ActivationReport_" + [string](Get-Date -UFormat %Y%m%d%S) + ".html"))


#table to capture our outputs
$serviceusage = New-Object System.Data.Datatable
[void]$serviceusage.Columns.Add("ServiceName")
[void]$serviceusage.Columns.Add("ActivatedUsers")

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



#check to see if the Microsoft Graph Modules are installed
if (get-installedmodule -Name Microsoft.Graph -ErrorAction SilentlyContinue) {
    Write-Host "Microsoft Graph Installed, Continuing with Script Execution"
}
else {
    $title    = 'Microsoft Graph is Not Installed'
    $question = 'Do you want to install it now?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Write-Host 'Your choice is Yes, installing module'
        Write-Host "This will take several minutes with no visible progress, please be patient" -foregroundcolor Yellow -backgroundcolor Magenta
        Install-Module Microsoft.Graph -Scope CurrentUser -SkipPublisherCheck -Force -Confirm:$false 
    } else {
        Write-Host 'Please install the module manually to continue https://docs.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-beta'
        Exit
}
}

#check to see if the MSOLlicense management module is installed and install it if it is not
if (get-installedmodule -Name MSOLLicenseManagement -MinimumVersion 3.0.4 -ErrorAction SilentlyContinue) {
    Write-Host "License Management Module Installed, Continuing with Script Execution"
}
else {
    $title    = 'License Management Module is Not Installed'
    $question = 'Do you want to install it now?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Write-Host 'Your choice is Yes, installing module'
        Install-Module MSOLLicenseManagement -scope CurrentUser -SkipPublisherCheck -Force -Confirm:$false
        Import-Module MSOLLicenseManagement -Force 
    } else {
        Write-Host 'Please install the module manually to continue https://github.com/Canthv0/MSOLLicenseManagement'
        Exit
}
}
#connect to the MS Graph Using an account specified in real time
Write-Host 'Connecting to the Microsoft Graph, Please logon in the new window' -ForegroundColor DarkYellow
connect-MgGraph -Scopes 'User.Read.All','Organization.Read.All','Directory.Read.All'

## download the latest version of the license file 
#list of all of the current friendly sku product names - https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference 
Write-Host "Downloading Microsoft 365 Product names and service plan identifiers" -ForegroundColor Yellow
$ProgressPreference = 'SilentlyContinue'
$SKUCSV = Invoke-WebRequest https://aka.ms/m365productsandplans
$SKUList = $SKUCSV | ConvertFrom-Csv
$SPlanFriendly = $SKUList | Select-Object Service_Plan_Id, Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique
$ProgressPreference = 'Continue' 

#identify the selection of users to process
# -largetenant = will just choose the first 100 users identified in the tenant
# -customlist = now we can define the specific users we want in a CSV format (header must be UserPrincipalName)
# -listpath = if you want to define the userlist path in the script, else we will ask for a file dialog

If($usecustomlist -eq $false){
    Write-Host "All users will be evaulated" -ForegroundColor Green
}

else{
    try{$userpath = import-csv $listpath}
    catch{
        Write-Host "No user list provided. Please select the file" -ForegroundColor Yellow
        Add-Type -AssemblyName System.Windows.Forms
        
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'CSV (*.csv)|*.csv'
        }
        
        $null = $FileBrowser.ShowDialog()
        $userpath = import-csv $FileBrowser.FileName
        
    }
}

if($LargeTenant){
    $largeuser = get-mguser | Select-Object userprincipalname -first 100 
    Get-MGUserLicenseReport -Users $largeuser
}
elseif($userpath){
    Get-MGUserLicenseReport -Users $userpath
}
else{
    Get-MGUserLicenseReport 
}


## before we import the list, we want to make sure that we have the most recent file output by get-mguserlicensereport, depending on how long the process the script runs it may spill over a day and not have the same date as the temp path check"
## for this to work, we are now making the assumption that the LAST License_Report file in the folder is the most recent. 
$reportitem = Get-ChildItem $env:LOCALAPPDATA | Where-Object {$_.Name -like "License_Report_*"} | Sort-Object LastWriteTime | Select-Object -expand FullName -last 1
$list = import-csv $reportitem

#Get all of the availible SKUs in tenant
$AllSku = Get-MgSubscribedSku 2>%1
if ($AllSku.count -le 0) {
    Write-Error ("No SKU found! Do you have permissions to run Get-MGSubscribedSKU? `nSuggested Command: Connect-MGGraph -scopes Organization.Read.All, Directory.Read.All, Organization.ReadWrite.All, Directory.ReadWrite.All")
} 

# Build a list of all of the plans from all of the SKUs
[array]$Plans = $null
foreach ($Sku in $AllSku) {
    $SKU.ServicePlans.ServicePlanName | ForEach-Object { [array]$Plans = $Plans + $_ }
}
$Plans = $Plans | Select-Object -Unique | Sort-Object

#use the license file and the active plans for the customer subscription
foreach ($plan in $plans){
$planlist = $plan
$holdlist = $list | Where-Object -property $planlist -eq 'Success'

#filter out services that are licensed on multiple subscriptions
$holdlist = $holdlist | Select-Object userprincipalname,$planlist -unique
[void]$serviceusage.Rows.Add($Planlist,$Holdlist.count)
}

#filterout the services that have a null assignment value
$serviceusage2 = $serviceusage | Where-Object -property ActivatedUsers -ge 0 -ErrorAction SilentlyContinue

##pull the license plan names that are specific for the tenant
$spdata = @()
foreach ($Sku in $allsku){
    ForEach ($SPlan in $Sku.ServicePlans) {
        $SPLine = [PSCustomObject][Ordered]@{  
            ServicePlanName        = $SPlan.ServicePlanName
            #use 'Service_Plans_Included_Friendly_Names' from $SPlanFriendly for 'ServicePlanDisplayName'
            ServicePlanDisplayName = ($SPlanFriendly | Where-Object { $_.Service_Plan_Id -eq $SPlan.ServicePlanId }).Service_Plans_Included_Friendly_Names | Select-Object -First 1 
        }
        $SPData += [PSCustomObject]$SPLine
    }
}
$spdata = $spdata | select-object * -unique
$spdata = $spdata | Group-Object ServicePlanName -AsHashtable

#construct our final output
if ($reporttype -match 'Simple'){
    $serviceusage2 = $serviceusage2 | Where-Object {($_.Servicename -like "RMS_S_*" -or $_.ServiceName -like "COMPLIANCE_MANAGER*" -or $_.ServiceName -like "LOCKBOX_*" -or $_.ServiceName -like "MIP_S_*" -or $_.Servicename -like "INFORMATION_Barriers" -or $_.ServiceName -like "CONTENT*" -or $_.ServiceName -like "M365_ADVACNED*" -or $_.ServiceName -like "MICROSOFT_COMMUNICATION*" -or $_.ServiceName -like "COMMUNICATIONS_*" -or $_.ServiceName -like "CUSTOMER_KE*" -or $_.ServiceName -like "INFO_GOV*" -or $_.ServiceName -like "INSIDER_RISK_MANAG*" -or $_.ServiceName -like "ML_CLASSIFI*" -or $_.ServiceName -like "RECORDS_*" -or $_.ServiceName -like "EQUIVIO*" -or $_.ServiceName -like "PAM*" -or $_.ServiceName -like "PRIVACY*" -or $_.ServiceName -like "M365_ADVANCED_AUDIT*")} 
    $outputlist = $serviceusage2 | Select-Object Servicename, @{ N='FriendlyName'; E={ $spdata[$_.Servicename].ServicePlanDisplayName }}, ActivatedUsers | Sort-Object FriendlyName
    Write-host "Generating Simple HTML Report"
}

elseif ($reporttype -match 'Detailed') {
    $outputlist = $serviceusage2 | Select-Object Servicename, @{ N='FriendlyName'; E={ $spdata[$_.Servicename].ServicePlanDisplayName }}, ActivatedUsers | Sort-Object FriendlyName
    Write-host "Generating Detailed HTML Report"
}

#generate the HTML Output

$tenantdetails = Get-MgOrganization
$scriptrunner = Get-MgContext

$reportstamp = "<p id='CreationDate'><b>Report Date:</b> $(Get-Date)<br>
<b>Tenant Name:</b> $($tenantdetails.DisplayName)<br>
<b>Tenant ID:</b> $($tenantdetails.ID)<br>
<b>Tenant Domain:</b> $($tenantdetails.VerifiedDomains | Where-Object {$_.isinitial -eq "True"} | select-object -expandproperty Name)<br>
<b>Executed by</b>: $($scriptrunner.Account)</p>"

$reporttitle = "<h1> Compliance Activation Assesment Report </h1>
<p>The following document shows the current status of the license and service usage within the customers Microsoft 365 environment</p>"

if($LargeTenant){
    $summarylist = $outputlist | ConvertTo-Html -Fragment -PreContent "<h2>Individual Service Summary</h2> $reportstamp <p>IMPORTANT: Only a sample of the users below are represented. This report shows the content for  $($largeuser.count) users</p>"
}
else{
    $summarylist = $outputlist | ConvertTo-Html -Fragment -PreContent "<h2>Individual Service Summary</h2> $reportstamp"   
}

$tenantlicensedetails = $AllSku | Select-Object SkuPartNumber, ConsumedUnits, @{ n = 'TotalUnits'; e = { $_.prepaidunits.enabled } } | convertto-html -Fragment -PreContent "<h2>Microsoft 365 License Summary</h2> $reportstamp"
Convertto-html -Head $header -Body "$reporttitle $tenantlicensedetails $summarylist" -Title "Microsoft 365 Service Assesment Report" | Out-File $outputfile 

#display report in browser
Write-Host "Report file available at: " $outputfile
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
Disconnect-MgGraph

<#
.SYNOPSIS
Creates a report of Microsoft 365 License and Feature Usage

.DESCRIPTION
The Compliance Activation Assessment is a PowerShell script-based assessment that leverages Microsoft Graph to gather information about current Microsoft 365 usage. The assessment will generate a report that provides details about license and service usage for the Microsoft tenant.

.PARAMETER ReportType
Specifies which products to display in the service summary:
Simple (Default) - Only Display Compliance Related Services
Detailed - Shows all services in the tenant

.PARAMETER ReportPath
Specifics the location to save the report and temporary files.  
Default location is the local appdata folder for the logged on user on Windows PCs if not specified.
MacOS and Linux clients will always prompt for the path

.PARAMETER LargeTenant
Troubleshooting paramter, useful if you are having timeout issues with the script in extermely large tenants.
Will return the first 100 users in the tenant and generate the report off their information.  Use the customlist option before using this option.

.PARAMETER UseCustomList
Prompts for a csv file of users to evaualte as part of the script processing. List must be a csv of valid user principal names.
Invalid users in the list will throw an exception during the running of the script.  This is the preferred option if you are running into timeout issues running against a very large tenant. 

.PARAMETER ListPath
Provide a direct path to the csv file used for the custom list.  If this paramter is not specificed when using the custom list, a dialog box will prompt to select a file.

.EXAMPLE
PS> .\ComplianceActivationAssessment.ps1
Provides the default report output for the customers tenant. Minimum required for workshop delivery

.EXAMPLE
PS> .\ComplianceActivationAssessment.ps1 -reportpath c:\temp
Saves the report output to the folder c:\temp

.EXAMPLE
PS> .\ComplianceActivationAssessment.ps1 -ReportType Detailed -UseCustomList -ListPath c:\temp\dataprotectusers.csv
Provides a full report of all active services for the user principal names identified in the dataprotectusers.csv file

.NOTES
Leverages the Microsoft Graph License Management report (with permission) from the following module https://github.com/Canthv0/MSOLLicenseManagement

.LINK
Find the most recent version of the script here:
https://github.com/microsoft/CompliancePartnerWorkshops

#>
