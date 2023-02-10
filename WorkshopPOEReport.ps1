# Copyright (c) Microsoft Corporation.
# Licensed under the MIT license.

############################
#Protect and Govern Sensitive Data Activator - Workshop POE Report
#Author: Jim Banach
#Version 1.0 - October,2022
##################################

#project variables
param ($reporttype='All',$reportpath=$env:LOCALAPPDATA)
$outputfile=(Join-path ($reportpath) ("DLPReport_"+$reporttype+"_" + [string](Get-Date -UFormat %Y%m%d%S) + ".html"))
# $a is a variable that helps to build out the HTML report body.  Will update to something more descriptive at a later date
$a=@()
$policycounts= @()
$sitcounts = @()

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
        CreationDate = $Policy.WhenCreated
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
function get-dlppolicydetails($param){
    
    $policy = get-dlpcompliancepolicy -Identity $param
    
    
    #the array we will return
    $dlppolicydetailtable = @()
    
    #the variables needed to be used for each policy we loop through
    $exchangemembers=@()
    $exchangememberscount = 0
    $sharepointlocations=@()
    $sharepointlocationscount = 0 
    $onedrivemembers=@()
    $onedrivememberscount = 0 
    $onedrivesites=@()
    $onedrivesitescount = 0 
    $teamslocations=@()
    $teamslocationscount = 0 
    $endpointdlplocations=@()
    $endpointdlplocationcounts = 0 
    $onpremlocations=@()
    $onpremlocationscount = 0 

    #the hashtable we are going to store everything in and return
    $dlphashtable=@{}

Write-Host "Processing the policy:" $policy.Name -ForegroundColor Green
#we now need to check to see which locations the policy is enabled for, we will do this by checking for the presence of data in the *location* variables

    #Exchange Activity Block
    if ($policy.ExchangeLocation.Count -gt 0){
        Write-host $policy.name "Policy is enabled for Exchange" -ForegroundColor Yellow
        if($policy.ExchangeSenderMemberOf.count -eq 0){
            $exchangemembers += "All Users"
            $exchangememberscount = "All Users"
        }
        
        else{        
        ##parse out who exchange is enabled for and store those items in the array 'exchangemembers'
        $exchangelist = $policy.ExchangeSenderMemberOf
        $exchangelist = $exchangelist | Select-Object -unique
            
        foreach($exmember in $exchangelist){
            $exmemberitem = $exmember.split(",")
            $exmemberitem = $exmemberitem | Select-String -Pattern 'Display'
            [string]$exmemberitemstring = $exmemberitem
            $exmemberitemstring = $exmemberitemstring.split(":")[1] -replace '["]'
            
            #get the number of users in the group and add it to the counter
            $exgroup = Get-AzureADGroup -SearchString $exmemberitemstring
            try{$exgroupcount = (Get-AzureADGroupMember -ObjectId $exgroup.objectid -all $true -ErrorAction SilentlyContinue).Count }
            catch{$exgroupcount = 0}
            $exchangememberscount += $exgroupcount
            
            #return the string with the count of that group in parentheses
            $exmemberitemstring = $exmemberitemstring + " (" + $exgroupcount + ")"
            $exchangemembers += $exmemberitemstring
            }
        }
            #write-host "Exchange is enabled for" $exchangemembers
    }
    
    #sharePoint Activity Block
    if ($policy.SharePointLocation.Count -gt 0){
        write-host $policy.name "Policy is Enabled for Sharepoint" -ForegroundColor Yellow
        
        if($policy.SharePointLocation.name -eq "All"){
            $sharepointlocations += "All Sites"
            $sharepointlocationscount = "All Sites"
            #write-host "OneDrive is enabled for" $sharepointlocations
        }

        else{
        foreach($site in $policy.SharePointLocation){
            $site = $site | Select-Object DisplayName,Name
            $site = ($site.Displayname,$site.name) -join ": "
            $sharepointlocations += $site 
            $sharepointlocationscount++
            }
        }
    
        #write-host "SharePoint is enabled for" $sharepointlocations
    }
    
    #onedrive Activity Block
    if ($policy.OneDriveLocation.Count -gt 0){
        write-host $policy.name "Policy is Enabled for OneDrive" -ForegroundColor Yellow
        
        if($policy.OneDriveSharedBy.count -eq 0){
            $onedrivemembers += "All Users"
            $onedrivememberscount = "All Users"
            #write-host "OneDrive is enabled for" $onedrivemembers
        }

        else{
            #get-individual onedrive users
            if ($policy.OneDriveSharedBy.Count -gt 0){
                $odmemberlist = $policy.OneDriveSharedBy
                $odmemberlist = $odmemberlist | Select-Object -unique
        
                foreach($odmember in $odmemberlist){
                    $odmemberitem = $odmember.split(",")
                    $odmemberitem = $odmemberitem | Select-String -Pattern 'Display'
                    [string]$odmemberstring = $odmemberitem
                    $odmemberstring = $odmemberstring.split(":")[1] -replace '["]'               
                    $onedrivemembers += $odmemberstring
                    $onedrivememberscount++
                    }
                    #write-host "OneDrive is enabled for" $onedrivemembers
                }
            }
            #get individual onedrive sites
            if ($policy.OneDriveSharedByMemberOf.Count -gt 0){
                $odsitelist = $policy.OneDriveSharedByMemberOf
                $odsitelist = $odsitelist | Select-Object -unique
        
                foreach($odsite in $odsitelist){
                    $odsiteitem = $odsite.split(",")
                    $odsiteitem = $odsiteitem | Select-String -Pattern 'Display'
                    [string]$odsitestring = $odsiteitem
                    $odsitestring = $odsitestring.split(":")[1] -replace '["]'
                    
                    #get the number of users in the group and add it to the counter
                    $odgroup = Get-AzureADGroup -SearchString $odsitestring
                    try{$odgroupcount = (Get-AzureADGroupMember -ObjectId $odgroup.objectid -all $true).Count}
                    catch{$odgroupcount = 0}
                    $onedrivesitescount += $odgroupcount
            
                    #return the string with the count of that group in parentheses
                    $odsitestring = $odsitestring + " (" + $odgroupcount + ")"                                
                    $onedrivesites += $odsitestring
                }
                #write-host "OneDrive is enabled for Members of:"
            }
    }

    #Teams Activity Block
    if ($policy.TeamsLocation.Count -gt 0){
    write-host $policy.name "Policy is Enabled for Teams" -ForegroundColor Yellow

        if($policy.TeamsLocation.name -eq "All"){
            $teamslocations += "All Teams"
            $teamslocationscount = "All Users"
        }
    
        else{
            foreach($teamitem in $policy.TeamsLocation) {
                try{$teamuser = Get-AzureADUser -ObjectId $teamitem.immutableidentity -ErrorAction SilentlyContinue}
                catch{$teamgroup = get-azureadgroup -ObjectId $teamitem.immutableidentity -ErrorAction SilentlyContinue}
            
                if ($teamuser.ObjectId -eq $teamitem.immutableidentity){
                    $usertext = "User: " 
                    $teamdata = $usertext + $teamitem.Displayname
                    $teamslocations += $teamdata
                    $teamslocationscount++
                }
                elseif($teamgroup.Objectid -eq $teamitem.immutableidentity){
                    #get the count of the members of each team group
                    try{$teamgroupcount = (Get-AzureADGroupMember -ObjectId $teamgroup.objectid -all $true).Count}
                    catch{$teamgroupcount = 0}
                    
                    $grouptext = "Group: " 
                    $teamdata = $grouptext + $teamitem.Displayname + " (" + $teamgroupcount + ")"                    
                    $teamslocations += $teamdata
                    $teamslocationscount += $teamgroupcount
                }
            }
        }        
        #write-host "Teams is enabled for" $teamslocations
    }

    #EndpointDLP Check
    if ($policy.EndpointDlpLocation.Count -gt 0){
        write-host $policy.name "Policy is Enabled for EndpointDLP" -ForegroundColor Yellow
        
        if($policy.EndpointDlpLocation.name -eq "All"){
            $endpointdlplocations += "All Enrolled Endpoints"
            $endpointdlplocationcounts = "All Users"
            #write-host "Endpoint DLP is enabled for" $endpointdlplocations
        }

        else{
            foreach($endpointitem in $policy.EndpointDlpLocation) {
                try{$endpointuser = Get-AzureADUser -ObjectId $endpointitem.immutableidentity -ErrorAction SilentlyContinue}
                catch{$endpointgroup = get-azureadgroup -ObjectId $endpointitem.immutableidentity -ErrorAction SilentlyContinue}
            
                if ($endpointuser.ObjectId -eq $endpointitem.immutableidentity){
                    $usertext = "User: " 
                    $endpointdata = $usertext + $endpointitem.Displayname
                    $endpointdlplocations += $endpointdata
                    $endpointdlplocationcounts++
                }
                elseif($endpointgroup.Objectid -eq $endpointitem.immutableidentity){
                    #get the count of the members of each team group
                    try{$endpointgroupcount = (Get-AzureADGroupMember -ObjectId $endpointgroup.objectid -all $true).Count}
                    catch{$endpointgroupcount = 0}
                    $grouptext = "Group: " 
                    $endpointdata = $grouptext + $endpointitem.Displayname + " (" + $endpointgroupcount + ")"
                    $endpointdlplocations += $endpointdata
                    $endpointdlplocationcounts += $endpointgroupcount 
                }
            }
        }
            
        #write-host "Endpoint DLP is enabled for" $endpointdlplocations
    }

    #MDCA Check
    if ($policy.ThirdPartyAppDlpLocation.Count -gt 0){
        write-host $policy.name "Policy is Enabled for Defender for Cloud Apps" -ForegroundColor Yellow
        $defenderforCAlocations = "Enabled"
    }

    #AIP Scanner Location
    if ($policy.OnPremisesScannerDlpLocation.Count -gt 0){
        write-host $policy.name "Policy is Enabled for On Premisies Locations" -ForegroundColor Yellow

        if($policy.OnPremisesScannerDlpLocation.name -eq "All"){
            $onpremlocations += "All Repositories"
            $onpremlocationscount 
        }
        else{
            foreach($opremlocation in $policy.OnPremisesScannerDlpLocation){
                $onpremlocations += $opremlocation.Displayname
                $onpremlocationscount++
            }
        }
        #write-host "On-Prem DLP is enabled for" $onpremlocations    
    }

    $dlphashtable=[Ordered]@{
        PolicyName = $param
        Exchange = ($exchangemembers -join ":::") | Out-String
        ProtectedExchangeUsers = $exchangememberscount
        OneDriveUsers = ($onedrivemembers -join ":::")| Out-String
        OneDriveGroups = ($onedrivesites -join ":::") | Out-String
        ProtectedOneDriveLocations = $onedrivememberscount + $onedrivesitescount
        SharePoint = ($sharepointlocations -join ":::") | Out-String
        ProtectedSharePointLocations = $sharepointlocationscount
        Teams = ($teamslocations -join ":::") | Out-String
        ProtectedTeamsUsers = $teamslocationscount
        Endpoints = ($endpointdlplocations -join ":::") | Out-String
        ProtectedEndPointUsers = $endpointdlplocationcounts
        DefenderforCA = ($defenderforCAlocations -join ":::") | Out-String
        OnPremDLP = ($onpremlocations -join ":::") | Out-String
        ProtectedOnPremLocations = $onpremlocationscount
    }

    $dlppolicydetailobject = [PSCustomObject]$dlphashtable    
    $dlppolicydetailtable += $dlppolicydetailobject    
    return $dlppolicydetailtable
}
function get-dlppolicyruledetails($param){
    $dlppolicyrule = Get-DlpCompliancerule  -Policy $param
    $sitlist = @()
    $labellist = @()
    
    foreach($rule in $dlppolicyrule){
        $rulename = $rule.DistinguishedName.split(",")[0] -replace 'CN='
        $ruledisabled

        if($rule.Disabled -like 'True'){
            $ruledisabled="No"
        }
        else{
            $ruledisabled="Yes"
        }
        ### new content to deal with advaned rules

        if($rule.IsAdvancedRule -like 'True'){
            #### do all the code for the advanced processing
            $rulejson = $rule.AdvancedRule | convertfrom-json | Select-Object -ExpandProperty Condition | Select-Object -ExpandProperty SubConditions

            foreach ($ruleitem in $rulejson){
                if ($ruleitem.ConditionName -like "ContentContainsSensitiveInformation"){
                    $sitgroups = $ruleitem | Select-Object -ExpandProperty Value | Select-Object -ExpandProperty Groups 

                    foreach($sitgroup in $sitgroups){
                    $sitruledetails = $sitgroup | Select-Object -ExpandProperty SensitiveTypes

                        foreach ($sitrule in $sitruledetails){
                            $shash = [ordered]@{
                            RuleGroup = $rulename
                            Name = $sitrule.name
                            RuleEnabled = $ruledisabled
                            ClassifierType = if($sitrule.ClassifierType){$sitrule.ClassifierType}else{"Content"}
                            MinCount = $sitrule.Mincount
                            MaxCount = $sitrule.Maxcount
                            ConfidenceLevel = $sitrule.Confidencelevel
                            }
                        $sobject = New-Object PSObject -Property $shash
                        $sitlist += $sobject
                        }
                    
                    }

                }

            }

        }

        elseif($null -ne $rule.ContentContainsSensitiveInformation){
        
            #check if more than one group of Sit's exsits in the policy
            if($null -ne $rule.ContentContainsSensitiveInformation.groups){
                $sitgroup = $rule.ContentContainsSensitiveInformation.groups
                
                foreach($group in $sitgroup){
                 $sensitivetypes = $group.sensitivetypes
                 $labels = $group.labels
                     foreach($sitg in $sensitivetypes){
                         $shash = [ordered]@{
                            RuleGroup = $rulename
                            Name = $sitg.name
                            RuleEnabled = $ruledisabled
                            ClassifierType = $sitg.ClassifierType
                            MinCount = $sitg.mincount
                            MaxCount = $sitg.maxcount
                            ConfidenceLevel = $sitg.confidencelevel
                         }
                         $sobject = New-Object PSObject -Property $shash
                         $sitlist += $sobject
                     }
                     foreach($label in $labels){
                        $lhash = [ordered]@{
                            RuleGroup = $rulename
                            LabelName = (get-label $label.name).DisplayName
                        }
                         $labellist += $lhash
                     }   
                }
             }
                 else{
                 $sits = $rule.ContentContainsSensitiveInformation
                    foreach($sit in $sits){
                        $shash = [ordered]@{
                            RuleGroup = $rulename
                            Name = $sit.name
                            RuleEnabled = $ruledisabled
                            ClassifierType = $sit.ClassifierType
                            MinCount = $sit.mincount
                            MaxCount = $sit.maxcount
                            ConfidenceLevel = $sit.confidencelevel
                         }
                         $sobject = New-Object PSObject -Property $shash
                         $sitlist += $sobject
                     }
                 }
        }      
    }
    
    return $sitlist
    }
function get-auditlogsummary(){
    $audittable = @()
    $auditlogpolicies = Get-UnifiedAuditLogRetentionPolicy
    
    foreach ($logpolicy in $auditlogpolicies){
        $audithashtable = @{}
    
        if ($logpolicy.operations -gt 0){
            $opslist = ($logpolicy.operations).split(",") 
        }
        else {
            $opslist = $null
        }
    
        $audithashtable= [Ordered]@{
            Priority = $logpolicy.Priority
            Name = $logpolicy.name
            WhenCreated = $logpolicy.whencreated
            Enabled = $logpolicy.enabled
            RecordType = ($logpolicy.recordtypes -join ":::") | Out-String
            Operations = ($opslist -join ":::") | Out-String
            UsersAudited = ($logpolicy.UserIds).split(",").count
            RetentionDuration = $logpolicy.RetentionDuration
            
        }
    
        #create the new object
        $auditlogpolicyobject = [PSCustomObject]$audithashtable
        $audittable += $auditlogpolicyobject
                
            
    }
    
    return $audittable
}


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

if (get-installedmodule -Name AzureADPreview -ErrorAction SilentlyContinue) {
    Write-Host "Azure AD Module Installed, Continuing with Script Execution"
}
else {
    $title    = 'Azure AD Module is Not Installed'
    $question = 'Do you want to install it now?'
    $choices  = '&Yes', '&No'

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) {
        Write-Host 'Your choice is Yes, installing module'
        Install-Module AzureADPreview -SkipPublisherCheck -Force -Confirm:$false 
    } else {
        Write-Host 'Please install the module manually to continue https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0'
        Exit
}
}

Write-Host 'Connecting to Security & Compliance Center. Please logon in the new window' -ForegroundColor DarkYellow
Connect-IPPSSession
Write-Host "Connecting to Azure AD. Please Logon in the new Window" -ForegroundColor DarkYellow
Connect-AzureAD
Write-Host "`r`n`r`nConnected to Microsoft 365, Continuing with Script`r`n`r`n" -ForegroundColor Yellow

#######################
#script activities
#each section below performs one part of the script
#section 1: collects all of the DLP Policies and identifies which workloads are enabled
#section 2: works through each individual policy and gathers pertninant data about each DLP Policy including
#           the locations (both users and groups) that are being evaulated.  It also goes through the Rules 
#           attached to each policy and lists out any SITs that are being evaluated.  There are a number of
#           other policy and rule settings that COULD be pulled, in this script we are focused on the settings
#           that we are configuring as part of the workshops
#section 3: we are creating a unified summary table for the top fo the report. This uses the chart created in 
#           section 1 and then combines that with a count of all of unique SITS in each policy along with a 
#           count of all of the covered users in a given workload combining individually defined users and
#           users in a group
#section 4: this is where we are constructing the report itself.  It involves merging data from the prior two
#           sections and then using convertto-html to place it all into a report that can be provided to the
#           customer and submitted as part of the final Proof of Execution (POE) for the workshop
#######################

### section 1 DLP Policies
$dlppolicysummary = get-dlpolicysummary

### section 2, collect information on individual policies
$dlppolicies = Get-DlpCompliancePolicy | Select-Object Name

foreach($dlppolicy in $dlppolicies){
    #get the policy details
    $dlpdetails = get-dlppolicydetails($dlppolicy.name)
        
    #get the rule details
    $dlpruledetails = Get-DlpComplianceRule -Policy $dlppolicy.name
    
    $htmloutput=@()
    $sitcountperrule=@()
    
    foreach($rule in $dlpruledetails){
        Write-Host "Processing DLP Rule:" $rule.Name "in Policy" $dlppolicy.name -ForegroundColor Green

        if($rule.Disabled -like 'True'){
            $ruledisabled="No"
        }
        else{
            $ruledisabled="Yes"
        }
        
        $rulename = $rule.DistinguishedName.split(",")[0] -replace 'CN='
        $adminalert = $rule.GenerateAlert
        $alertthreshold = $rule.AlertProperties.threshold
        $notifyuser = ($null -ne $rule.NotifyUser) -or ($null -ne $rule.NotifyEndpointUser)
        $advancedrule = if($rule.IsAdvancedRule -like "True"){"Yes"}else{"No"}
        
        $allruleschart = get-dlppolicyruledetails ($dlppolicy.name) 
        $rulechart = $allruleschart | Where-Object {$_.RuleGroup -eq $rulename} | Select-Object Name,RuleEnabled,classifiertype,mincount,maxcount,confidencelevel
        $rulehtml = ($rulechart | convertto-html -Fragment) -replace ("-1","Any") 

        ####get a count of sits by policy name
        #count the number of times a specific SIT appears in a policy
        $sitcount = @($allruleschart | where-object {$_.RuleEnabled -eq "Yes"} | Select-Object Name -Unique).count
        $sitcounthash = [PSCustomObject]@{
            DLPPolicyName = $dlppolicy.Name
            DLPRUleName = $rulename
            CountofSits = $sitcount
        }
        $sitcountperrule += $sitcounthash

        $htmloutput += "<p> <b>Rule Name:</b> $($rulename)<br>
        <b>Rule Enabled:</b> $($ruledisabled)<br>
        <b>Advanced Rule (Complex Conditions):</b> $($advancedrule)<br>
        <b>Admin Alerts:</b> $($adminalert)<br>
        <b>Alert Threshold:</b> $($alertthreshold)<br>
        <b>User Notification:</b> $($notifyuser)</p>"
        $htmloutput += $rulehtml

    }

    ##get a unique count for the sit counts
    $sitcounts += $sitcountperrule | Select-Object DLPPolicyName,CountofSits -Unique

    #capture up all of the policy counts for user later
    $policycounts += ($dlpdetails | Select-Object PolicyName,ProtectedExchangeUsers,ProtectedOnedriveLocations,ProtectedSharePointLocations,ProtectedTeamsUsers,ProtectedEndpointUsers,ProtectedOnPremLocations)

    #build out the html file to user later
    $chartbuild = ($dlpdetails |Select-Object Exchange,onedriveusers,onedrivegroups,Sharepoint,Teams,endpoints,defenderforca,OnPremDLP | convertto-html  -Fragment -PreContent "<h4>Protected Users and Groups by Workload</h4>") -replace (":::","<br/>")
    $a += "<h3><b>$($dlppolicy.name)</b></h3>"
    $a += $chartbuild
    $a += "<h4><i>Policy Rule Snapshot</i></h4>"
    $a += $htmloutput
    $a += "<br>"
    $a += "<hr>"
}

### section 3, gather the audit log configuration
$auditlogsummary = get-auditlogsummary
$auditchart = ($auditlogsummary | ConvertTo-Html -Fragment -PreContent "<h2>Unified Audit Log Policy Summary</h2>") -replace (":::","<br/>")
$auditchart += "<hr>"

### section ,3 construct a unified summary table
###### doing this here as i couldn't get the function to return properly
###### future improvement - create a function that takes these 3 items and returns the formatted data
$dlpsummarychart = $dlppolicysummary
$dlppolicycounts = $policycounts
$sitpolicycounts = $sitcounts
$POEChart = [array]@()

foreach ($item in $dlpsummarychart){

    $coveredaccounts = $dlppolicycounts | Where-Object {$_.Policyname -eq $item.policyname}
    $coveredsits = $sitpolicycounts | Where-Object {$_.DLPPolicyName -eq $item.policyname}

    #create the new output hashtable
    $itemtable=[ordered]@{
        DLPPolicyName = $item.PolicyName
        CreationDate = $item.CreationDate
        PolicyMode = $item.PolicyMode
        SITSUsed = $coveredsits.CountofSits
        ExchangeOnline = $item.ExchangeOnline + " (" + $coveredaccounts.ProtectedExchangeUsers + ")"
        OneDrive = $item.OneDrive + " (" + $coveredaccounts.ProtectedOneDriveLocations + ")"
        SharePoint = $item.SharePointOnline + " (" + $coveredaccounts.ProtectedSharePointLocations + ")"
        Teams = $item.Teams + " (" + $coveredaccounts.ProtectedTeamsUsers + ")"
        Endpoints = $item.EndPoints + " (" + $coveredaccounts.ProtectedEndPointUsers + ")"
        DefenderforCA = $item.DefenderforCA 
        OnPremises = $item.OnPremises + " (" + $coveredaccounts.ProtectedOnPremLocations + ")"
    }

    $summarychart = [PSCustomObject]$itemtable
    $POEChart += $summarychart
}

###section 4, build our html file
$tenantdetails = Get-AzureADTenantDetail
$scriptrunner = Get-AzureADCurrentSessionInfo

$reportstamp = "<p id='CreationDate'><b>Report Creation Date:</b> $(Get-Date)<br>
<b>Tenant Name:</b> $($tenantdetails.DisplayName)<br>
<b>Tenant ID:</b> $($scriptrunner.TenantID)<br>
<b>Tenant Domain:</b> $($scriptrunner.TenantDomain)<br>
<b>Executed by</b>: $($scriptrunner.Account)</p>
"

$reportintro = "<h1> Compliance Workshop: Policy Configuration Report</h1>
<p><b>The following report shows a snapshot of the current status of Audit and DLP Policy Configuration within the Microsoft 365 environment.</b> </p>
<p>Units in <b>()</b> indicate the number of protected users or sites </p>"
$reportintro+= $reportstamp


$poeintro = "<h1> Compliance Workshop: Policy Configuration Report (POE Only)</h1>
<p><b>The following report shows a snapshot of the current status of Audit and DLP Policy Configuration within the Microsoft 365 environment.</b> </p>
<p>Units in <b>()</b> indicate the number of protected users or sites </p>"
$poeintro += $reportstamp


$reportdetails = "<h2>Individual Policy Details<h2>
</hr2>"

if($reporttype -match 'All'){
    $poehtml = ($poechart | ConvertTo-Html -PreContent "<h2>Compliance Workshop DLP POE Summary</h2>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poethml += "<hr>"
    #saving each of the individual reports here in case they are ever needed for troubleshooting the report rollup
    #$summaryhtml = $dlppolicysummary | ConvertTo-Html -Fragment
    #$policyhtml = $policycounts | convertto-html -Fragment
    #$sithtml = $sitcounts | ConvertTo-Html -Fragment
    Convertto-html -Head $header -Body "$reportintro $poehtml $auditchart $reportdetails $a" -Title "Compliance Workshop Policy Configuration Report" | Out-File $outputfile 
}
elseif($reporttype -match'POEReport'){
    $poehtml = ($poechart | ConvertTo-Html -PreContent "<h2>Compliance Workshop DLP POE Summary</h2>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.Exchangeonline -like "Yes*" -and $_.PolicyMode -match "Enable"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,SITSUsed,ExchangeOnline | ConvertTo-Html -PreContent "<h3>Exchange DLP Policies</h3>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.OneDrive -like "Yes*" -and $_.PolicyMode -match "Enable"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,SITSUsed,OneDrive | ConvertTo-Html -PreContent "<h3>OneDrive DLP Policies</h3>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.SharePoint -like "Yes*" -and $_.PolicyMode -match "Enable"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,SITSUsed,SharePoint | ConvertTo-Html -PreContent "<h3>SharePoint DLP Policies</h3>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += ($poechart | Where-Object {$_.Teams -like "Yes*" -and $_.PolicyMode -match "Enable"} | Select-Object DLPPolicyName,CreationDate,PolicyMode,SITSUsed,Teams | ConvertTo-Html -PreContent "<h3>Teams DLP Policies</h3>") -replace ("(\([0]\))","") -replace ("(s\d+\))","s)")
    $poehtml += "<hr>"

    Convertto-html -Head $header -Body "$poeintro $poehtml $auditchart" -Title "Compliance Workshop POE Report" | Out-File $outputfile 
}

#display report in browser
Write-Host "`nReport file available at:" $outputfile -ForegroundColor Yellow -BackgroundColor Blue
Write-host "`n`r"
Start-Process $outputfile

#cleanup
Write-Host "Disconnecting Services" -ForegroundColor Yellow
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction:SilentlyContinue  -InformationAction Ignore
Disconnect-AzureAD -Confirm:$false -InformationAction Ignore