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
$FriendlyLicenses= @{}
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

#list of all of the current friendly sku product names - https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference 
# how to build below list,  
# 1) download file
# 2),  import-csv into variable, 
# 3) <variable> | select Service_Plan_Name, Service_Plans_Included_Friendly_Names -Unique | Export-Csv <some csv list>
# 4) remove duplicates of Service Plan Names based on which one is the most readable
$FriendlyLicenses = @{
    "AAD_BASIC"="MICROSOFT AZURE ACTIVE DIRECTORY BASIC"
    "AAD_BASIC_EDU"="Azure Active Directory Basic for Education"
    "AAD_EDU"="Azure Active Directory for Education"
    "AAD_PREMIUM"="Azure Active Directory Premium Plan 1"
    "AAD_PREMIUM_P2"="Azure Active Directory Premium P2"
    "AAD_SMB"="Azure Active Directory"
    "ADALLOM_FOR_AATP"="SecOps Investigation for MDI"
    "ADALLOM_S_DISCOVERY"="Microsoft Defender for Cloud Apps Discovery"
    "ADALLOM_S_O365"="Office 365 Cloud App Security"
    "ADALLOM_S_STANDALONE"="Microsoft Defender for Cloud Apps"
    "ADALLOM_S_STANDALONE_DOD"="Microsoft Defender for Cloud Apps for DOD"
    "ATA"="Microsoft Defender for Identity"
    "ATP_ENTERPRISE"="Microsoft Defender for Office 365 (Plan 1)"
    "ATP_ENTERPRISE_GOV"="Microsoft Defender for Office 365 (Plan 1) for Government"
    "BI_AZURE_P_2_GOV"="Power BI Pro for Government"
    "BI_AZURE_P0"="Power BI (free)"
    "BI_AZURE_P1"="Microsoft Power BI Reporting and Analytics Plan 1"
    "BI_AZURE_P2"="Power BI Pro"
    "BI_AZURE_P3"="Power BI Premium Per User"
    "BPOS_S_DlpAddOn"="Data Loss Prevention"
    "BPOS_S_TODO_1"="To-Do (Plan 1)"
    "BPOS_S_TODO_2"="To-Do (Plan 2)"
    "BPOS_S_TODO_3"="To-Do (Plan 3)"
    "BPOS_S_TODO_FIRSTLINE"="To-Do (Firstline)"
    "CCIBOTS_PRIVPREV_VIRAL"="Dynamics 365 AI for Customer Service Virtual Agents Viral"
    "CDS_ATTENDED_RPA"="Common Data Service Attended RPA"
    "CDS_CUSTOMER_INSIGHTS_TRIAL"="Common Data Service for Customer Insights Trial"
    "CDS_DB_CAPACITY"="Common Data Service for Apps Database Capacity"
    "CDS_DB_CAPACITY_GOV"="Common Data Service for Apps Database Capacity for Government"
    "CDS_Flow_Business_Process"="Common data service for Flow per business process plan"
    "CDS_FORM_PRO_USL"="Common Data Service"
    "CDS_LOG_CAPACITY"="Common Data Service for Apps Log Capacity"
    "CDS_O365_E5_KM"="Common Data Service for SharePoint Syntex"
    "CDS_O365_F1"="Common Data Service for Teams"
    "CDS_O365_F1_GCC"="Common Data Service for Teams_F1 GCC"
    "CDS_O365_P1"="Common Data Service for Teams"
    "CDS_O365_P1_GCC"="Common Data Service for Teams_P1 GCC"
    "CDS_O365_P2"="Common Data Service for Teams_P2"
    "CDS_O365_P2_GCC"="COMMON DATA SERVICE FOR TEAMS_P2 GCC"
    "CDS_O365_P3"="Common Data Service for Teams_P3"
    "CDS_O365_P3_GCC"="CDS_O365_P3_GCC"
    "CDS_PER_APP"="CDS PowerApps per app plan"
    "CDS_PER_APP_IWTRIAL"="CDS Per app baseline access"
    "CDS_POWERAPPS_PORTALS_LOGIN"="Common Data Service Power Apps Portals Login Capacity"
    "CDS_POWERAPPS_PORTALS_LOGIN_GCC"="Common Data Service Power Apps Portals Login Capacity for GCC"
    "CDS_POWERAPPS_PORTALS_PAGEVIEW_GCC"="CDS PowerApps Portals page view capacity add-on for GCC"
    "CDS_REMOTE_ASSIST"="Common Data Service for Remote Assist"
    "CDS_UNATTENDED_RPA"="Common Data Service Unattended RPA"
    "CDS_VIRTUAL_AGENT_BASE"="Common Data Service for Virtual Agent Base"
    "CDSAICAPACITY"="AI Builder capacity add-on"
    "CDSAICAPACITY_PERAPP"="AI Builder capacity Per App add-on"
    "CDSAICAPACITY_PERUSER"="AI Builder capacity Per User add-on"
    "CDSAICAPACITY_PERUSER_NEW"="AI Builder capacity Per User add-on"
    "COMMUNICATIONS_COMPLIANCE"="Microsoft Communications Compliance"
    "COMMUNICATIONS_DLP"="Microsoft Communications DLP"
    "COMPLIANCE_MANAGER_PREMIUM_ASSESSMENT_ADDON"="Compliance Manager Premium Assessment Add-On"
    "Content_Explorer"="Information Protection and Governance Analytics - Premium"
    "ContentExplorer_Standard"="Information Protection and Governance Analytics - Standard"
    "CORTEX"="Viva Topics"
    "CPC_1"="Windows 365 Enterprise 2 vCPU 4 GB 128 GB"
    "CPC_2"="Windows 365 Enterprise 2 vCPU 8 GB 128 GB"
    "CPC_B_1C_2RAM_64GB"="Windows 365 Business 1 vCPU 2 GB 64 GB"
    "CPC_B_2C_4RAM_128GB"="Windows 365 Business 2 vCPU 4 GB 128 GB"
    "CPC_B_2C_4RAM_256GB"="Windows 365 Business 2 vCPU 4 GB 256 GB"
    "CPC_B_2C_4RAM_64GB"="Windows 365 Business 2 vCPU 4 GB 64 GB"
    "CPC_B_2C_8RAM_256GB"="Windows 365 Business 2 vCPU 8 GB 256 GB"
    "CPC_B_4C_16RAM_128GB"="Windows 365 Business 4 vCPU 16 GB 128 GB"
    "CPC_B_4C_16RAM_256GB"="Windows 365 Business 4 vCPU 16 GB 256 GB"
    "CPC_B_4C_16RAM_512GB"="Windows 365 Business 4 vCPU 16 GB 512 GB"
    "CPC_B_8C_32RAM_128GB"="Windows 365 Business 8 vCPU 32 GB 128 GB"
    "CPC_B_8C_32RAM_256GB"="Windows 365 Business 8 vCPU 32 GB 256 GB"
    "CPC_B_8C_32RAM_512GB"="Windows 365 Business 8 vCPU 32 GB 512 GB"
    "CPC_E_1C_2GB_64GB"="Windows 365 Enterprise 1 vCPU 2 GB 64 GB"
    "CPC_E_2C_4GB_256GB"="Windows 365 Enterprise 2 vCPU 4 GB 256 GB"
    "CPC_E_2C_4GB_64GB"="Windows 365 Enterprise 2 vCPU 4 GB 64 GB"
    "CPC_E_2C_8GB_256GB"="Windows 365 Enterprise 2 vCPU 8 GB 256 GB"
    "CPC_E_4C_16GB_128GB"="Windows 365 Enterprise 4 vCPU 16 GB 128 GB"
    "CPC_E_4C_16GB_256GB"="Windows 365 Enterprise 4 vCPU 16 GB 256 GB"
    "CPC_E_4C_16GB_512GB"="Windows 365 Enterprise 4 vCPU 16 GB 512 GB"
    "CPC_E_8C_32GB_128GB"="Windows 365 Enterprise 8 vCPU 32 GB 128 GB"
    "CPC_E_8C_32GB_256GB"="Windows 365 Enterprise 8 vCPU 32 GB 256 GB"
    "CPC_E_8C_32GB_512GB"="Windows 365 Enterprise 8 vCPU 32 GB 512 GB"
    "CPC_S_2C_4GB_128GB"="Windows 365 Shared Use 2 vCPU 4 GB 128 GB"
    "CPC_S_2C_4GB_256GB"="Windows 365 Shared Use 2 vCPU 4 GB 256 GB"
    "CPC_S_2C_4GB_64GB"="Windows 365 Shared Use 2 vCPU 4 GB 64 GB"
    "CPC_S_2C_8GB_128GB"="Windows 365 Shared Use 2 vCPU 8 GB 128 GB"
    "CPC_S_2C_8GB_256GB"="Windows 365 Shared Use 2 vCPU 8 GB 256 GB"
    "CPC_S_4C_16GB_128GB"="Windows 365 Shared Use 4 vCPU 16 GB 128 GB"
    "CPC_S_4C_16GB_256GB"="Windows 365 Shared Use 4 vCPU 16 GB 256 GB"
    "CPC_S_4C_16GB_512GB"="Windows 365 Shared Use 4 vCPU 16 GB 512 GB"
    "CPC_S_8C_32GB_128GB"="Windows 365 Shared Use 8 vCPU 32 GB 128 GB"
    "CPC_S_8C_32GB_256GB"="Windows 365 Shared Use 8 vCPU 32 GB 256 GB"
    "CPC_S_8C_32GB_512GB"="Windows 365 Shared Use 8 vCPU 32 GB 512 GB"
    "CPC_SS_2"="Windows 365 Business 2 vCPU, 8 GB, 128 GB"
    "CRM_ONLINE_PORTAL"="Microsoft Dynamics CRM Online - Portal Add-On"
    "CRMINSTANCE"="Microsoft Dynamics CRM Online Instance"
    "CRMPLAN2"="MICROSOFT DYNAMICS CRM ONLINE BASIC"
    "CRMSTANDARD"="MICROSOFT DYNAMICS CRM ONLINE PROFESSIONA"
    "CRMSTORAGE"="Microsoft Dynamics CRM Online Storage Add-On"
    "CRMTESTINSTANCE"="Microsoft Dynamics CRM Online Additional Test Instance"
    "CUSTOMER_KEY"="Microsoft Customer Key"
    "CUSTOMER_VOICE_ADDON"="Dynamics Customer Voice Add-On"
    "Customer_Voice_Base"="Dynamics 365 Customer Voice Base Plan"
    "CUSTOMER_VOICE_DYN365_VIRAL_TRIAL"="Customer Voice for Dynamics 365 vTrial"
    "D365_AssetforSCM"="Asset Maintenance Add-in"
    "D365_CSI_EMBED_CE"="Dynamics 365 Customer Service Insights for CE Plan"
    "D365_CSI_EMBED_CSEnterprise"="Dynamics 365 Customer Service Insights for CS Enterprise"
    "D365_FIELD_SERVICE_ATTACH"="Dynamics 365 for Field Service Attach"
    "D365_Finance"="Microsoft Dynamics 365 for Finance"
    "D365_IOTFORSCM"="Iot Intelligence Add-in for D365 Supply Chain Management"
    "D365_IOTFORSCM_ADDITIONAL"="IoT Intelligence Add-in Additional Machines"
    "D365_ProjectOperations"="Dynamics 365 Project Operations"
    "D365_ProjectOperationsCDS"="Dynamics 365 Project Operations CDS"
    "D365_SALES_ENT_ATTACH"="Dynamics 365 for Sales Enterprise Attach"
    "D365_SALES_PRO_ATTACH"="Dynamics 365 for Sales Pro Attach"
    "D365_SALES_PRO_IW"="Dynamics 365 for Sales Professional Trial"
    "D365_SALES_PRO_IW_Trial"="Dynamics 365 for Sales Professional Trial"
    "D365_SCM"="DYNAMICS 365 FOR SUPPLY CHAIN MANAGEMENT"
    "DATA_INVESTIGATIONS"="Microsoft Data Investigations"
    "DATAVERSE_FOR_POWERAUTOMATE_DESKTOP"="Dataverse for PAD"
    "DATAVERSE_POWERAPPS_PER_APP_NEW"="Dataverse for Power Apps per app"
    "DDYN365_CDS_DYN_P2"="COMMON DATA SERVICE"
    "Deskless"="Microsoft StaffHub"
    "DYN365_AI_SERVICE_INSIGHTS"="Dynamics 365 AI for Customer Service Trial"
    "DYN365_BUSCENTRAL_DB_CAPACITY"="Dynamics 365 Business Central Database Capacity"
    "DYN365_BUSCENTRAL_ENVIRONMENT"="Dynamics 365 Business Central Additional Environment Addon"
    "DYN365_BUSCENTRAL_PREMIUM"="Dynamics 365 Business Central Premium"
    "DYN365_BUSINESS_Marketing"="Dynamics 365 Marketing"
    "DYN365_CDS_CCI_BOTS"="Common Data Service for CCI Bots"
    "DYN365_CDS_DEV_VIRAL"="Common Data Service - DEV VIRAL"
    "DYN365_CDS_DYN_APPS"="Common Data Service"
    "DYN365_CDS_FINANCE"="Common Data Service for Dynamics 365 Finance"
    "DYN365_CDS_FOR_PROJECT_P1"="Common Data Service for Project P1"
    "DYN365_CDS_FORMS_PRO"="Common Data Service"
    "DYN365_CDS_GUIDES"="Common Data Service"
    "DYN365_CDS_O365_F1"="Common Data Service - O365 F1"
    "DYN365_CDS_O365_F1_GCC"="Common Data Service - O365 F1"
    "DYN365_CDS_O365_P1"="Common Data Service - O365 P1"
    "DYN365_CDS_O365_P1_GCC"="Common Data Service - O365 P1 GCC"
    "DYN365_CDS_O365_P2"="Common Data Service - O365 P2"
    "DYN365_CDS_O365_P2_GCC"="COMMON DATA SERVICE - O365 P2 GCC"
    "DYN365_CDS_O365_P3"="Common Data Service - O365 P3"
    "DYN365_CDS_P1_GOV"="Common Data Service for Government"
    "DYN365_CDS_P2"="Common Data Service - P2"
    "DYN365_CDS_P2_GOV"="Common Data Service for Government"
    "DYN365_CDS_PROJECT"="Common Data Service for Project"
    "DYN365_CS_ENTERPRISE_VIRAL_TRIAL"="Dynamics 365 Customer Service Enterprise vTrial"
    "DYN365_CS_MESSAGING_VIRAL_TRIAL"="Dynamics 365 Customer Service Digital Messaging vTrial"
    "DYN365_CS_VOICE_VIRAL_TRIAL"="Dynamics 365 Customer Service Voice vTrial"
    "DYN365_CUSTOMER_INSIGHTS_ENGAGEMENT_INSIGHTS_BASE_TRIAL"="Dynamics 365 Customer Insights Engagement Insights Viral"
    "DYN365_CUSTOMER_INSIGHTS_VIRAL"="Dynamics 365 Customer Insights Viral Plan"
    "DYN365_CUSTOMER_SERVICE_PRO"="Dynamics 365 for Customer Service Pro"
    "DYN365_ENTERPRISE_CASE_MANAGEMENT"="Dynamics 365 for Case Management"
    "DYN365_ENTERPRISE_CUSTOMER_SERVICE"="MICROSOFT SOCIAL ENGAGEMENT - SERVICE DISCONTINUATION"
    "DYN365_ENTERPRISE_FIELD_SERVICE"="Dynamics 365 for Field Service"
    "DYN365_ENTERPRISE_P1"="DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN"
    "DYN365_ENTERPRISE_P1_IW"="DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS"
    "DYN365_ENTERPRISE_SALES"="DYNAMICS 365 FOR SALES"
    "DYN365_ENTERPRISE_TALENT_ATTRACT_TEAMMEMBER"="DYNAMICS 365 FOR TALENT - ATTRACT EXPERIENCE TEAM MEMBER"
    "DYN365_ENTERPRISE_TALENT_ONBOARD_TEAMMEMBER"="DYNAMICS 365 FOR TALENT - ONBOARD EXPERIENCE"
    "DYN365_ENTERPRISE_TEAM_MEMBERS"="DYNAMICS 365 FOR TEAM MEMBERS"
    "DYN365_FINANCIALS_ACCOUNTANT"="Dynamics 365 Business Central External Accountant"
    "DYN365_FINANCIALS_BUSINESS"="Dynamics 365 for Business Central Essentials"
    "DYN365_FINANCIALS_TEAM_MEMBERS"="Dynamics 365 for Team Members"
    "DYN365_FS_ENTERPRISE_VIRAL_TRIAL"="Dynamics 365 Field Service Enterprise vTrial"
    "DYN365_MARKETING_MSE_USER"="Dynamics 365 for Marketing MSE User"
    "DYN365_MARKETING_USER"="Dynamics 365 for Marketing USL"
    "DYN365_REGULATORY_SERVICE"="Dynamics 365 for Finance and Operations Enterprise edition - Regulatory Service"
    "DYN365_RETAIL_DEVICE"="Dynamics 365 for Retail Device"
    "DYN365_SALES_ENTERPRISE_VIRAL_TRIAL"="Dynamics 365 Sales Enterprise vTrial"
    "DYN365_SALES_INSIGHTS_VIRAL_TRIAL"="Dynamics 365 Sales Insights vTrial"
    "DYN365_SALES_PRO"="Dynamics 365 for Sales Professional"
    "DYN365_TALENT_ENTERPRISE"="DYNAMICS 365 FOR TALENT"
    "DYN365_TEAM_MEMBERS"="DYNAMICS 365 TEAM MEMBERS"
    "DYN365BC_MS_INVOICING"="Microsoft Invoicing"
    "Dynamics_365_for_HCM_Trial"="Dynamics 365 for HCM Trial"
    "Dynamics_365_for_Operations"="DYNAMICS 365 FOR_OPERATIONS"
    "Dynamics_365_for_Operations_Sandbox_Tier2"="Dynamics 365 for Operations non-production multi-box instance for standard acceptance testing (Tier 2)"
    "Dynamics_365_for_Operations_Sandbox_Tier4"="Dynamics 365 for Operations Enterprise Edition - Sandbox Tier 4:Standard Performance Testing"
    "DYNAMICS_365_FOR_OPERATIONS_TEAM_MEMBERS"="DYNAMICS 365 FOR OPERATIONS TEAM MEMBERS"
    "Dynamics_365_for_OperationsDevices"="Dynamics 365 for Operations Devices"
    "Dynamics_365_for_Retail"="DYNAMICS 365 FOR RETAIL"
    "DYNAMICS_365_FOR_RETAIL_TEAM_MEMBERS"="DYNAMICS 365 FOR RETAIL TEAM MEMBERS"
    "DYNAMICS_365_FOR_TALENT_TEAM_MEMBERS"="DYNAMICS 365 FOR TALENT TEAM MEMBERS"
    "DYNAMICS_365_HIRING_FREE_PLAN"="DYNAMICS 365 HIRING FREE PLAN"
    "Dynamics_365_Onboarding_Free_PLAN"="Dynamics 365 for Talent: Onboard"
    "Dynamics_365_Talent_Onboard"="DYNAMICS 365 FOR TALENT: ONBOARD"
    "DYNB365_CSI_VIRAL_TRIAL"="Dynamics 365 Customer Service Insights vTrial"
    "EducationAnalyticsP1"="Education Analytics"
    "EOP_ENTERPRISE"="Exchange Online Protection"
    "EOP_ENTERPRISE_PREMIUM"="Exchange Enterprise CAL Services (EOP DLP)"
    "EQUIVIO_ANALYTICS"="Office 365 Advanced eDiscovery"
    "EQUIVIO_ANALYTICS_GOV"="Office 365 Advanced eDiscovery for Government"
    "ERP_TRIAL_INSTANCE"="Dynamics 365 Operations Trial Environment"
    "EXCEL_PREMIUM"="Microsoft Excel Advanced Analytics"
    "EXCHANGE_ANALYTICS"="Microsoft MyAnalytics (Full)"
    "EXCHANGE_ANALYTICS_GOV"="Microsoft MyAnalytics for Government (Full)"
    "EXCHANGE_B_STANDARD"="EXCHANGE ONLINE POP"
    "EXCHANGE_FOUNDATION_GOV"="EXCHANGE FOUNDATION FOR GOVERNMENT"
    "EXCHANGE_L_STANDARD"="EXCHANGE ONLINE (P1)"
    "EXCHANGE_S_ARCHIVE"="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
    "EXCHANGE_S_ARCHIVE_ADDON"="Exchange Online Archiving"
    "EXCHANGE_S_DESKLESS"="Exchange Online Kiosk"
    "EXCHANGE_S_DESKLESS_GOV"="Exchange Online (Kiosk) for Government"
    "EXCHANGE_S_ENTERPRISE"="Exchange Online (Plan 2)"
    "EXCHANGE_S_ENTERPRISE_GOV"="Exchange Online (Plan 2) for Government"
    "EXCHANGE_S_ESSENTIALS"="EXCHANGE ESSENTIALS"
    "EXCHANGE_S_FOUNDATION"="Exchange Foundation"
    "EXCHANGE_S_FOUNDATION_GOV"="Exchange Foundation for Government"
    "EXCHANGE_S_STANDARD"="Exchange Online (Plan 1)"
    "EXCHANGE_S_STANDARD_GOV"="Exchange Online (Plan 1) for Government"
    "EXCHANGE_S_STANDARD_MIDMARKET"="EXCHANGE ONLINE PLAN"
    "EXCHANGEONLINE_MULTIGEO"="Exchange Online Multi-Geo"
    "EXPERTS_ON_DEMAND"="Microsoft Threat Experts - Experts on Demand"
    "FLOW_BUSINESS_PROCESS"="Flow per business process plan"
    "FLOW_CCI_BOTS"="Flow for CCI Bots"
    "FLOW_CUSTOMER_SERVICE_PRO"="Power Automate for Customer Service Pro"
    "FLOW_DEV_VIRAL"="Flow for Developer"
    "FLOW_DYN_APPS"="Power Automate for Dynamics 365"
    "FLOW_DYN_P2"="Flow for Dynamics 365"
    "FLOW_DYN_TEAM"="Power Automate for Dynamics 365"
    "FLOW_FOR_PROJECT"="Power Automate for Project"
    "FLOW_FORMS_PRO"="Power Automate for Dynamics 365 Customer Voice"
    "FLOW_O365_P1"="Power Automate for Office 365"
    "FLOW_O365_P1_GOV"="Power Automate for Office 365 for Government"
    "FLOW_O365_P2"="Power Automate for Office 365"
    "FLOW_O365_P2_GOV"="POWER AUTOMATE FOR OFFICE 365 FOR GOVERNMENT"
    "FLOW_O365_P3"="Power Automate for Office 365"
    "FLOW_O365_P3_GOV"="Power Automate for Office 365 for Government"
    "FLOW_O365_S1"="Power Automate for Office 365 F3"
    "FLOW_O365_S1_GOV"="Power Automate for Office 365 F3 for Government"
    "FLOW_P1_GOV"="Power Automate (Plan 1) for Government"
    "FLOW_P2"="Power Automate (Plan 2)"
    "FLOW_P2_VIRAL"="Flow Free"
    "FLOW_P2_VIRAL_REAL"="Flow P2 Viral"
    "Flow_Per_APP"="Power Automate for Power Apps per App Plan"
    "Flow_Per_APP_IWTRIAL"="Flow per app baseline access"
    "FLOW_PER_USER"="Flow per user plan"
    "FLOW_PER_USER_GCC"="Power Automate per User Plan for Government"
    "Flow_PowerApps_PerUser"="Power Automate for Power Apps per User Plan"
    "Flow_PowerApps_PerUser_GCC"="Power Automate for Power Apps per User Plan for GCC"
    "FLOW_VIRTUAL_AGENT_BASE"="Power Automate for Virtual Agent"
    "FORMS_GOV_E1"="Forms for Government (Plan E1)"
    "FORMS_GOV_E3"="FORMS FOR GOVERNMENT (PLAN E3)"
    "FORMS_GOV_E5"="Microsoft Forms for Government (Plan E5)"
    "FORMS_GOV_F1"="Forms for Government (Plan F1)"
    "FORMS_PLAN_E1"="Microsoft Forms (Plan E1)"
    "FORMS_PLAN_E3"="Microsoft Forms (Plan E3)"
    "FORMS_PLAN_E5"="Microsoft Forms (Plan E5)"
    "FORMS_PLAN_K"="Microsoft Forms (Plan F1)"
    "FORMS_PRO"="Dynamics 365 Customer Voice"
    "Forms_Pro_AddOn"="Microsoft Dynamics 365 Customer Voice Add-on"
    "Forms_Pro_CE"="Microsoft Dynamics 365 Customer Voice for Customer Engagement Plan"
    "Forms_Pro_Customer_Insights"="Microsoft Dynamics 365 Customer Voice for Customer Insights"
    "Forms_Pro_FS"="Microsoft Dynamics 365 Customer Voice for Field Service"
    "Forms_Pro_Marketing"="Microsoft Dynamics 365 Customer Voice for Marketing"
    "Forms_Pro_Service"="Microsoft Dynamics 365 Customer Voice for Customer Service Enterprise"
    "Forms_Pro_USL"="Microsoft Dynamics 365 Customer Voice USL"
    "GRAPH_CONNECTORS_SEARCH_INDEX"="Graph Connectors Search with Index"
    "GRAPH_CONNECTORS_SEARCH_INDEX_TOPICEXP"="Graph Connectors Search with Index (Viva Topics)"
    "GUIDES"="Dynamics 365 Guides"
    "INFO_GOVERNANCE"="Microsoft Information Governance"
    "INFORMATION_BARRIERS"="Information Barriers"
    "INSIDER_RISK"="Microsoft Insider Risk Management"
    "INSIDER_RISK_MANAGEMENT"="Microsoft Insider Risk Management"
    "Intelligent_Content_Services"="SharePoint Syntex"
    "Intelligent_Content_Services_SPO_type"="SharePoint Syntex - SPO type"
    "INTUNE_A"="Microsoft Intune"
    "Intune_Defender"="MDE_SecurityManagement"
    "INTUNE_EDU"="Microsoft Intune for Education"
    "INTUNE_O365"="Mobile Device Management for Office 365"
    "INTUNE_SMBIZ"="Microsoft Intune"
    "IT_ACADEMY_AD"="MS IMAGINE ACADEMY"
    "KAIZALA_O365_P1"="Microsoft Kaizala Pro Plan 1"
    "KAIZALA_O365_P2"="Microsoft Kaizala Pro Plan 2"
    "KAIZALA_O365_P3"="Microsoft Kaizala Pro Plan 3"
    "KAIZALA_STANDALONE"="Microsoft Kaizala"
    "LOCKBOX_ENTERPRISE"="Customer Lockbox"
    "LOCKBOX_ENTERPRISE_GOV"="Customer Lockbox for Government"
    "M365_ADVANCED_AUDITING"="Microsoft 365 Advanced Auditing"
    "M365_LIGHTHOUSE_CUSTOMER_PLAN1"="Microsoft 365 Lighthouse (Plan 1)"
    "M365_LIGHTHOUSE_PARTNER_PLAN1"="Microsoft 365 Lighthouse (Plan 2)"
    "MCO_TEAMS_IW"="MICROSOFT TEAMS"
    "MCO_VIRTUAL_APPT"="Microsoft Teams Premium Virtual Appointments"
    "MCOEV"="Microsoft 365 Phone System"
    "MCOEV_GOV"="Microsoft 365 Phone System for Government"
    "MCOEV_VIRTUALUSER"="Microsoft 365 Phone Standard Resource Account"
    "MCOEV_VIRTUALUSER_GOV"="Microsoft 365 Phone Standard Resource Account for Government"
    "MCOEVSMB"="SKYPE FOR BUSINESS CLOUD PBX FOR SMALL AND MEDIUM BUSINESS"
    "MCOFREE"="MCO FREE FOR MICROSOFT TEAMS (FREE)"
    "MCOIMP"="Skype for Business Online (Plan 1)"
    "MCOIMP_GOV"="Skype for Business Online (Plan 1) for Government"
    "MCOLITE"="SKYPE FOR BUSINESS ONLINE (PLAN P1)"
    "MCOMEETACPEA"="Microsoft 365 Audio Conferencing Pay-Per-Minute"
    "MCOMEETADV"="Microsoft 365 Audio Conferencing"
    "MCOMEETADV_GOV"="Microsoft 365 Audio Conferencing for Government"
    "MCOMEETBASIC"="Microsoft Teams Audio Conferencing with dial-out to select geographies"
    "MCOPSTN1"="Microsoft 365 Domestic Calling Plan"
    "MCOPSTN1_GOV"="Domestic Calling for Government"
    "MCOPSTN2"="DOMESTIC AND INTERNATIONAL CALLING PLAN"
    "MCOPSTN3"="MCOPSTN3"
    "MCOPSTN5"="MICROSOFT 365 DOMESTIC CALLING PLAN (120 min)"
    "MCOPSTN8"="Microsoft 365 Domestic Calling Plan (120 min) at User Level"
    "MCOPSTNC"="COMMUNICATIONS CREDITS"
    "MCOPSTNEAU"="AUSTRALIA CALLING PLAN"
    "MCOSTANDARD"="Skype for Business Online (Plan 2)"
    "MCOSTANDARD_GOV"="Skype for Business Online (Plan 2) for Government"
    "MCOSTANDARD_MIDMARKET"="SKYPE FOR BUSINESS ONLINE (PLAN 2) FOR MIDSIZ"
    "MCOVOICECONF"="SKYPE FOR BUSINESS ONLINE (PLAN 3)"
    "MDE_LITE"="Microsoft Defender for Endpoint Plan 1"
    "MDE_SMB"="Microsoft Defender for Business"
    "MDM_SALES_COLLABORATION"="MICROSOFT DYNAMICS MARKETING SALES COLLABORATION - ELIGIBILITY CRITERIA APPLY"
    "MFA_PREMIUM"="Microsoft Azure Multi-Factor Authentication"
    "MICROSOFT_BUSINESS_CENTER"="MICROSOFT BUSINESS CENTER"
    "MICROSOFT_COMMUNICATION_COMPLIANCE"="Microsoft 365 Communication Compliance"
    "MICROSOFT_ECDN"="Microsoft eCDN"
    "MICROSOFT_REMOTE_ASSIST"="Microsoft Remote Assist"
    "MICROSOFT_SEARCH"="Microsoft Search"
    "MICROSOFTBOOKINGS"="Microsoft Bookings"
    "MICROSOFTENDPOINTDLP"="Microsoft Endpoint DLP"
    "MICROSOFTSTREAM"="MICROSOFT STREAM"
    "MINECRAFT_EDUCATION_EDITION"="Minecraft Education Edition"
    "MIP_S_CLP1"="Information Protection for Office 365 - Standard"
    "MIP_S_CLP2"="Information Protection for Office 365 - Premium"
    "MIP_S_Exchange"="Data Classification in Microsoft 365"
    "ML_CLASSIFICATION"="Microsoft ML-based classification"
    "MMR_P1"="Meeting Room Managed Services"
    "MTP"="Microsoft 365 Defender"
    "MYANALYTICS_P2"="Insights by MyAnalytics"
    "MYANALYTICS_P2_GOV"="Insights by MyAnalytics for Government"
    "NBENTERPRISE"="Microsoft Social Engagement Enterprise"
    "NBPROFESSIONALFORCRM"="MICROSOFT SOCIAL ENGAGEMENT PROFESSIONAL - ELIGIBILITY CRITERIA APPLY"
    "NONPROFIT_PORTAL"="Nonprofit Portal"
    "Nucleus"="Nucleus"
    "O365_SB_Relationship_Management"="OUTLOOK CUSTOMER MANAGER"
    "OFFICE_BUSINESS"="Microsoft 365 Apps for Business"
    "OFFICE_FORMS_PLAN_2"="Microsoft Forms (Plan 2)"
    "OFFICE_FORMS_PLAN_3"="Microsoft Forms (Plan 3)"
    "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"="OFFICE 365 SMALL BUSINESS SUBSCRIPTION"
    "OFFICE_PROPLUS_DEVICE"="Microsoft 365 Apps for Enterprise (Device)"
    "OFFICE_SHARED_COMPUTER_ACTIVATION"="Office Shared Computer Activation"
    "OFFICEMOBILE_SUBSCRIPTION"="Office Mobile Apps for Office 365"
    "OFFICEMOBILE_SUBSCRIPTION_GOV"="Office Mobile Apps for Office 365 for GCC"
    "OFFICESUBSCRIPTION"="Microsoft 365 Apps for Enterprise"
    "OFFICESUBSCRIPTION_GOV"="Microsoft 365 Apps for enterprise G"
    "OFFICESUBSCRIPTION_unattended"="Microsoft 365 Apps for Enterprise (Unattended)"
    "ONEDRIVE_BASIC"="OneDrive for Business (Basic)"
    "ONEDRIVE_BASIC_GOV"="ONEDRIVE FOR BUSINESS BASIC FOR GOVERNMENT"
    "ONEDRIVEENTERPRISE"="ONEDRIVEENTERPRISE"
    "ONEDRIVESTANDARD"="OneDrive for Business (Plan 1)"
    "PAM_ENTERPRISE"="Office 365 Privileged Access Management"
    "PBI_PREMIUM_P1_ADDON"="Power BI Premium P"
    "POWER_APPS_DYN365_VIRAL_TRIAL"="Power Apps for Dynamics 365 vTrial"
    "POWER_AUTOMATE_ATTENDED_RPA"="Power Automate RPA Attended"
    "POWER_AUTOMATE_DYN365_VIRAL_TRIAL"="Power Automate for Dynamics 365 vTrial"
    "Power_Automate_For_Project_P1"="Power Automate for Project P1"
    "POWER_AUTOMATE_UNATTENDED_RPA"="Power Automate Unattended RPA add-on"
    "POWER_VIRTUAL_AGENTS_O365_F1"="Power Virtual Agents for Office 365 F1"
    "POWER_VIRTUAL_AGENTS_O365_P1"="Power Virtual Agents for Office 365 P1"
    "POWER_VIRTUAL_AGENTS_O365_P2"="Power Virtual Agents for Office 365 P2"
    "POWER_VIRTUAL_AGENTS_O365_P3"="Power Virtual Agents for Office 365 P3"
    "POWERAPPS_CUSTOMER_SERVICE_PRO"="Power Apps for Customer Service Pro"
    "POWERAPPS_DEV_VIRAL"="PowerApps for Developer"
    "POWERAPPS_DYN_APPS"="PowerApps for Dynamics 365"
    "POWERAPPS_DYN_P2"="Power Apps for Dynamics 365"
    "POWERAPPS_DYN_TEAM"="Power Apps for Dynamics 365"
    "POWERAPPS_GUIDES"="Power Apps for Guides"
    "POWERAPPS_O365_P1"="Power Apps for Office 365"
    "POWERAPPS_O365_P1_GOV"="Power Apps for Office 365 for Government"
    "POWERAPPS_O365_P2"="Power Apps for Office 365"
    "POWERAPPS_O365_P2_GOV"="POWER APPS FOR OFFICE 365 FOR GOVERNMENT"
    "POWERAPPS_O365_P3"="Power Apps for Office 365 (Plan 3)"
    "POWERAPPS_O365_P3_GOV"="Power Apps for Office 365 for Government"
    "POWERAPPS_O365_S1"="Power Apps for Office 365 F3"
    "POWERAPPS_O365_S1_GOV"="Power Apps for Office 365 F3 for Government"
    "POWERAPPS_P1_GOV"="PowerApps Plan 1 for Government"
    "POWERAPPS_P2"="Power Apps (Plan 2)"
    "POWERAPPS_P2_VIRAL"="PowerApps Trial"
    "POWERAPPS_PER_APP"="Power Apps per App Plan"
    "POWERAPPS_PER_APP_IWTRIAL"="PowerApps per app baseline access"
    "POWERAPPS_PER_APP_NEW"="Power Apps per app"
    "POWERAPPS_PER_USER"="Power Apps per User Plan"
    "POWERAPPS_PER_USER_GCC"="Power Apps per User Plan for Government"
    "POWERAPPS_PORTALS_LOGIN"="Power Apps Portals Login Capacity Add-On"
    "POWERAPPS_PORTALS_LOGIN_GCC"="Power Apps Portals Login Capacity Add-On for Government"
    "POWERAPPS_PORTALS_PAGEVIEW_GCC"="Power Apps Portals Page View Capacity Add-On for Government"
    "POWERAPPS_SALES_PRO"="Power Apps for Sales Pro"
    "POWERAPPSFREE"="MICROSOFT POWERAPPS"
    "POWERAUTOMATE_DESKTOP_FOR_WIN"="PAD for Windows"
    "POWERFLOWSFREE"="LOGIC FLOWS"
    "POWERVIDEOSFREE"="MICROSOFT POWER VIDEOS BASIC"
    "PREMIUM_ENCRYPTION"="Premium Encryption in Office 365"
    "PROJECT_CLIENT_SUBSCRIPTION"="Project Online Desktop Client"
    "PROJECT_CLIENT_SUBSCRIPTION_GOV"="Project Online Desktop Client for Government"
    "PROJECT_ESSENTIALS"="Project Online Essentials"
    "PROJECT_ESSENTIALS_GOV"="Project Online Essentials for Government"
    "PROJECT_FOR_PROJECT_OPERATIONS"="Project for Project Operations"
    "PROJECT_MADEIRA_PREVIEW_IW"="Dynamics 365 Business Central for IWs"
    "PROJECT_O365_F3"="Project for Office (Plan F)"
    "PROJECT_O365_P1"="Project for Office (Plan E1)"
    "PROJECT_O365_P2"="Project for Office (Plan E3)"
    "PROJECT_O365_P3"="Project for Office (Plan E5)"
    "PROJECT_P1"="Project P1"
    "PROJECT_PROFESSIONAL"="Project P3"
    "PROJECT_PROFESSIONAL_FACULTY"="Project P3 for Faculty"
    "PROJECTWORKMANAGEMENT"="Microsoft Planner"
    "PROJECTWORKMANAGEMENT_GOV"="Office 365 Planner for Government"
    "RECORDS_MANAGEMENT"="Microsoft Records Management"
    "RMS_S_ADHOC"="Rights Management Adhoc"
    "RMS_S_BASIC"="Microsoft Azure Rights Management Service"
    "RMS_S_ENTERPRISE"="Azure Rights Management"
    "RMS_S_ENTERPRISE_GOV"="Azure Rights Management"
    "RMS_S_PREMIUM"="Azure Information Protection Premium P1"
    "RMS_S_PREMIUM_GOV"="Azure Information Protection Premium P1 for GCC"
    "RMS_S_PREMIUM2"="Azure Information Protection Premium P2"
    "RMS_S_PREMIUM2_GOV"="Azure Information Protection Premium P2 for GCC"
    "SAFEDOCS"="Office 365 SafeDocs"
    "SCHOOL_DATA_SYNC_P1"="School Data Sync (Plan 1)"
    "SCHOOL_DATA_SYNC_P2"="School Data Sync (Plan 2)"
    "SharePoint Plan 1G"="SharePoint Plan 1G"
    "SHAREPOINT_PROJECT"="Project Online Service"
    "SHAREPOINT_PROJECT_EDU"="Project Online Service for Education"
    "SHAREPOINT_PROJECT_GOV"="Project Online Service for Government"
    "SHAREPOINT_S_DEVELOPER"="SHAREPOINT FOR DEVELOPER"
    "SHAREPOINTDESKLESS"="SharePoint Online Kiosk"
    "SHAREPOINTDESKLESS_GOV"="SharePoint KioskG"
    "SHAREPOINTENTERPRISE"="SharePoint Online (Plan 2)"
    "SHAREPOINTENTERPRISE_EDU"="SharePoint (Plan 2) for Education"
    "SHAREPOINTENTERPRISE_GOV"="SharePoint Plan 2G"
    "SHAREPOINTENTERPRISE_MIDMARKET"="SHAREPOINT PLAN 1"
    "SHAREPOINTLITE"="SHAREPOINTLITE"
    "SHAREPOINTONLINE_MULTIGEO"="SharePoint Multi-Geo"
    "SHAREPOINTSTANDARD"="SharePoint (Plan 1)"
    "SHAREPOINTSTANDARD_EDU"="SharePoint (Plan 1) for Education"
    "SHAREPOINTSTORAGE"="Office 365 Extra File Storage"
    "SHAREPOINTSTORAGE_GOV"="SHAREPOINTSTORAGE_GOV"
    "SHAREPOINTWAC"="Office for the web"
    "SHAREPOINTWAC_DEVELOPER"="OFFICE ONLINE FOR DEVELOPER"
    "SHAREPOINTWAC_EDU"="Office for the web (Education)"
    "SHAREPOINTWAC_GOV"="Office for the web (Government)"
    "SOCIAL_ENGAGEMENT_APP_USER"="Dynamics 365 AI for Market Insights - Free"
    "SPZA"="APP CONNECT"
    "SQL_IS_SSIM"="Microsoft Power BI Information Services Plan 1"
    "STREAM_O365_E1"="Microsoft Stream for Office 365 E1"
    "STREAM_O365_E1_GOV"="Microsoft Stream for O365 for Government (E1)"
    "STREAM_O365_E3"="Microsoft Stream for Office 365 E3"
    "STREAM_O365_E3_GOV"="MICROSOFT STREAM FOR O365 FOR GOVERNMENT (E3)"
    "STREAM_O365_E5"="Microsoft Stream for Office 365 E5"
    "STREAM_O365_E5_GOV"="Stream for Office 365 for Government (E5)"
    "STREAM_O365_K"="Microsoft Stream for O365 K SKU"
    "STREAM_O365_K_GOV"="Microsoft Stream for O365 for Government (F1)"
    "STREAM_O365_SMB"="Stream for Office 365"
    "STREAM_P2"="Microsoft Stream Plan 2"
    "STREAM_STORAGE"="Microsoft Stream Storage Add-On"
    "SWAY"="Sway"
    "TEAMS_ADVCOMMS"="Microsoft 365 Advanced Communications"
    "TEAMS_AR_DOD"="Microsoft Teams for DOD (AR)"
    "TEAMS_AR_GCCHIGH"="Microsoft Teams for GCCHigh (AR)"
    "TEAMS_FREE"="MICROSOFT TEAMS (FREE)"
    "TEAMS_FREE_SERVICE"="TEAMS FREE SERVICE"
    "TEAMS_GOV"="Microsoft Teams for Government"
    "Teams_Room_Standard"="Teams Room Standard"
    "TEAMS1"="Microsoft Teams"
    "TeamsEss"="Microsoft Teams Essentials"
    "TEAMSMULTIGEO"="Teams Multi-Geo"
    "TEAMSPRO_CUST"="Microsoft Teams Premium Personalized"
    "TEAMSPRO_MGMT"="Microsoft Teams Premium Intelligent"
    "TEAMSPRO_PROTECTION"="Microsoft Teams Premium Secure"
    "TEAMSPRO_VIRTUALAPPT"="Microsoft Teams Premium Virtual Appointment"
    "TEAMSPRO_WEBINAR"="Microsoft Teams Premium Webinar"
    "THREAT_INTELLIGENCE"="Microsoft Defender for Office 365 (Plan 2)"
    "THREAT_INTELLIGENCE_GOV"="Microsoft Defender for Office 365 (Plan 2) for Government"
    "UNIVERSAL_PRINT_01"="Universal Print"
    "UNIVERSAL_PRINT_NO_SEEDING"="Universal Print Without Seeding"
    "VIRTUAL_AGENT_BASE"="Virtual Agent Base"
    "Virtualization Rights for Windows 10 (E3/E5+VDA)"="Windows 10/11 Enterprise"
    "VISIO_CLIENT_SUBSCRIPTION"="Visio Desktop App"
    "VISIO_CLIENT_SUBSCRIPTION_GOV"="VISIO DESKTOP APP FOR Government"
    "VISIOONLINE"="Visio Web App"
    "VISIOONLINE_GOV"="VISIO WEB APP FOR GOVERNMENT"
    "VIVA_LEARNING_SEEDED"="Viva Learning Seeded"
    "VIVAENGAGE_CORE"="Viva Engage Core"
    "WHITEBOARD_FIRSTLINE1"="Whiteboard (Firstline)"
    "WHITEBOARD_PLAN1"="Whiteboard (Plan 1)"
    "WHITEBOARD_PLAN2"="Whiteboard (Plan 2)"
    "WHITEBOARD_PLAN3"="Whiteboard (Plan 3)"
    "WIN10_ENT_LOC_F1"="Windows 10 Enterprise E3 (Local Only)"
    "WIN10_PRO_ENT_SUB"="WINDOWS 10 ENTERPRISE"
    "WINBIZ"="Windows 10/11 Business"
    "WINDEFATP"="Microsoft Defender for Endpoint"
    "Windows Autopatch"="Windows Autopatch"
    "Windows Store for Business EDU Store_faculty"="Windows Store for Business EDU Store_faculty"
    "Windows_Autopatch"="Windows Autopatch"
    "WINDOWS_STORE"="Windows Store Service"
    "WINDOWSUPDATEFORBUSINESS_DEPLOYMENTSERVICE"="Windows Update for Business Deployment Service"
    "WORKPLACE_ANALYTICS"="Microsoft Workplace Analytics"
    "WORKPLACE_ANALYTICS_INSIGHTS_BACKEND"="Microsoft Workplace Analytics Insights Backend"
    "WORKPLACE_ANALYTICS_INSIGHTS_USER"="Microsoft Workplace Analytics Insights User"
    "YAMMER_EDU"="Yammer for Academic"
    "YAMMER_ENTERPRISE"="Yammer Enterprise"
    "YAMMER_MIDSIZE"="YAMMER MIDSIZE"
    "PRIVACY_MANGEMENT_RISK_EXCHANGE"="Priva Privacy Risk Management"  
    "PRIVACY_MANGEMENT_DSR_EXCHANGE"="Priva Privacy DSR"  
    "PRIVACY_MANGEMENT_RISK"="Priva Privacy Risk Management"  
    "PRIVACY_MANGEMENT_DSR"="Priva Privacy DSR"
    "MIP_S_EXCHANGE_CO"="Microsoft Information Protection"
    "DYN365_CDS_VIRAL"="Microsoft Dynamics Common Data Service (Free)"
}    
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

#construct our final output
if ($reporttype -match 'Simple'){
    $serviceusage2 = $serviceusage2 | Where-Object {($_.Servicename -like "RMS_S_*" -or $_.ServiceName -like "COMPLIANCE_MANAGER*" -or $_.ServiceName -like "LOCKBOX_*" -or $_.ServiceName -like "MIP_S_*" -or $_.Servicename -like "INFORMATION_Barriers" -or $_.ServiceName -like "CONTENT*" -or $_.ServiceName -like "M365_ADVACNED*" -or $_.ServiceName -like "MICROSOFT_COMMUNICATION*" -or $_.ServiceName -like "COMMUNICATIONS_*" -or $_.ServiceName -like "CUSTOMER_KE*" -or $_.ServiceName -like "INFO_GOV*" -or $_.ServiceName -like "INSIDER_RISK_MANAG*" -or $_.ServiceName -like "ML_CLASSIFI*" -or $_.ServiceName -like "RECORDS_*" -or $_.ServiceName -like "EQUIVIO*" -or $_.ServiceName -like "PAM*" -or $_.ServiceName -like "PRIVACY*" -or $_.ServiceName -like "M365_ADVANCED_AUDIT*")} 
    $outputlist = $serviceusage2 | Select-Object Servicename, @{ n = 'FriendlyName'; e= {$_ | ForEach-Object { $FriendlyLicenses[$_.ServiceName] } } },ActivatedUsers  | Sort-Object FriendlyName
    Write-host "Generating Simple HTML Report"
}

elseif ($reporttype -match 'Detailed') {
    $outputlist = $serviceusage2 | Select-Object Servicename, @{ n = 'FriendlyName'; e= {$_ | ForEach-Object { $FriendlyLicenses[$_.ServiceName] } } },ActivatedUsers  | Sort-Object FriendlyName
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
