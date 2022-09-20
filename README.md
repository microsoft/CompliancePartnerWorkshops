# Project

## ComplianceActivationAssessment
Use the ComplianceActivationAssesment.ps1 file as part of the Protect and Govern Sensitive Data Activator
Follow the instructions in the workshop guide run the script and include the output of the report as part of your final results for your customer

### Current Issues and Limitations
1) The ComplianceActivationAssessment Report has only been tested against Commercial Office 365 Tenants.  If you need to connect to a GCC or Regional(China / Germany) Tenant, please update the powershell connection strings inside the code
2) Scripts have only been tested against English/Unicode lanuguages
3) License Friendly Names MAY not exist for non commercial license SKUs

## WorkshopPOEReport
Use the workshoppoereport.ps1 file as part of the Protect and Govern Sensitive Data Activator
Follow the instructions in the workshop guide run the script and include the output of the report as part of your final results for your customer

### Current Issues and Limitations:
1) The WorkshopPOE Report only works against Commercial Office 365 Tenants.  If you need to connect to a GCC or Regional(China / Germany) Tenant, please update the powershell connection strings inside the code
2) The WorkshopPOE Report currently uses the AzureAD powershell Module.  It will be updated to GraphAPI in a future version
3) Scripts have only been tested against English/Unicode lanuguages

## ComplianceEnvrionmentPrep
Use the complianceenvriomentprep.ps1 file as part of the Mitigate Complinace and Prviacy Risks Activator
Follow the instructions in the workshop guide and run the script to prepare the isolated Microsoft 365 Developer Tenant.

### Current Issues and Limitations:
1) The ComplianceEnvriomentPrep script is designed to be used against tenants that are provisioned as part of the Microsoft 365 Developer Subscription. It has not been tested against other Microsoft 365 envrioments
2) use the startup switch '-debug' to enable basic logging and get an output of information logged to the screen
3) Scripts have only been tested against English/Unicode lanuguages

### Other Files
The additional files in this repository are developed for the Mitigate Compliance and Privacy Risks Activator. Please refer to the engagement master delivery guide on how to leverage them
1) Rulepack.xml - Custom sensitive information type rule pack
2) DeleteFileFlow.zip - Power Automate Flow
3) FileCopyFlow.zip - Power Automate Flow
4) FileCreationFlow.zip - Power Automate Flow
5) Mark8.zip - Sample files

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
