# Project

Scripts and Files to support the Compliance Partner Build Intent Engagements.

## EngagementPOEReport

Use the Engagement POE Report as part of the Data Security Engagement. Please see the delivery guide on how to use the output as part of the Proof of execution. The most recent version is 3.25 (published January 2025). This most recent version aligns to Version 7 of the Data Security Engagement and includes reporting for the Data Security for AI module a

**V3.5 Updates**
Added report output for Data Security for AI module

**V3.4 Updates**
Updated formatting of date time outputs to force Invariant (universal) date time output to help expedite POE reviews for organizations providing POE's that do not default to English abbreviations for Months in MMM-dd-YYYY format.

**V3.3 Updates**
Updated the DLP and Content Search logic to only include items created in the last 90 to limit the amount of data returned in large environments while still capturing what was created during the engagement
-Added a ZIP file containing current version of signed script.

**V3.2 updates**
Changed sort order on Content Search to put most recent searches on top.
-Signed Script (the file should have 547 lines including trailing blank line after digital signature)

**V3.0 updates**
-Transition from Microsoft Graph Powershell to Exchange Online Powershell
-Signed Script (the file should have 549 lines including trailing blank line after digital signature)
-- A zip file version of the script is also available in the repository (EngagementPOEReport.zip)

### Current issues or limitations

1) Current version has only been tested against Commercial Office 365 tenants. If you need to connect to a GCC or Regional(China / Germany) Tenant, please update the powershell connection strings inside the code

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit <https://cla.opensource.microsoft.com>.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
