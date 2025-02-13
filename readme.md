# PowerShell Scripts to Create Shared Templates

Author: Andy Prosser
Email: [mailto:andy@theexectechs.com]
Web: [https://theexectechs.com] 

## Overview
Document templates are a key element to delivering a consistent brand across an organization. Microsoft 365 allows the creation of a Brand Hub, yet the automatic synchronization of document templates are limited to Microsoft 365 E3 & E5 licenses. This package has been created to assist smaller organizations achieve a similar centralized brand hub as available to enterprise customers.

### Creating a Brand Hub (Enterprise Customers)
Creating a Brand Hub for customers with Microsoft E3 or E5 licenses require little configuration. 
Simply follow Microsoft's documentation:
[https://learn.microsoft.com/en-us/sharepoint/organization-assets-library]

The included <Create-TemplateLibrary.ps1> script can be used to automate the creation of the Brand Hub site. This creates the required Document Libraries and sets the required permissions to allow employees access to the hub. 

### Creating a Brand Hub (Business Customers)
It is possible to replicate the features of a Brand Hub for smaller organizations, yet it requires a little more work.
Use the included <Create-TemplateLibrary.ps1> script to automate the creation of the Brand Hub site. This creates the required Document Libraries and sets the required permissions to allow employees access to the hub. 

The deployment & configuration of the templates themselves require the use of the <Sync-Templates.ps1> script. Each User **MUST** run this script at least *ONCE* to setup the OneDrive Library Sync and configure Microsoft Office 365. This script will ONLY work on Windows 10 or Window 11.

## <Create-TemplateLibrary.ps1> Script
The <Create-TemplateLibrary.ps1> script automatically connects to SharePoint Online to create the required SharePoint Site and the required document libraries. It has been developed using PowerShell and leverages the PnP PowerShell Library.
[https://pnp.github.io/powershell/index.html]

It is recommended that the <Create-TemplateLibrary.ps1> script is run by a SharePoint Administrator.
You will need to edit the "Configuration" section of the script before running it. Please ensure it matches your SharePoint Online tenant required settings.

Most of these settings can remain as-is, except the '$sharePointUrl' parameter. This **MUST** match your tenant.

```
<# CONFIGURATION #>
$isTenantAdmin = $true
$sharePointUrl = "https://exectechs.sharepoint.com/"
$templateSite = [PSCustomObject]@{
    Name    = "Company Templates"
    Alias   = "templates"
    Locale  = 1033  # en-US
    Tz      = [PnP.Framework.Enums.TimeZone]::UTCPLUS1000_CANBERRA_MELBOURNE_SYDNEY
    Assets  = "Assets"
    Library = "Templates"
}
```

Once the script completes, it will output a list of "SharePoint Online Connection Parameters".
Make note of these values as the will be needed for the <Sync-Templates.ps1> script.

## <Sync-Templates.ps1> Script
This script does not require any special permissions or packages. The only requirement is that OneDrive is installed on the target Windows computer.
Each user will need to execute the <Sync-Templates.ps1> script at least once. If it is run mul.tiple times, it will check that the require library is being synchronised and skip if it is setup.

Edit the "Configuration" section of the script and deploy to your users to run. The values you need are the "SharePoint Online Connection Parameters" from the <Create-TemplateLibrary.ps1> script. We recommend you use a login script and/or management tool to deploy this script.

```
<# CONFIGURATION #>
$siteId    = 'd0a2f731-3f1e-40d2-b429-d26eff835ca9'
$webId     = 'b3f1c3fd-24bd-40e2-a8b5-86fb07efb3d9'
$listId    = '34640710-9bcf-435f-86cf-4701cf5583a6'
$listTitle = 'Templates'
$webUrl    = 'https://exectechs.sharepoint.com/teams/templates'
$webTitle  = 'Company Templates'
$orgName   = 'TheExecTechs'
```

## Deploying <Sync-Templates.ps1> Script with Intune
Deploying scripts with Microsoft Endpoint Manager (aka Intune) is very easy, but does require your IT team to setup the required configuration.
When creating the script deployment configuration, ensure the following values are set;
* **Run this script using the logged on credentials** to **Yes**.  
* **Enforce script signature check** to **No** (if you sign the script you can set yours to **Yes**).  
* **Run script in 64-bit PowerShell Host** to **Yes**

Make sure the script is assigned to the require user group (not devices).

All parameters for this script are self-contained.

## Manually Deploying <Sync-Templates.ps1> Script
In order toi manually deploy the <Sync-Templates.ps1> script, you will need to copy this script somewhere that is easy for users to access. Most email systems will block PowerShell scripts for security reasons. The file will need to be downloaded to the local computer.
You can run a PowerShell script by *right-clicking* it and then clicking **Run with PowerShell**.

## <Import-FilesToSharePoint.ps1> Script
This bonus script allows the easy importing files from a File Share or other local file source. The script supports *synchronization* which allows changes to files to be made and allowing large file shares to be migrated without significantly impacting use of the files. Once the bulk of the files have been transferred, a final "sync" can be performed to ensure any changed files have been updated.

This file is provided *as-is, without warranty*


## Editing PowerShell Scripts
The are many tools you can use to edit a PowerShell Script. The recommended tool is Visual Studio Code.
[https://code.visualstudio.com/Download]


