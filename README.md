---
Version: 1.2 
Title: M365 Apps Intune scripted dynamic install using Office Deployment Toolkit
Authors: JankeSkanke, sandytsang
Owner: JankeSkanke
Date: 23/11/2022
Description: Deploying M365 with Intune dynamically as a Win32App
---
  [![HitCount](https://hits.dwyl.com/msendpointmgr/m365apps.svg?style=flat)](http://hits.dwyl.com/msendpointmgr/m365apps)

# M365 Apps Intune scripted dynamic install using Office Deployment Toolkit 

## This solution covers installation of the following products

* [M365 Apps(Office)](#main-office-package-getting-configuration-from-external-url)
* [Project](#project-and-visio)
* [Visio](#project-and-visio)
* [Proofing tools](#proofing-tools)

Each product is made of the following components

* Install script (PowerShell)
* Local Configuration.xml or External URL
* Detection (script or documented)

***
## Main Office Package (getting configuration from External URL)

1. Define your config XML (Example below, can be generated at office.com)

```xml
<Configuration ID="9aa11e20-2e29-451a-b0ba-f1ae3e89d18d">
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" MigrateArch="TRUE">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Updates Enabled="TRUE" />
  <AppSettings>
    <Setup Name="Company" Value="Company Name" />
  </AppSettings>
  <Display Level="None" AcceptEULA="FALSE" />
</Configuration>
```

1. Create a .Intunewim using the Win32 Content Prep tool [Prepare Win32 app content for upload](https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-prepare?WT.mc_id=EM-MVP-5002085) containing the InstallM365Apps.ps1 
2. Upload .Intunewim and define the following parameters during install 
    * Install Command:
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"```
    * Uninstall Command:
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1``` (Not working yet)

    <img src="/.images/officeinstall.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">

    * Install behaviour: System 
    * Requirements (probable 64 bit Windows something)
    * Detection: Use PowerShell detection Script M365AppsWin32DetectionScript.ps1 
  1. Assign

***

## Main Office Package (using configuration.xml inside package)

1. Define your config XML (Example below, can be generated at office.com)
```xml
<Configuration ID="9aa11e20-2e29-451a-b0ba-f1ae3e89d18d">
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" MigrateArch="TRUE">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Updates Enabled="TRUE" />
  <AppSettings>
    <Setup Name="Company" Value="Company Name" />
  </AppSettings>
  <Display Level="None" AcceptEULA="FALSE" />
</Configuration>
```
2. Create a .Intunewim using the Win32 Content Prep tool [Prepare Win32 app content for upload](https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-prepare?WT.mc_id=EM-MVP-5002085) containing the configuration.xml and the InstallM365Apps.ps1 
2. Upload .Intunewim and define the following parameters during install 
    * Install Command:
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1```
    * Uninstall Command:
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1``` (Not working yet)
    * Install behaviour: System
    * Requirements (probable 64 bit Windows something)
    * Detection: Use PowerShell detection Script M365AppsWin32DetectionScript.ps1
  1. Assign

***
## Project and Visio

1. Define your config XML 
```xml
<Configuration ID="fc6a02c8-622f-4cf4-bf7f-6c57847b0580">
  <Add OfficeClientEdition="64" Version="MatchInstalled">
    <Product ID="ProjectProRetail">
      <Language ID="MatchInstalled" Fallback="en-us" TargetProduct="O365ProPlusRetail"/>
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="OneDrive" />
    </Product>
  </Add>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="TRUE" />
  <AppSettings>
    <Setup Name="Company" Value="Company Name" />
  </AppSettings>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
```
> The provided configuration.xml can not be imported on config.office.com and must be stored in for instance a Azure blob storage when using external url option. 

* This configuration file example will match the language installed (M365 Apps). 
* TargetProduct is required for that to work. 
* This example will shutdown running Office Apps for end users during install. 
* Follow others steps from the main office package. 
 <img src="/.images/visioinstall.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">

Project Install Command (Local):

```PowerShell
powershell.exe -executionpolicy bypass -file InstallProject.ps1 
```

Project Install Command (ExternalXML):

```PowerShell
powershell.exe -executionpolicy bypass -file InstallProject.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"
```

Visio Install Command (Local):

```PowerShell
powershell.exe -executionpolicy bypass -file InstallVisio.ps1 
```

Visio Install Command (ExternalXML):

```PowerShell
powershell.exe -executionpolicy bypass -file InstallVisio.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"
```

***

## Proofing tools or LanguagePacks

We recommend installing only 1 language on the computers unless your requirements are very specific. But there might still be some users that would need a complete language pack or proofing tools for various languages. The main thinking here is to have all possible proofing tools in your environment as available to end user to install by their own choosing.

For these 2 options the included configuration.xml files are just "templates" as the script it self will rewrite the XML dynamically based on the parameters you send to the script.

>There is no need to maintain this XML as long as Microsoft does not change the XML requirements.

***
**EXAMPLES for LanguagePacks**

```PowerShell
powershell.exe -executionpolicy bypass -file InstallLanguagePacks.ps1 -LanguageID nb-no -Action Install
powershell.exe -executionpolicy bypass -file InstallLanguagePacks.ps1 -LanguageID nb-no -Action Uninstall
```


**EXAMPLES for Proofing Tools**

```PowerShell
powershell.exe -executionpolicy bypass -file InstallProofingTools.ps1 -LanguageID nb-no -Action Install
powershell.exe -executionpolicy bypass -file InstallProofingTools.ps1 -LanguageID nb-no -Action Uninstall
```

<img src="/.images/proofinginstall_1.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">
<img src="/.images/proofinginstall_2.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">

***

It is also recommended that you have a requirement to check if Main Office is installed on the device as the install will fail if you try to install the a language pack or proofing tools without Office installed. This can be done using a registry key check or using the provided requirement script.

<img src="/.images/proofingrequire.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">

**ProofingRequirementScript.ps1**

***

Detection of language packs is best using registry detection. Example Norwegian: 

* Key Path: **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - nb-no**
* Value Name: **DisplayName**
* String Comparison - Equals
* Microsoft 365 Apps for enterprise - nb-no

Detection of the proofing tools can be done either with the provided detection script, customized for each LanguageID or by using a registry key check(Recommended).

EXAMPLE Detection Rule:
Registry
<img src="/.images/proofingdetect.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">

**HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - nb-no.proof**

***
Get URL from config.office.com : 

<img src="/.images/configofficeurl.png" alt="Office Install XML" title="Office Install XML" style="display: inline-block; margin: 0 auto; max-width: 300px">


For more details and instructions go to [MSEndpointMgr Blog](https://msendpointmgr.com/2022/10/23/installing-m365-apps-as-win32-app-in-intune/)

This solution has been developed by @JankeSkanke with assistance from @sandytsang and @maurice-daly

