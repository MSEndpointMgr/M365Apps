---
Version: 1.1 
Title: M365 Apps Intune scripted dynamic install using Office Deployment Toolkit
Authors: JankeSkanke, sandytsang
Owner: JankeSkanke
Date: 28/10/2022
Description: Deploying M365 with Intune dynamically as a Win32App
---
  [![HitCount](https://hits.dwyl.com/msendpointmgr/m365apps.svg?style=flat)](http://hits.dwyl.com/msendpointmgr/m365apps)

# M365 Apps Intune scripted dynamic install using Office Deployment Toolkit 
## This solution covers installation of the following products 
* [M365 Apps(Office)](#Main-Office-Package)
* [Project](#Project-and-Visio)
* [Visio](#Project-and-Visio)
* [Proofing tools](#Proofing-tools)

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
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1 ``` (Not working yet)
    * Install behaviour: System 
    * Requirements (probable 64 bit Windows something)
    * Detection: Use PowerShell detection Script M365AppsWin32DetectionScript.ps1 
  3. Assign 

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
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"```
    * Uninstall Command: 
      * ```powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1 ``` (Not working yet)
    * Install behaviour: System 
    * Requirements (probable 64 bit Windows something)
    * Detection: Use PowerShell detection Script M365AppsWin32DetectionScript.ps1 
  3. Assign 

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

Project Install Command (Local):
```  
powershell.exe -executionpolicy bypass -file InstallProject.ps1 
```
Project Install Command (ExternalXML):
```
powershell.exe -executionpolicy bypass -file InstallProject.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"
```

Visio Install Command (Local):

```  
powershell.exe -executionpolicy bypass -file InstallVisio.ps1 
```
Visio Install Command (ExternalXML):
```
powershell.exe -executionpolicy bypass -file InstallVisio.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"
```
***
## Proofing tools

We recommend installing only 1 language on the computers unless your requirements are very specific. But there might still be need for proofing tools for multiple languages. The main thinking here is to have all possible proofing tools in your environment as available to end user to install by their own choosing. 

For proofing tool the included configuration.xml files are just "templates" as the script it self will rewrite the XML dynamically based on the parameters you send to the script. 

>There is no need to maintain this XML as long as Microsoft does not change the XML requirements. 

***
**EXAMPLES for Proofing Tools**
```
powershell.exe -executionpolicy bypass -file Install-Proofing-Tools.ps1 -LanguageID nb-no -Action Install
powershell.exe -executionpolicy bypass -file Install-Proofing-Tools.ps1 -LanguageID nb-no -Action Uninstall
```
***
It is also recommended that you have a requirement to check if Main Office is installed on the device as the install will fail if you try to install the proofing tools without Office installed. 
This can be done using a registry key check or using the provided requirement script 
**ProofingRequirementScript.ps1**
***
Detection of the proofing tools can be done either with the provided detection script, customized for each LanguageID or by having a registry key check. 

EXAMPLE Detection Rule: 
Registry

**HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - nb-no.proof**


***
For more details and instructions go to [MSEndpointMgr Blog](https://msendpointmgr.com/2022/10/23/installing-m365-apps-as-win32-app-in-intune/)

This solution has been developed by @JankeSkanke with assistance from @sandytsang and @maurice-daly

