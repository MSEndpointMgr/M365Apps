# Changelog

## 1.2 - 23.11.2023

* Changed to Evergreen Url for setup.exe to avoid any issues with "scraping" the ODT download site and simplification. (Requested by [@aaronparker](https://github.com/aaronparker))
* Added cleanup action to remove temporary setup files after installation
* Added Languagepacks as an separated option. Install with parameters same way as Proofing tools. The XML files in the folder is only templates used by script to configure given language.

Example

```PowerShell
powershell.exe -executionpolicy bypass -file InstallLanguagePacks.ps1 -LanguageID nb-no -Action Install
powershell.exe -executionpolicy bypass -file InstallLanguagePacks.ps1 -LanguageID nb-no -Action Uninstall
```

***

## 1.1 - 28.10.2023

* Added support for external XML configuration.xml for M365 Apps main package, Visio and Project as an option
* No change for Proofing Tools

Example

Without external XML (Requires configuration.xml in the package) - Same as Version 1.0

```PowerShell
powershell.exe -executionpolicy bypass -file <installfile>.ps1
```

With external XML (Requires XML to be provided by URL)  - New option in the scripts

```PowerShell
powershell.exe -executionpolicy bypass -file <installfile>.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"
```

***

## 1.0 Initial release - 23.10.2022

No additional information - see Readme

***
