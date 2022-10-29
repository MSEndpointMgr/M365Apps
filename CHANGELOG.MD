# Changelog 
***
## 1.0 Initial release - 23.10.2022
No additional information - see Readme

***
## 1.1 - 28.10.2023
* Added support for external XML configuration.xml for M365 Apps main package, Visio and Project as an option 
* No change for Proofing Tools
 
Example

Without external XML (Requires configuration.xml in the package) - Same as Version 1.0 
```
powershell.exe -executionpolicy bypass -file <installfile>.ps1
```

With external XML (Requires XML to be provided by URL)  - New option in the scripts 
```
powershell.exe -executionpolicy bypass -file <installfile>.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"
```
***