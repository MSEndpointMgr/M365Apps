<#
.SYNOPSIS
  Script to install M365 Apps as a Win32 App

.DESCRIPTION
    Script to install Office as a Win32 App during Autopilot by downloading the latest Office setup exe from evergreen url
    Running Setup.exe from downloaded files with provided config.xml file. 

.EXAMPLE
    Without external XML (Requires configuration.xml in the package)
    powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1
    With external XML (Requires XML to be provided by URL)  
    powershell.exe -executionpolicy bypass -file InstallM365Apps.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"

    If you want to cleanup OEM Based Office during ESP phase: InstallM365Apps.ps1 -XMLURL "https://mydomain.com/xmlfile.xml" -CleanOEM

.NOTES
    Version:        1.3
    Author:         Jan Ketil Skanke
    Contact:        @JankeSkanke
    Creation Date:  01.07.2021
    Updated:        (2023-06-06)
    Version history:
        1.0.0 - (2022-23-10) Script released 
        1.1.0 - (2022-25-10) Added support for External URL as parameter 
        1.2.0 - (2022-23-11) Moved from ODT download to Evergreen url for setup.exe 
        1.2.1 - (2022-01-12) Adding function to validate signing on downloaded setup.exe
        1.3.0 - (2023-06-06) Adding functionality to cleanup OEM version of Office. 
#>
#region parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [string]$XMLUrl,
    [Parameter(Mandatory = $false)]
    [switch]$CleanOEM
)
#endregion parameters
#Region Functions
function Write-LogEntry {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,
        [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = $LogFileName
    )
    # Determine log file location
    $LogFilePath = Join-Path -Path $env:SystemRoot -ChildPath $("Temp\$FileName")
	
    # Construct time stamp for log entry
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
	
    # Construct date for log entry
    $Date = (Get-Date -Format "MM-dd-yyyy")
	
    # Construct context for log entry
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
	
    # Construct final log entry
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
	
    # Add value to log file
    try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
        if ($Severity -eq 1) {
            Write-Verbose -Message $Value
        }
        elseif ($Severity -eq 3) {
            Write-Warning -Message $Value
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}
function Start-DownloadFile {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    Begin {
        # Construct WebClient object
        $WebClient = New-Object -TypeName System.Net.WebClient
    }
    Process {
        # Create path if it doesn't exist
        if (-not(Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }

        # Start download of file
        $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
    }
    End {
        # Dispose of the WebClient object
        $WebClient.Dispose()
    }
}
function Invoke-FileCertVerification {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath
    )
    # Get a X590Certificate2 certificate object for a file
    $Cert = (Get-AuthenticodeSignature -FilePath $FilePath).SignerCertificate
    $CertStatus = (Get-AuthenticodeSignature -FilePath $FilePath).Status
    if ($Cert){
        #Verify signed by Microsoft and Validity
        if ($cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid"){
            #Verify Chain and check if Root is Microsoft
            $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
            $chain.Build($cert) | Out-Null
            $RootCert = $chain.ChainElements | ForEach-Object {$_.Certificate}| Where-Object {$PSItem.Subject -match "CN=Microsoft Root"}
            if (-not [string ]::IsNullOrEmpty($RootCert)){
                #Verify root certificate exists in local Root Store
                $TrustedRoot = Get-ChildItem -Path "Cert:\LocalMachine\Root" -Recurse | Where-Object { $PSItem.Thumbprint -eq $RootCert.Thumbprint}
                if (-not [string]::IsNullOrEmpty($TrustedRoot)){
                    Write-LogEntry -Value "Verified setupfile signed by : $($Cert.Issuer)" -Severity 1
                    Return $True
                }
                else {
                    Write-LogEntry -Value  "No trust found to root cert - aborting" -Severity 2
                    Return $False
                }
            }
            else {
                Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2 
                Return $False
            }
        }
        else {
            Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2 
            Return $False
        }  
    }
    else {
        Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
        Return $False
    }
}
function Test-OfficeExists{
    Write-LogEntry -Value "Check if M365 Apps exists on device" -Severity 1
    $RegistryKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $M365Apps = "Microsoft 365 Apps"
    $M365AppsCheck = $RegistryKeys | Get-ItemProperty | Where-Object { $_.DisplayName -match $M365Apps }
    if ($M365AppsCheck) {
        Write-LogEntry -Value "Microsoft 365 Apps detected with version $($M365AppsCheck[0].DisplayVersion)" -Severity 1
        return $true
       }else{
        Write-LogEntry -Value "Microsoft 365 Apps not detected" -Severity 1
        return $false
    }
}
function Test-HasDeviceESPCompleted{
    $KeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\Autopilot\EnrollmentStatusTracking\Device\Setup\"
    $Property = "HasProvisioningCompleted"
    try {
        $regKey = Get-ItemProperty -Path $KeyPath -ErrorAction Stop
        $propertyExists = $regKey.PSObject.Properties.Name -contains $Property

        if ($propertyExists) {
           return $true
        } else {
            return $false
        }
    } catch {
        return $false
    }
}
function Test-IsAutopilotDevice{
    $autopilotDiagPath = "HKLM:\software\microsoft\provisioning\Diagnostics\Autopilot"
    $values = Get-ItemProperty "$autopilotDiagPath"
    if (-not $values.CloudAssignedTenantId) {
        return $false
    }
    else{
        return $true
    }
}
#Endregion Functions

#Region Initialisations
$LogFileName = "M365AppsSetup.log"
$DetectionRegKeyName = "MSEndpointMgr" #Change with company name if needed 

#Endregion Initialisations

#Initate Install
Write-LogEntry -Value "Initiating Office setup process" -Severity 1
#Attempt Cleanup of SetupFolder
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup") {
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
}

$SetupFolder = (New-Item -ItemType "directory" -Path "$($env:SystemRoot)\Temp" -Name OfficeSetup -Force).FullName

try {
    #Download latest Office setup.exe
    $SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"
    
    try {
        #Start install preparations
        $SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
        if (-Not (Test-Path $SetupFilePath)) {
            Throw "Error: Setup file not found"
        }
        Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1
        try {
        # Initate OEM Cleanup 
            if ($CleanOEM){
                Write-LogEntry -Value "OEM Cleanup requested" -Severity 1
                if (Test-OfficeExists){
                    Write-LogEntry -Value "Office found - attempting cleanup" -Severity 1
                    if (Test-IsAutopilotDevice){
                        # Device is an Autopilot device
                        Write-LogEntry -Value "Device is an Autopilot Device" -Severity 1
                        if (-not (Test-HasDeviceESPCompleted)){
                            # Device is in Device ESP Phase 
                            Write-LogEntry -Value "Provisioning is not completed - starting OEM Cleanup" -Severity 1
                            
                            # Generate Removal XML 
                            $XmlFilePath = "$SetupFolder\remove.xml"
                            $xmlDocument = New-Object System.Xml.XmlDocument
                            $rootElement = $xmlDocument.CreateElement("Configuration")
                            $xmlDocument.AppendChild($rootElement)
                            $removeElement = $xmlDocument.CreateElement("Remove")
                            $removeElement.SetAttribute("All", "TRUE")
                            $rootElement.AppendChild($removeElement)
                            $removeElement.
                            $propertyElement = $xmlDocument.CreateElement("Property")
                            $propertyElement.SetAttribute("Name", "FORCEAPPSHUTDOWN")
                            $propertyElement.SetAttribute("Value", "TRUE")
                            $rootElement.AppendChild($propertyElement)
                            $displayElement = $xmlDocument.CreateElement("Display")
                            $displayElement.SetAttribute("Level", "None")
                            $displayElement.SetAttribute("AcceptEULA", "TRUE")
                            $rootElement.AppendChild($displayElement)
                            $xmlDocument.Save($XmlFilePath)
                            Write-LogEntry -Value "Starting OEM Office Removal using $XmlFilePath" -Severity 1
        
                            if (Invoke-FileCertVerification -FilePath $SetupFilePath){
                                 #Starting Office removal with configuration file               
                                Try {
                                    #Running office installer
                                    Write-LogEntry -Value "Starting M365 Apps OEM Cleanup" -Severity 1
                                    $OfficeRemoval = Start-Process $SetupFilePath -ArgumentList "/configure $($SetupFolder)\remove.xml" -Wait -PassThru -ErrorAction Stop
                                    Write-LogEntry -Value "M365 Apps OEM Cleanup completed" -Severity 1
                                }
                                catch [System.Exception] {
                                    Write-LogEntry -Value  "Error running M365 Apps OEM Cleanup. Errormessage: $($_.Exception.Message)" -Severity 3
                                }
                            }
                        }
                        else {
                            Write-LogEntry -Value "OEM Cleanup should only be done during ESP Device Phase" -Severity 1
                        } 
                    }
                    else {
                        Write-LogEntry -Value "Device is not an Autopilot Device - OEM Cleanup will not be attempted" -Severity 1
                    }                       
                }
                else{
                    Write-LogEntry -Value "Office not found - OEM Cleanup not needed" -Severity 1
                } 
                #Adding custom detection keys
                New-Item -Path "HKLM:SOFTWARE\$($DetectionRegKeyName)" -Name "M365AppsInstall" -Force | Out-Null
                Set-ItemProperty -Path "HKLM:SOFTWARE\$($DetectionRegKeyName)\M365AppsInstall" -Name "OEMClean" -Value "Yes"
            }
    
        # After OEM Cleanup install Office 
            try {
                #Prepare Office Installation
                $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$($SetupFolder)\setup.exe").FileVersion 
                Write-LogEntry -Value "Office C2R Setup is running version $OfficeCR2Version" -Severity 1
                if (Invoke-FileCertVerification -FilePath $SetupFilePath){
                #Check if XML URL is provided, if true, use that instead of trying local XML in package
                    if ($XMLUrl) {
                        Write-LogEntry -Value "Attempting to download configuration.xml from external URL" -Severity 1
                        try {
                            #Attempt to download file from External Source
                            Start-DownloadFile -URL $XMLURL -Path $SetupFolder -Name "configuration.xml"
                            Write-LogEntry -Value "Downloading configuration.xml from external URL completed" -Severity 1
                        }
                        catch [System.Exception] {
                            Write-LogEntry -Value "Downloading configuration.xml from external URL failed. Errormessage: $($_.Exception.Message)" -Severity 3
                            Write-LogEntry -Value "M365 Apps setup failed" -Severity 3
                            exit 1
                        }
                    }
                    else {
                        #Local configuration file only 
                        Write-LogEntry -Value "Running with local configuration.xml" -Severity 1
                        Copy-Item "$($PSScriptRoot)\configuration.xml" $SetupFolder -Force -ErrorAction Stop
                        Write-LogEntry -Value "Local Office Setup configuration filed copied" -Severity 1
                    }
                    #Starting Office setup with configuration file               
                    Try {
                        #Running office installer
                        Write-LogEntry -Value "Starting M365 Apps Install with Win32App method" -Severity 1
                        $OfficeInstall = Start-Process $SetupFilePath -ArgumentList "/configure $($SetupFolder)\configuration.xml" -Wait -PassThru -ErrorAction Stop
                    }
                    catch [System.Exception] {
                        Write-LogEntry -Value  "Error running the M365 Apps install. Errormessage: $($_.Exception.Message)" -Severity 3
                    }
                }
                else {
                    Throw "Error: Unable to verify setup file signature"
                }
            }
            catch [System.Exception] {
                Write-LogEntry -Value  "Error preparing office installation. Errormessage: $($_.Exception.Message)" -Severity 3
            }
        }
        catch {
            Write-LogEntry -Value  "Error during OEM Cleanup. Errormessage: $($_.Exception.Message)" -Severity 3
        }
          
       
    }
    catch [System.Exception] {
        Write-LogEntry -Value  "Error finding office setup file. Errormessage: $($_.Exception.Message)" -Severity 3
    }
    
}
catch [System.Exception] {
    Write-LogEntry -Value  "Error downloading office setup file. Errormessage: $($_.Exception.Message)" -Severity 3
}
#Cleanup 
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
}
Write-LogEntry -Value "M365 Apps setup completed" -Severity 1
# Complete