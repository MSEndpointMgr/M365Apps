<#
.SYNOPSIS
  Script to install additional languagepacks as a Win32 App 

.DESCRIPTION
    Script to install additional languagepacks as a Win32 App by downloading the latest office setup.exe from evergreen url
    Running Setup.exe from downloaded files with provided config.xml file. 

.PARAMETER LanguageID
    Set the language ID in the correct formatting (like nb-no or en-us)
.PARAMETER Action 
    Supported actions are Install or Uninstall 

.EXAMPLE 
    InstallLanguagePacks.ps1 -LanguageID "nb-no" -Action Install
    InstallLanguagePacks.ps1 -LanguageID "nb-no" -Action Uninstall

.NOTES
  Version:          1.2
  Author:           Jan Ketil Skanke
  Creation Date:    23.11.2022
  Purpose/Change:   Initial script development
  Author:           Jan Ketil Skanke
  Contributor       Sandy Zeng 
  Contact:          @JankeSkanke @sandytsang
  Updated:          2022-23-11
    Version history:
    1.2 - (2022-23-11) Script created - Matching M365 Apps solution version
    1.2.1 - (2022-01-12) Adding function to validate signing on downloaded setup.exe
#>

#region parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    [string]$LanguageID,

    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall")]
    [string]$Action 
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
		} elseif ($Severity -eq 3) {
			Write-Warning -Message $Value
		}
	} catch [System.Exception] {
		Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}
function Start-DownloadFile {
    param(
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [parameter(Mandatory=$true)]
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
function Invoke-XMLUpdate  {
    param(
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LanguageID,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Filename,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
		[ValidateSet("Install", "Uninstall")]
        [string]$Action 
    )
    if ($Action -eq "Install"){
        $xmlDoc = [System.Xml.XmlDocument](Get-Content $FileName)
        $xmlDoc.Configuration.Add.Product.Language.ID = $LanguageID
        $xmlDoc.Save($FileName); 
    }
    else {
        $xmlDoc = [System.Xml.XmlDocument](Get-Content $FileName)
        $xmlDoc.Configuration.Remove.Product.Language.ID = $LanguageID
        $xmlDoc.Save($FileName);
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
#Endregion Functions

#Region Initialisations
$LogFileName = "M365AppsSetup.log"
#Endregion Initialisations
switch -Wildcard ($Action) { 
    {($PSItem -match "Install")}{
      $FileName = "install.xml"
    }
    {($PSItem -match "Uninstall")}{
        $FileName = "uninstall.xml"
    }
}

#Initate Install
Write-LogEntry -Value "Initiating LanguagePack $($LanguageID) $($Action) process" -Severity 1
#Attempt Cleanup of SetupFolder
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
}

$SetupFolder = (New-Item -ItemType "directory" -Path "$($env:SystemRoot)\Temp" -Name OfficeSetup -Force).FullName

try{
    #Download latest Office Deployment Toolkit
    $SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"
    
    try{
        #Start install preparations
        $SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
        if (-Not (Test-Path $SetupFilePath)) {
            Throw "Error: Setup file not found"
        }
        Write-LogEntry -Value "Setup file found at $($SetupFilePath)" -Severity 1
        try{
            #Prepare language pack installation or removal
            $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion 
            Write-LogEntry -Value "Office C2R Setup is running version $OfficeCR2Version" -Severity 1
            if (Invoke-FileCertVerification -FilePath $SetupFilePath){
                Invoke-XMLUpdate -LanguageID $LanguageID -Filename "$($PSScriptRoot)\$($Filename)" -Action $Action
                Copy-Item "$($PSScriptRoot)\$($Filename)" $SetupFolder -Force -ErrorAction Stop
                Write-LogEntry -Value "LanguagePack $($LanguageID) configuration file copied" -Severity 1           
                Try{
                    #Running office installer
                    Write-LogEntry -Value "Starting LanguagePack $($LanguageID) $($Action) with Win32App method" -Severity 1
                    $OfficeInstall = Start-Process $SetupFilePath -ArgumentList "/configure $($SetupFolder)\$($Filename)" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                }
                catch [System.Exception]{
                    Write-LogEntry -Value  "Error running the LanguagePack $($LanguageID) $($Action). Errormessage: $($_.Exception.Message)" -Severity 3
                }
            }
            else {
                Throw "Error: Unable to verify setup file signature"
            }

        }
        catch [System.Exception]{
            Write-LogEntry -Value  "Error preparing LanguagePack $($LanguageID) $($Action). Errormessage: $($_.Exception.Message)" -Severity 3
        }
    }
    catch [System.Exception]{
        Write-LogEntry -Value  "Error finding setup.exe Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3
    }
    
}
catch [System.Exception]{
    Write-LogEntry -Value  "Error downloading setup.exe from evergreen url. Errormessage: $($_.Exception.Message)" -Severity 3
}
#Cleanup 
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
}
Write-LogEntry -Value "LanguagePack $($LanguageID) $($Action) completed" -Severity 1
