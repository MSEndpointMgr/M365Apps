<#
.SYNOPSIS
  Script to install Visio as a Win32 App during Autopilot

.DESCRIPTION
    Script to install Visio as a Win32 App during Autopilot by downloading the latest Office setup exe from evergreen url
    Running Setup.exe from downloaded files with provided config.xml file. 

.EXAMPLE
    Without external XML (Requires configuration.xml in the package)
    powershell.exe -executionpolicy bypass -file InstallVisio.ps1
    With external XML (Requires XML to be provided by URL)  
    powershell.exe -executionpolicy bypass -file InstallVisio.ps1 -XMLURL "https://mydomain.com/xmlfile.xml"

.NOTES
    Version: 1.2
    Author: Jan Ketil Skanke
    Contact:     @JankeSkanke
    Creation Date:  01.07.2021
    Updated: (2022-25-10)
    Version history:
        1.0.0 - (2022-23-10) Script released
        1.1.0 - (2022-25-10) Added support for External URL as parameter 
        1.2.0 - (2022-23-11) Moved from ODT download to Evergreen url for setup.exe 
#>
#region parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [string]$XMLUrl
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
#Endregion Functions

#Region Initialisations
$LogFileName = "M365AppsSetup.log"
#Endregion Initialisations

#Initate Install
Write-LogEntry -Value "Initiating Visio setup process" -Severity 1
#Attempt Cleanup of SetupFolder
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
}

$SetupFolder = (New-Item -ItemType "directory" -Path "$($env:SystemRoot)\Temp" -Name OfficeSetup -Force).FullName

try{
    #Download latest Office setup.exe
    $SetupEverGreenURL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "setup.exe"
    
    try{
        #Start install preparations
        $SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "setup.exe"
        if (-Not (Test-Path $SetupFilePath)) {
            Throw "Error: Setup file not found"
        }
        Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1
        try{
            #Prepare Office Installation
            $OfficeCR2Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$($SetupFolder)\setup.exe").FileVersion 
            Write-LogEntry -Value "Office C2R Setup is running version $OfficeCR2Version" -Severity 1

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
                    Write-LogEntry -Value "Visio setup failed" -Severity 3
                    exit 1
                }
            }
            else {
                #Local configuration file only 
                Write-LogEntry -Value "Running with local configuration.xml" -Severity 1
                Copy-Item "$($PSScriptRoot)\configuration.xml" $SetupFolder -Force -ErrorAction Stop
                Write-LogEntry -Value "Local Visio setup configuration filed copied" -Severity 1
            }
            #Starting Office setup with configuration file          
            Try{
                #Running office installer
                Write-LogEntry -Value "Starting Visio Install with Win32App method" -Severity 1
                $OfficeInstall = Start-Process $SetupFilePath  -ArgumentList "/configure $($SetupFolder)\configuration.xml" -Wait -PassThru -ErrorAction Stop
              }
            catch [System.Exception]{
                Write-LogEntry -Value  "Error running the Visio install. Errormessage: $($_.Exception.Message)" -Severity 3
            }
        }
        catch [System.Exception]{
            Write-LogEntry -Value  "Error preparing office installation. Errormessage: $($_.Exception.Message)" -Severity 3
        }
    }
    catch [System.Exception]{
        Write-LogEntry -Value  "Error extraction setup.exe from ODT Package. Errormessage: $($_.Exception.Message)" -Severity 3
    }
    
}
catch [System.Exception]{
    Write-LogEntry -Value  "Error downloading Office Deployment Toolkit. Errormessage: $($_.Exception.Message)" -Severity 3
}
#Cleanup 
if (Test-Path "$($env:SystemRoot)\Temp\OfficeSetup"){
    Remove-Item -Path "$($env:SystemRoot)\Temp\OfficeSetup" -Recurse -Force -ErrorAction SilentlyContinue
}
Write-LogEntry -Value "Visio setup completed" -Severity 1
