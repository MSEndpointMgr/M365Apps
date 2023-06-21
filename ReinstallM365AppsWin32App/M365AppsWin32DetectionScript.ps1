#DetectionScript
#For version 1.3 - enable OEM Cleanup with setting variable $CleanOEM = $true
#For version 1.4 - Remove detection of Autopilot and ESP, because this script is for reinstalling Office

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

# CleanOEM 
$CleanOEM = $true
# Script 
$LogFileName = "M365AppsSetup.log"
$DetectionRegKeyName = "MSEndpointMgr" # Only used if OEM Cleanup is enabled 

Write-LogEntry -Value "Start Office Install detection logic" -Severity 1
$RegistryKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$M365Apps = "Microsoft 365 Apps"

if ($CleanOEM){
    Write-LogEntry -Value "Testing for OEM Cleanup" -Severity 1
    #Check if Office with Clean OEM has already run 
    $checkOEMClean = Get-Item -Path "HKLM:SOFTWARE\$($DetectionRegKeyName)\M365AppsInstall\" -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object {$_.OemClean -eq "Yes"}

    #check if M365 Apps is installed and OEM Cleanup has not run
    if ($checkOEMClean -and (Test-OfficeExists)) {          
        Write-LogEntry -Value "OEM Cleanup already performed - M365 Apps Detected OK" -Severity 1
        Write-Output "Microsoft 365 Apps detected OK"
        Exit 0       
    }
    else{
		Write-LogEntry -Value "Microsoft 365 Apps not detected" -Severity 2
		Exit 1
    }   
}
else {
    Write-LogEntry -Value "Detecting Microsoft 365 Apps" -Severity 1
	if (Test-OfficeExists ){
		Write-LogEntry -Value "Microsoft 365 Apps detected OK" -Severity 1
		Write-Output "Microsoft 365 Apps detected OK"
		Exit 0
	}
	else {
		Write-LogEntry -Value "Microsoft 365 Apps noted detected" -Severity 2
        Write-Output "Microsoft 365 Apps noted detected"
		Exit 1
	}
}
