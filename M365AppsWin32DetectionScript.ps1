#DetectionScript
#For version 1.3 - enable OEM Cleanup with setting variable $CleanOEM = $true 
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
    if (Test-OfficeExists){
        #Check if Office with Clean OEM has already run
        if (Get-Item -Path "HKLM:SOFTWARE\$($DetectionRegKeyName)\M365AppsInstall\" -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object {$_.OemClean -eq "Yes"}){          
            # Path exist - cleanup should have already runned - Exit 0 
            Write-LogEntry -Value "OEM Cleanup already performed - M365 Apps Detected OK" -Severity 1
            Write-Output "Microsoft 365 Apps detected OK"
		    Exit 0
        }
        else {
            Write-LogEntry -Value "Microsoft 365 Apps found - checking if device is in Autopilot" -Severity 1
            if (Test-IsAutopilotDevice){
                # Device is an Autopilot device
                Write-LogEntry -Value "Device is an Autopilot Device" -Severity 1
                if (-not (Test-HasDeviceESPCompleted)){
                    # Device is in Device ESP Phase 
                    Write-LogEntry -Value "Device in ESP Device Phase - Initiate Installer with OEM Cleanup" -Severity 1
                    Exit 1
                }
                else {
                    Write-LogEntry -Value "Device is not in ESP Device Phase - Microsoft 365 Apps detected OK" -Severity 1
                    Write-Output "Microsoft 365 Apps detected OK"
                    Exit 0
                } 
            }
            else {
                Write-LogEntry -Value "Device is not an Autopilot Device - Microsoft 365 Apps detected OK" -Severity 1
                Write-Output "Microsoft 365 Apps detected OK"
                Exit 0
            }   
        }             
    }
    else{
		Write-LogEntry -Value "Microsoft 365 Apps not detected" -Severity 2
		Exit 1
    }   
}
else {
	if (Test-OfficeExists ){
		Write-LogEntry -Value "Microsoft 365 Apps detected OK" -Severity 1
		Write-Output "Microsoft 365 Apps detected OK"
		Exit 0
	}
	else {
		Write-LogEntry -Value "Microsoft 365 Apps noted detected" -Severity 2
		Exit 1
	}
}


