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
$LogFileName = "TeamsSetup.log"
Write-LogEntry -Value "Initiating New Teams setup process" -Severity 1
#Attempt Cleanup of SetupFolder
if (Test-Path "$($env:SystemRoot)\Temp\TeamsSetup") {
    Remove-Item -Path "$($env:SystemRoot)\Temp\TeamsSetup" -Recurse -Force -ErrorAction SilentlyContinue
}
#Create SetupFolder
$SetupFolder = (New-Item -ItemType "directory" -Path "$($env:SystemRoot)\Temp" -Name TeamsSetup -Force).FullName

try {
    #Download latest teamsbootstrapper.exe
    $SetupEverGreenURL = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    Write-LogEntry -Value "Attempting to download latest teamsbootstrapper.exe executable" -Severity 1
    Start-DownloadFile -URL $SetupEverGreenURL -Path $SetupFolder -Name "teamsbootstrapper.exe" 
    try {
        #Start install preparations
        $SetupFilePath = Join-Path -Path $SetupFolder -ChildPath "teamsbootstrapper.exe"
        if (-Not (Test-Path $SetupFilePath)) {
            Throw "Error: Setup file not found"
        }
        Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1
        try {
            #Prepare Installation
            $SetupExeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$($SetupFolder)\teamsbootstrapper.exe").FileVersion 
            Write-LogEntry -Value "teamsbootstrapper.exe is running version $SetupExeVersion" -Severity 1
            if (Invoke-FileCertVerification -FilePath $SetupFilePath){
                #Starting Office setup with configuration file               
                Try {
                    #Running office installer
                    Write-LogEntry -Value "Starting Teams Install with Win32App method" -Severity 1
                    $Install = Start-Process $SetupFilePath -ArgumentList "-p" -WindowStyle Hidden -PassThru -ErrorAction Stop
                    $Install.WaitForExit()
                }
                catch [System.Exception] {
                    Write-LogEntry -Value  "Error running the Teams install. Errormessage: $($_.Exception.Message)" -Severity 3
                }
            }
            else {
                Throw "Error: Unable to verify setup file signature"
            }
        }
        catch [System.Exception] {
            Write-LogEntry -Value  "Error preparing installation. Errormessage: $($_.Exception.Message)" -Severity 3
        }
    }
catch [System.Exception] {
    Write-LogEntry -Value  "Error finding  setup file. Errormessage: $($_.Exception.Message)" -Severity 3
}

}
catch [System.Exception] {
Write-LogEntry -Value  "Error downloading office setup file. Errormessage: $($_.Exception.Message)" -Severity 3
}
#Cleanup 
if (Test-Path "$($env:SystemRoot)\Temp\TeamsSetup"){
    Remove-Item -Path "$($env:SystemRoot)\Temp\TeamsSetup" -Recurse -Force -ErrorAction SilentlyContinue
}
Write-LogEntry -Value "Teams setup completed" -Severity 1
# Complete