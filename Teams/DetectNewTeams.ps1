#Detection 
$TeamsPackage = Get-AppPackage -AllUsers -name "MSTeams"
    If ($TeamsPackage){
        $Status = ($TeamsPackage | Select-Object PackageUserInformation).PackageUserInformation
        if ($Status.InstallState -match "Installed"){
            Write-host "Installed"
            Exit 0
        } 
        else { 
            "Not Installed"
            Exit 1
        }
    }
    else { 
        "Not Installed"
         Exit 1
    }    
   
