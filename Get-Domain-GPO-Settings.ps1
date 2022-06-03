<#
Author: Stan Crider
Date: 3June2022
What this crap does:
Gets GPO settings from specified domains and outputs in HTML format to specified folder
#>

# User variables
$DateLabel = Get-Date -Format yyMMdd
$LogFolder = "C:\Logs\GPO"
$Domains = @(
    "acme.com"
    "contoso.com"
)

# Script below
If(Test-Path $LogFolder){
    ForEach($Domain in $Domains){
        $DomainError = $null
        Try{
            $PrimaryDC = Get-ADDomain -Server $Domain | Select-Object PDCEmulator -ErrorAction Stop
        }
        Catch{
            $DomainError = "No PDC could be found for domain $Domain."
            Write-Warning $DomainError
        }
        If(!($DomainError)){
            $LogFileName = ($LogFolder + "\" + $Domain + "-" + $DateLabel + ".html")
            If(Test-Path $LogFileName){
                Write-Warning "The file $LogFileName already exists. Report for $Domain has been skipped."
            }
            Else{
                Try{
                    Get-GPOReport -All -Domain $Domain -Server $PrimaryDC.PDCEmulator -ReportType Html -Path $LogFileName -ErrorAction Stop
                }
                Catch{
                    Write-Warning "Unable to generate GPO report for $Domain"
                }
            }
        }
    }
}
Else{
    Write-Warning "The path $LogFolder does not exist. Script terminated."
}
