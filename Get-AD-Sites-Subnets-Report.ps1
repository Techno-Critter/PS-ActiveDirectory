<#
Author: Stan Crider
Date: 6Feb2018
What this crap does:
Get a specified list of sites and subnets from AD
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#Requires -Module ImportExcel

$Date = Get-Date -Format yyMMdd
$LogFile = "C:\Temp\Domain_Sites_$Date.xlsx"

#region Functions
# Pull Names of sites in ADReplication queries and combine into single string for visual output
Function Get-SitePartners($SiteIncludes){
    $SitePartners = @()
    ForEach($LSite in $SiteIncludes){
        $SitePartners += Get-ADReplicationSite -Identity $LSite
    }
    $Output = $SitePartners.Name -join ", "
    $Output
}

# Convert number of object items into Excel column headers
Function Get-ColumnName ([int]$ColumnCount){
    If(($ColumnCount -le 702) -and ($ColumnCount -ge 1)){
        $ColumnCount = [Math]::Floor($ColumnCount)
        $CharStart = 64
        $FirstCharacter = $null

        # Convert number into double letter column name (AA-ZZ)
        If($ColumnCount -gt 26){
            $FirstNumber = [Math]::Floor(($ColumnCount)/26)
            $SecondNumber = ($ColumnCount) % 26

            # Reset increment for base-26
            If($SecondNumber -eq 0){
                $FirstNumber--
                $SecondNumber = 26
            }

            # Left-side column letter (first character from left to right)
            $FirstLetter = [int]($FirstNumber + $CharStart)
            $FirstCharacter = [char]$FirstLetter

            # Right-side column letter (second character from left to right)
            $SecondLetter = $SecondNumber + $CharStart
            $SecondCharacter = [char]$SecondLetter

            # Combine both letters into column name
            $CharacterOutput = $FirstCharacter + $SecondCharacter
        }

        # Convert number into single letter column name (A-Z)
        Else{
            $CharacterOutput = [char]($ColumnCount + $CharStart)
        }
    }
    Else{
        $CharacterOutput = "ZZ"
    }

    # Output column name
    $CharacterOutput
}
#endregion

#region Subnets
$SubnetsRaw = Get-ADReplicationSubnet -Filter * -Properties Name,Location,Site,Description | Sort-Object Name | Select-Object Name, Location, Description, Site
$Subnets= @()
ForEach($Subnet in $SubnetsRaw){
    $Subnets += [PSCustomObject]@{
        "Subnet"      = $Subnet.Name
        "Location"    = $Subnet.Location
        "Description" = $Subnet.Description
        "Site"        = Get-SitePartners $Subnet.Site #Change DistinguishedName to Name
    }
}
#endregion

# Get Sites
$Sites = Get-ADReplicationSite -Properties * -Filter * | Sort-Object Name | Select-Object Name,Location,Description

#region Site Links
$SiteLinksRaw = Get-ADReplicationSiteLink -Filter * -Properties * | Sort-Object Name | Select-Object Name,Description,Cost,ReplicationFrequencyInMinutes,SitesIncluded
$SiteLinks = @()
ForEach($SiteLink in $SiteLinksRaw){
    $SiteLinks += [PSCustomObject]@{
        "Name"         = $SiteLink.Name
        "Description"  = $SiteLink.Description
        "Cost"         = $SiteLink.Cost
        "RepFreq"      = $SiteLink.ReplicationFrequencyInMinutes
        "Member Sites" = Get-SitePartners $SiteLink.SitesIncluded #Change DistinguishedName to Name
    }
}
#endregion

#region Output to Excel
# Create Excel standard configuration properties
$ExcelProps = @{
    Autosize = $true;
    FreezeTopRow = $true;
    BoldTopRow = $true;
}

$ExcelProps.Path = $LogFile

# Subnets sheet
$SubnetHeaderCount = Get-ColumnName ($Subnets | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$SubnetHeaderRow = "`$A`$1:`$$SubnetHeaderCount`$1"
$SubnetSheetStyle = New-ExcelStyle -Range "'File Servers'$SubnetHeaderRow" -HorizontalAlignment Center
$Subnets | Export-Excel @ExcelProps -WorkSheetname "Subnets" -Style $SubnetSheetStyle

# Sites sheet
$SitesHeaderCount = Get-ColumnName ($Sites | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$SitesHeaderRow = "`$A`$1:`$$SitesHeaderCount`$1"
$SitesSheetStyle = New-ExcelStyle -Range "'File Servers'$SitesHeaderRow" -HorizontalAlignment Center
$Sites | Export-Excel @ExcelProps -WorkSheetname "Sites" -Style $SitesSheetStyle

# SiteLinks sheet
$SiteLinkHeaderCount = Get-ColumnName ($SiteLinks | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$SiteLinkHeaderRow = "`$A`$1:`$$SiteLinkHeaderCount`$1"
$SiteLinkSheetStyle = New-ExcelStyle -Range "'File Servers'$SiteLinkHeaderRow" -HorizontalAlignment Center
$SiteLinks | Export-Excel @ExcelProps -WorkSheetname "Links" -Style $SiteLinkSheetStyle
#endregion

#region Output to screen
Write-Output $Subnets
Write-Output ("Subnet count: " + ($Subnets | Measure-Object).Count)
Write-Output $Sites
Write-Output ("Site count: " + ($Sites | Measure-Object).Count)
Write-Output $SiteLinks
Write-Output ("Links count: " + ($SiteLinks | Measure-Object).Count)
#endregion
