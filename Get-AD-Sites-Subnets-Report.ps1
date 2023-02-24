<#
Author: Stan Crider
Date: 6Feb2018
What this crap does:
Get a specified list of sites and subnets from AD
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#Requires -Module ImportExcel

## User variables
$Date = Get-Date -Format yyMMdd
$LogFile = "C:\Temp\Domain_Sites_$Date.xlsx"
$Domains = @('production.acme.com'
            'development.acme.com'
            'acme.com'
            )

## Functions
# Retrieve site partners and combine into string output
Function Get-SitePartners{
    <#
    .SYNOPSIS
    Converts AD object array into single comma-separated string

    .DESCRIPTION
    Takes a provided AD object property with multiple values and combines them into a single string

    .EXAMPLE
    Get-SitePartners -SiteIncludes (Get-ADReplicationSiteLink).StiesIncluded -DomainName 'development.acme.com'

    .INPUTS
    String array

    .OUTPUTS
    String

    .NOTES
    Author: Stan Crider
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,
            HelpMessage = 'What AD object property array would you like to convert to a single string?')]
        [string[]]
        $SiteIncludes,

        [Parameter(Mandatory,
            HelpMessage = 'What domain would you like to retrieve AD objects for?')]
        [string]
        $DomainName
    )

    Process{
        $SitePartners = @()
        ForEach($LSite in $SiteIncludes){
            $SitePartners += Get-ADReplicationSite -Server $DomainName -Identity $LSite
        }
        $Output = $SitePartners.Name -join ", "
        $Output
    }
}

# Convert number of object items into Excel column headers
Function Get-ColumnName ([int]$ColumnCount){
    <#
    .SYNOPSIS
    Converts integer into Excel column headers

    .DESCRIPTION
    Takes a provided number of columns in a table and converts the number into Excel header format
    Input: 27 - Output: AA
    Input: 2 - Ouput: B

    .EXAMPLE
    Get-ColumnName 27

    .INPUTS
    Integer

    .OUTPUTS
    String

    .NOTES
    Author: Stan Crider and Dennis Magee
    #>

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

## Script below
$Subnets = @()
$Sites = @()
$SiteLinks = @()

ForEach($Domain in $Domains){
    # Subnets
    $SubnetsRaw = Get-ADReplicationSubnet -Server $Domain -Filter * -Properties Name,Location,Site,Description | Sort-Object Name

    ForEach($Subnet in $SubnetsRaw){
            $SubnetSite = $null
            If($Subnet.Site){
                $SubnetSite = Get-SitePartners -SiteIncludes $Subnet.Site -DomainName $Domain
            }

        $Subnets += [PSCustomObject]@{
            'Subnet'      = $Subnet.Name
            'Location'    = $Subnet.Location
            'Description' = $Subnet.Description
            'Site'        = $SubnetSite
            'Domain'      = $Domain
        }
    }

    # Sites
    $SitesRaw = Get-ADReplicationSite -Server $Domain -Filter * -Properties Name,Location,Description | Sort-Object Name
    ForEach($SiteObj in $SitesRaw){
        $Sites += [PSCustomObject]@{
            'Name'        = $SiteObj.Name
            'Location'    = $SiteObj.Location
            'Description' = $SiteObj.Description
            'Domain'      = $Domain
        }
    }

    # Site links
    $SiteLinksRaw = Get-ADReplicationSiteLink -Server $Domain -Filter * -Properties Name,Description,Cost,ReplicationFrequencyInMinutes,SitesIncluded | Sort-Object Name
    ForEach($SiteLink in $SiteLinksRaw){
        $MemberSites = $null
        If($SiteLink.SitesIncluded){
            $MemberSites = Get-SitePartners -SiteIncludes $SiteLink.SitesIncluded -DomainName $Domain
        }

        $SiteLinks += [PSCustomObject]@{
            'Name'         = $SiteLink.Name
            'Description'  = $SiteLink.Description
            'Cost'         = $SiteLink.Cost
            'RepFreq'      = $SiteLink.ReplicationFrequencyInMinutes
            'Member Sites' = $MemberSites
            'Domain'       = $Domain
        }
    }
}

## Output
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
