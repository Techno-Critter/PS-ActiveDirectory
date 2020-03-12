<#
Author: Stan Crider
Date: 6Feb2018
What this crap does:
Get a specified list of sites and subnets from AD
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#Requires -Module ImportExcel

#region Set logfile name
$Date = Get-Date -Format yyMMdd
$LogFile = "C:\Temp\Domain_Sites_$Date.xlsx"
#endregion

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
#endregion

#region Gather data
# Get subnets
$SubnetsRaw = Get-ADReplicationSubnet -Filter * -Properties Name,Location,Site,Description | Sort-Object Name | Select-Object Name, Location, Description, Site
$Subnets= @()
ForEach($Subnet in $SubnetsRaw){
    $SubnetObj = [PSCustomObject]@{
        Subnet = $Subnet.Name
        Location = $Subnet.Location
        Description = $Subnet.Description
        Site = Get-SitePartners $Subnet.Site #Change DistinguishedName to Name
    }
    $Subnets += $SubnetObj
}

# Get sites
$Sites = Get-ADReplicationSite -Properties * -Filter * | Sort-Object Name | Select-Object Name,Location,Description

# Get site links
$SiteLinksRaw = Get-ADReplicationSiteLink -Filter * -Properties * | Sort-Object Name | Select-Object Name,Description,Cost,ReplicationFrequencyInMinutes,SitesIncluded
$SiteLinks = @()
ForEach($SiteLink in $SiteLinksRaw){
    $SiteLinksObj = [PSCustomObject]@{
        Name = $SiteLink.Name
        Description = $SiteLink.Description
        Cost = $SiteLink.Cost
        RepFreq = $SiteLink.ReplicationFrequencyInMinutes
        MemberSites = Get-SitePartners $SiteLink.SitesIncluded #Change DistinguishedName to Name
        MemberCount = ($SiteLink.SitesIncluded | Measure-Object).Count
    }
    $SiteLinks += $SiteLinksObj
}
#endregion

#region Export to Excel
$Subnets | Sort-Object Subnet | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Subnets"
$Sites | Sort-Object Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Sites"
$SiteLinks | Sort-Object Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Links"
#endregion

#region Output to screen
Write-Output $Subnets | Format-Table
Write-Output $Sites | Format-Table
Write-Output $SiteLinks | Select-Object Name,Description,Cost,RepFreq,MemberCount,MemberSites | Format-Table
Write-Output ("Subnet count: " + ($Subnets | Measure-Object).Count)
Write-Output ("Site count: " + ($Sites | Measure-Object).Count)
Write-Output ("Link count: " + ($SiteLinks | Measure-Object).Count)
