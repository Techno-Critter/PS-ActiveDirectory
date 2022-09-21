<#
Author: Stan Crider
Date: 15April2020
Crap: Get the properties of every user in specified domain or OU and outputs results to Excel
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#Requires -Modules ImportExcel

## Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Temp\ADUser_Report_$Date.xlsx"
$DomainNames = @(
    "prod.acme.com"
    "test.acme.com"
    "dev.acme.com"
)

# List properties to be gathered
$ADUserProperties = "Name",
                    "SamAccountName",
                    "SID",
                    "ObjectGUID",
                    "DisplayName",
                    "SurName",
                    "GivenName",
                    "Initials",
                    "Enabled",
                    "msDS-UserPasswordExpiryTimeComputed",
                    "LastLogonDate",
                    "EmployeeID",
                    "EmployeeType",
                    "DistinguishedName",
                    "CanonicalName",
                    "HomeDirectory",
                    "HomeDrive",
                    "Mail",
                    "PhysicalDeliveryOfficeName",
                    "Department",
                    "Description",
                    "Division",
                    "TelephoneNumber",
                    "Title",
                    "PersonalTitle",
                    "Company",
                    "Manager",
                    "streetAddress",
                    "postOfficeBox",
                    "l",
                    "st",
                    "postalCode",
                    "co",
                    "info"

## FUNCTIONS
#Convert number of object items into Excel column headers
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

# Split FQDN into Active Directory DC format
Function Get-ADDomainDistinguishedName{
    <#
    .SYNOPSIS
    Converts fully qualified domain name into Active Directory DC format

    .DESCRIPTION
    For use when both accessing Active Directory root structure and
    working with a domain fully qualified domain name is necessary.
    Especially useful when using an entire domain as a search base.
    Input: resource.acme.com
    Output: DC=resource,DC=acme,DC=com

    .PARAMETER DomainFQDN 
    A fully qualified domain name in the DOT format; example: resource.acme.com

    .EXAMPLE
    Get-ADDomainDistinguishedName -DomainFQDN 'resource.acme.com'

    .EXAMPLE
    'resource.acme.com','development.acme.com' | Get-ADDomainDistinguishedName

    .INPUTS
    String

    .OUTPUTS
    String

    .NOTES
    Author: Stan Crider
    Date:   5May2022
    Crap:   Yes, I wrote a function for 6 lines of code. Sue me.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='What is the root fully qualified domain name you would like to convert?')]

        [string]
        $DomainFQDN
    )

    Process{
        $DomainDistinguishedName = @()
        $DomainNameSplit = $DomainFQDN -split '\.'
        ForEach($DC in $DomainNameSplit){
            $DomainDistinguishedName += "DC=$DC"
        }
        $DomainDistinguishedName -join ","
    }
}

## Script below
# Check if logfile exists and terminate if it does
If(Test-Path $LogFile){
    Write-Warning "The file $LogFile already exists. Script terminated."
}
Else{
    # Stage array
    $UserArray = @()

    ForEach($Domain in $DomainNames){
        $SearchBase = Get-ADDomainDistinguishedName -DomainFQDN $Domain
        # Get users from specified location
        $UserProps = Get-ADUser -Server $Domain -SearchBase $SearchBase -Properties $ADUserProperties -Filter *
        ForEach($User in $UserProps){
            $LastLogonDays = "N/A"
            If($null -ne $User.LastLogonDate){
                $LastLogonDays = ((Get-Date) - ($User.LastLogonDate)).Days
            }

            $Manager = $null
            If($User.Manager){
                $Manager = (Get-ADuser -Server $Domain -Identity $User.Manager).Name
            }

            $POBox = $null
            If($User.postOfficeBox){
                $POBox = $User.postOfficeBox -join ", "
            }

            Try{
                $PWExpireDate = [DateTime]::FromFileTime($User."msDS-UserPasswordExpiryTimeComputed")
            }
            Catch{
                $PWExpireDate = $null
            }

            $UserArray += [PSCustomObject]@{
                "Name"            = $User.Name
                "Account Name"    = $User.SamAccountName
                "Display Name"    = $User.DisplayName
                "Last Name"       = $User.Surname
                "First Name"      = $User.GivenName
                "Initials"        = $User.Initials
                "EmployeeID"      = $User.EmployeeID
                "SID"             = $User.SID
                "GUID"            = $User.ObjectGUID
                "Enabled"         = $User.Enabled
                "Last Logon Date" = $User.LastLogonDate
                "Last Logon Days" = $LastLogonDays
                "PasswordExpires" = $PWExpireDate
                "Home Drive"      = $User.HomeDrive
                "Home Directory"  = $User.HomeDirectory
                "Email"           = $User.Mail
                "Phone"           = $User.TelephoneNumber
                "Office"          = $User.PhysicalDeliveryOfficeName
                "Department"      = $User.Department
                "Division"        = $User.Division
                "Description"     = $User.Description
                "Title"           = $User.Title
                "Personal Title"  = $User.PersonalTitle
                "Company"         = $User.Company
                "Manager"         = $Manager
                "Address"         = $User.streetAddress
                "PO Box"          = $POBox
                "City"            = $User.l
                "State"           = $User.st
                "ZIP"             = $User.PostalCode
                "Country"         = $User.co
                "Notes"           = $User.info
                "AD Path"         = $User.DistinguishedName
                "Canonical Name"  = $User.CanonicalName
                "Domain"          = $Domain
            }
        }
    }

    ## Output to Excel
    $UserSheetLastRow = ($UserArray | Measure-Object).Count + 1
    If($UserSheetLastRow -gt 1){
        # Define columns
        $UserSheetHeaderCount = Get-ColumnName ($UserArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $UserSheetHeaderRow   = "'Users'!`$A`$1:`$$UserSheetHeaderCount`$1"
        $UserEmployeeID       = "'Users'!`$G`$2:`$G`$$UserSheetLastRow"
        $UserEnabledColumn    = "'Users'!`$J`$2:`$J`$$UserSheetLastRow"
        $UserLastDaysColumn   = "'Users'!`$L`$2:`$L`$$UserSheetLastRow"

        # Format style for User sheet
        $UserSheetStyle = @()
        $UserSheetStyle += New-ExcelStyle -Range $UserSheetHeaderRow -HorizontalAlignment Center

        # Format text for User sheet
        $UserSheetConditionalText = @()
        $UserSheetConditionalText += New-ConditionalText -Range $UserEmployeeID -ConditionalType NotContainsText -BackgroundColor Wheat
        $UserSheetConditionalText += New-ConditionalText -Range $UserEnabledColumn -ConditionalType ContainsText "FALSE" -ConditionalTextColor Brown -BackgroundColor Wheat
        $UserSheetConditionalText += New-ConditionalText -Range $UserLastDaysColumn -ConditionalType GreaterThan 180 -ConditionalTextColor DarkRed -BackgroundColor LightPink

        # Create Excel standard configuration properties
        $ExcelProps = @{
            Autosize = $true;
            FreezeTopRow = $true;
            BoldTopRow = $true;
        }

        $ExcelProps.Path = $LogFile
        $ExcelProps.WorkSheetname = "Users"
        $ExcelProps.Style = $UserSheetStyle
        $ExcelProps.ConditionalFormat = $UserSheetConditionalText

        # Apply Style and Format, sort and output
        $UserArray | Sort-Object "Domain Name","Name" | Export-Excel @ExcelProps
    }
}
