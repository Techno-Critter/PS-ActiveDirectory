<#
Author: Stan Crider
Date: 8May2020
Crap: Gets a list of all computer objects in an AD environment
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

## Prepare script
# Configure variables
$Date = Get-Date -Format yyyyMMMdd
$Workbook = "C:\Temp\Computer Report $Date.xlsx"
$Searchbase = "DC=Acme,DC=com"

# Configure arrays
$ComputerArray = @()
$OperatingSystemArray = @()

## Functions
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

## Script: gather data from AD and add to array
If(Test-Path $Workbook){
    Write-Warning "The file $Workbook already exists. Script terminated."
}
Else{
    $Computers = Get-ADComputer -Filter * -SearchBase $Searchbase -Properties * | Sort-Object CanonicalName
    ForEach($Computer in $Computers){
        # Set computer account password age in days
        $PwdAge = $null
        If($null -ne $Computer.PasswordLastSet){
            $PwdAge = ((Get-Date)-$Computer.PasswordLastSet).Days
        }

        # Ping computer if account enabled to check if online
        $Pingable = "Not Tested"
        <#If($Computer.Enabled -eq $true){
            $Pingable = Test-Connection $Computer.Name -Count 2 -Quiet
        }#>

        # Add properties to array
        $ComputerArray += [PSCustomObject]@{
            "Name"               = $Computer.Name
            "IP Address"         = $Computer.IPv4Address
            "Enabled"            = $Computer.Enabled
            "Operating System"   = $Computer.OperatingSystem
            "Created"            = $Computer.Created
            "Last Logon"         = $Computer.LastLogonDate
            "Last Password"      = $Computer.PasswordLastSet
            "Password Age"       = $PwdAge
            "Online"             = $Pingable
            "Description"        = $Computer.Description
            "Location"           = $Computer.Location
            "Last Logon By"      = $Computer.employeeNumber
            "Domain Name"        = $Computer.CanonicalName
            "Distinguished Name" = $Computer.DistinguishedName
        }
    }

    # Count operating systems and amount of each OS
    $OperatingSystems = $ComputerArray | Select-Object "Operating System" | Sort-Object "Operating System" | Get-Unique -AsString
    ForEach($OS in $OperatingSystems){
        $OperatingSystemArray += [PSCustomObject]@{
            "OS"    = $OS.'Operating System'
            "Count" = ($ComputerArray | Where-Object{$_.'Operating System' -eq $OS.'Operating System'} | Measure-Object).Count
        }
    }

    ## Output
    # Computer worksheet
    $ComputerArrayColumnCount = Get-ColumnName ($ComputerArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ComputerArrayHeaderRow = "`$A`$1:`$$ComputerArrayColumnCount`$1"
    $ComputerArrayLastRow = ($ComputerArray | Measure-Object).Count + 1
    $EnabledColumn = "'Computers'!`$C`$2:`$C`$$ComputerArrayLastRow"
    $PasswordAgeColumn = "'Computers'!`$H`$2:`$H`$$ComputerArrayLastRow"

    $ComputerArrayStyle = @()
    $ComputerArrayStyle += New-ExcelStyle -Range "'Computers'!$ComputerArrayHeaderRow" -HorizontalAlignment Center

    $ComputerArrayConditionalFormatting = @()
    $ComputerArrayConditionalFormatting += New-ConditionalText -Range $EnabledColumn -ConditionalType ContainsText "FALSE" -ConditionalTextColor Brown -BackgroundColor Yellow
    $ComputerArrayConditionalFormatting += New-ConditionalText -Range $PasswordAgeColumn -ConditionalType GreaterThan 180 -ConditionalTextColor Maroon -BackgroundColor Pink

    $ComputerArray | Export-Excel -Path $Workbook -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Computers" -ConditionalText $ComputerArrayConditionalFormatting -Style $ComputerArrayStyle
    
    # OS Count worksheet
    $OSArrayColumnCount = Get-ColumnName ($OperatingSystemArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $OSArrayHeaderRow = "`$A`$1:`$$OSArrayColumnCount`$1"

    $OSArrayStyle = New-ExcelStyle -Range "'Computers'!$OSArrayHeaderRow" -HorizontalAlignment Center

    $OperatingSystemArray | Export-Excel -Path $Workbook -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "OS Count" -Style $OSArrayStyle
}
