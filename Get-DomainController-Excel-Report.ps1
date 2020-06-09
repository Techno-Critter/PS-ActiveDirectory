<#
Author: Stan Crider
Date: 18Nov2019
Crap: Get domain controller properties from local domain and outputs to Excel report
### Must have ImportExcel module installed! ###
### https://github.com/dfinke/ImportExcel  ###
### Transpose function stolen from: https://gallery.technet.microsoft.com/scriptcenter/Transpose-Object-cf517eb5
#>

#Requires -Module ImportExcel

# User variables
$DateName = Get-Date -Format yyyyMMdd
$Domain = (Get-ADDomain).DNSRoot
$LogFile = "C:\Temp\Domain Controllers\DC_$DateName.xlsx"

## FUNCTIONS
# Transpose object for adjusting output to Excel
Function Format-TransposeObject{
    [CmdletBinding()]
    Param([OBJECT][Parameter(ValueFromPipeline = $TRUE)]$InputObject)

    BEGIN{
        $Props = @()
        $PropNames = @()
        $InstanceNames = @()
    }

    PROCESS{
        If($Props.Length -eq 0){
            $PropNames = $InputObject.PSObject.Properties | Select-Object -ExpandProperty Name
            $InputObject.PSObject.Properties | ForEach-Object{
                $Props += New-Object -TypeName PSObject -Property @{Property = $_.Name }
            }
        }

        If($InputObject.Name){
            $Property = $InputObject.Name
        }
        Else{
            $Property = $InputObject | Out-String
        }

        If($InstanceNames -contains $Property){
            $COUNTER = 0
            Do{
                $COUNTER++
                $Property = "$($InputObject.Name) ({0})" -f $COUNTER
            }
            While($InstanceNames -contains $Property)
        }
        $InstanceNames += $Property

        $COUNTER = 0
        $PropNames | ForEach-Object{
            If($InputObject.($_)){
                $Props[$COUNTER] | Add-Member -Name $Property -Type NoteProperty -Value $InputObject.($_)
            }
            Else{
                $Props[$COUNTER] | Add-Member -Name $Property -Type NoteProperty -Value $null
            }
            $COUNTER++
        }
    }

    END{
        $Props
    }
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

## Begin script
# Exit if logfile already exists
If(Test-Path $LogFile){
    Write-Host "The file $LogFile already exists. Script terminated."
}
# Get domain controllers and info from domain
Else{
    Import-Module ActiveDirectory
    $DomainControllers = Get-ADDomainController -Filter * -Server $Domain | Sort-Object HostName
    $ForestProps = Get-ADForest
    $DomainProps = Get-ADDomain
    $GCCount = 0
    $DCUpCount = 0
    $DCCount = ($DomainControllers | Measure-Object).Count
    $DCArray = @()
    $ErrorArray = @()

    # Get info for each DC
    ForEach($DC in $DomainControllers){
        Write-Host ("Processing " + $DC.HostName + "...")
        $Uptime = $null
        $Reboot = $null
        $DCIPSettings = $null
        $DCDNSSettings = $null
        $TimeZone = $null
        $CurrentTime = $null
        $ADCompProp = Get-ADComputer -Identity $DC.Name -Properties Description
       # FSMO roles
        If($null -ne $DC.OperationMasterRoles){
            $FSMORoles = $DC.OperationMasterRoles -join ", "
        }
        # Ping
        $Online = (Test-Connection -ComputerName $DC.HostName -Quiet -Count 2)
        # Gather data
        If($Online){
            $DCUpCount++
            Try{
                $CimSession = New-CimSession -ComputerName $DC.HostName -ErrorAction Stop
                $DCIPSettings = Get-NetIPAddress -CimSession $CimSession -ErrorAction Stop | Where-Object{$_.IPAddress -eq $DC.IPv4Address}
                $DCDNSSettings = Get-DnsClientServerAddress -CimSession $CimSession -ErrorAction Stop | Where-Object{($_.InterfaceIndex -eq $DCIPSettings.InterfaceIndex) -and ($_.AddressFamily -eq $DCIPSettings.AddressFamily)}
                $Class = Get-CimInstance -Class Win32_OperatingSystem -CimSession $CimSession -ErrorAction Stop
                $TimeZone = Get-CimInstance -Class Win32_TimeZone -CimSession $CimSession -ErrorAction Stop
                $CurrentTime = Invoke-Command -ComputerName $DC.HostName -ScriptBlock {Get-Date -Format "MM/dd/yyyy HH:mm:ss"} -ErrorAction Stop
                $Reboot = $Class.LastBootUpTime
                $Uptime = (((Get-Date) - $Reboot).ToString("d' days 'hh':'mm':'ss"))
                Remove-CimSession -CimSession $CimSession -ErrorAction Stop
            }
            Catch{
                $ErrorArray += [PSCustomObject]@{
                    "Name"  = $DC.HostName
                    "Error" = $_.Exception.Message
                }
            }
        }
        Else{
            $ErrorArray += [PSCustomObject]@{
                "Name"  = $DC.HostName
                "Error" = "Not responding to ping"
            }
        }
        
        If($DC.IsGlobalCatalog){
            $GCCount++
        }
        # Add data to array
        $DCArray += [PSCustomObject]@{
            "Name"          = $DC.HostName
            "GC"            = $DC.IsGlobalCatalog
            "OS"            = $DC.OperatingSystem
            "IP"            = $DC.IPv4Address
            "DNS Addresses" = $DCDNSSettings.ServerAddresses -join ", "
            "Site"          = $DC.Site
            "Online"        = $Online
            "Enabled"       = $ADCompProp.Enabled
            "Uptime"        = $Uptime
            "LastBoot"      = $Reboot
            "Current Time"  = $CurrentTime
            "TimeZone"      = $TimeZone.Caption
            "Description"   = $ADCompProp.Description
            "FSMO"          = $FSMORoles
        }
    }

    $DomainStatObj = [PSCustomObject]@{
        "Name"         = $DomainProps.Name
        "Forest Name"  = $ForestProps.Name
        "Forest Level" = $ForestProps.ForestMode
        "Domain Name"  = $DomainProps.Name
        "Domain Level" = $DomainProps.DomainMode
        "DCs"          = $DCCount
        "GCs"          = $GCCount
        "Online"       = $DCUpCount
    }

    ## Export to Excel
    # Create Excel standard configuration properties
    $ExcelProps = @{
        Autosize = $true;
        FreezeTopRow = $true;
        BoldTopRow = $true;
    }

    $ExcelProps.Path = $LogFile

    # DC worksheet
    $DCArrayHeaderCount = Get-ColumnName ($DCArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $DCArrayHeaderRow = "`$A`$1:`$$DCArrayHeaderCount`$1"
    $LastRow = ($DCArray | Measure-Object).Count + 1
    $GC = "DCs!`$B`$2:`$B`$$LastRow"
    $Online = "DCs!`$G`$2:`$G`$$LastRow"

    $DCArrayStyle = @()
    $DCArrayStyle += New-ExcelStyle -Range "'DCs'$DCArrayHeaderRow" -HorizontalAlignment Center

    $DCArrayConditionalText = @()
    $DCArrayConditionalText += New-ConditionalText -Range $GC -ConditionalType BeginsWith "FALSE" -ConditionalTextColor Brown -BackgroundColor Wheat
    $DCArrayConditionalText += New-ConditionalText -Range $Online -ConditionalType BeginsWith "FALSE" -ConditionalTextColor Maroon -BackgroundColor Pink

    $DCArray | Sort-Object "Name" | Export-Excel @ExcelProps -WorksheetName "DCs" -ConditionalText $DCArrayConditionalText -Style $DCArrayStyle
    
    # Status worksheet
    $DomainStatHeaderCount = Get-ColumnName ($DomainStatObj | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $DomainStatHeaderRow = "`$A`$1:`$$DomainStatHeaderCount`$1"
    $DomainStatStyle = @()
    $DomainStatStyle += New-ExcelStyle -Range "'Status'$DomainStatHeaderRow" -HorizontalAlignment Center
    $DomainStatObj | Format-TransposeObject | Export-Excel @ExcelProps -WorksheetName "Status" -Style $DomainStatStyle
    
    # Errors worksheet
    If($ErrorArray){
        $ErrorArrayHeaderCount = Get-ColumnName ($DCArray | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
        $ErrorArrayHeaderRow = "`$A`$1:`$$ErrorArrayHeaderCount`$1"
        $ErrorArrayStyle = @()
        $ErrorArrayStyle += New-ExcelStyle -Range "'Errors'$ErrorArrayHeaderRow" -HorizontalAlignment Center
        $ErrorArray | Sort-Object "Name" | Export-Excel @ExcelProps -WorksheetName "Errors" -Style $ErrorArrayStyle
    }
}
