<#
Author: Stan Crider
Date: 18Nov2019
Crap: Get domain controller properties and outputs to Excel report
### Must have ImportExcel module installed! ###
### https://github.com/dfinke/ImportExcel  ###
### Function stolen from: https://gallery.technet.microsoft.com/scriptcenter/Transpose-Object-cf517eb5
#>

#Requires -Module ImportExcel

# User variables
$DateName = Get-Date -Format yyyyMMdd
$Domain = "acme.com"
$LogFile = "C:\Temp\Domain Controllers\DC_$DateName.xlsx"

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

# Begin script; exit if logfile already exists
If(Test-Path $LogFile){
    Write-Host "The file $LogFile already exists. Script terminated."
}

# Get domain controllers and info from domain
Else{
    Import-Module ActiveDirectory
    $DomainControllers = Get-ADDomainController -Filter { Domain -eq $Domain } | Sort-Object HostName
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
        # Up-time
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

    # Export to Excel
    $LastRow = ($DCArray | Measure-Object).Count + 1
    $GC = "DCs!`$B`$2:`$B`$$LastRow"
    $Online = "DCs!`$G`$2:`$G`$$LastRow"
    $DCArray | Sort-Object Name | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "DCs" -ConditionalText $(
        New-ConditionalText -Range $GC -ConditionalType BeginsWith "FALSE" -ConditionalTextColor Brown -BackgroundColor Wheat
        New-ConditionalText -Range $Online -ConditionalType BeginsWith "FALSE" -ConditionalTextColor Maroon -BackgroundColor Pink
    )
    $DomainStatObj | Format-TransposeObject | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Status"
    If($ErrorArray){
        $ErrorArray | Sort-Object "Name" | Export-Excel -Path $LogFile -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Errors"
    }
}
