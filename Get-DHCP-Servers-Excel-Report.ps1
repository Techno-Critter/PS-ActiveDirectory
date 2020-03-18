<#
Author: Stan Crider
Date: 4Apr2019
What this crap does:
Retreives DHCP servers, scopes and options from domain; outputs to Excel file
### Must have ImportExcel module installed! ###
###  https://github.com/dfinke/ImportExcel  ###
#>
#Requires -Modules ImportExcel

#region Variables
$Date = Get-Date -Format yyyyMMdd
$Workbook = "C:\Temp\DHCP_Servers_$Date.xlsx"
#endregion

# FUNCTION: Convert number of object items into Excel column headers
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

# If workbook already exists, terminate; else continue
If(Test-Path $Workbook){
    Write-Warning "The file $Workbook already exists. Script terminated."
}
Else{
#region Arrays
    $ServerSheet = @()
    $ADListSheet = @()
    $ScopeSheet = @()
    $ServerOptionsSheet = @()
    $ScopeOptionsSheet = @()
    $SuperScopeSheet = @()
    $Errors = @()

#endregion

    $ADListSheet = Get-DhcpServerInDC | Select-Object DnsName,IPAddress | Sort-Object DnsName

#region Get scopes for each server
    ForEach($Server in $ADListSheet){
        Write-Output ("Processing DHCP server $($Server.DnsName)...")

        If(-Not (Test-Connection -ComputerName $Server.DnsName -Quiet)){
            Write-Warning ("Server $($Server.DnsName) is not responding.")
            $Errors += [PSCustomObject]@{
                "Server Name" = $Server.DnsName
                "Error"       = "Server $($Server.DnsName) is not responding."
            }
            Continue
        }

        Try{
            $Scopes = Get-DhcpServerv4Scope -ComputerName $Server.DnsName -ErrorAction Stop
            $DNSSettings = Get-DhcpServerv4DnsSetting -ComputerName $Server.DnsName -ErrorAction Stop
            $Settings = Get-DhcpServerSetting -ComputerName $Server.DnsName -ErrorAction Stop
 
            $ServerSheet += [PSCustomObject]@{
                "Server Name"          = $Server.DnsName
                "Server IP"           = $Server.IPAddress
                "OS"                  = (Get-ADComputer -Identity ($Server.DnsName -split "\.")[0] -Properties Operatingsystem).OperatingSystem
                "Authorized"          = $Settings.IsAuthorized
                "Conflict Detections" = $Settings.ConflictDetectionAttempts
                "Dynamic DNS Updates" = $DNSSettings.DynamicUpdates
                "Update Old Clients"  = $DNSSettings.UpdateDnsRRForOlderClients
                "Delete DNS Expiry"   = $DNSSettings.DeleteDnsRROnLeaseExpiry
            }

            $ServerOptions = Get-DhcpServerv4OptionValue -ComputerName $Server.DnsName -ErrorAction Stop
            ForEach($SrvOption in $ServerOptions){
                $ServerOptionsSheet += [PSCustomObject]@{
                    "Server Name" = $Server.DnsName
                    "Name"       = $SrvOption.Name
                    "OptionID"   = $SrvOption.OptionID
                    "Type"       = $SrvOption.Type
                    "Value"      = ($SrvOption | Select-Object -ExpandProperty Value) -join ", "
                }
            }

            $SuperScopes = Get-DhcpServerv4Superscope -ComputerName $Server.DnsName -ErrorAction Stop | Select-Object SuperscopeName,ScopeId
            ForEach($SuperScope in $SuperScopes){
                If($SuperScope.SuperscopeName -ne ""){
                    ForEach($ScopeID in $SuperScope.ScopeID){
                        $SuperScopeMembers = [PSCustomObject]@{
                            "DHCPServer"     = $Server.DnsName
                            "SuperScopeName" = $SuperScope.SuperscopeName
                            "Member"         = $ScopeID.IPAddressToString
                        }
                        $SuperScopeSheet += $SuperScopeMembers
                    }
                }
            }
        
            ForEach($Scope in $Scopes){
                $ScopeSheet += [PSCustomObject]@{
                    "Server Name"         = $Server.DnsName
                    "Server IP"           = $Server.IPAddress
                    "Scope"               = $Scope.ScopeId
                    "Scope Name"          = $Scope.Name
                    "Subnet Mask"         = $Scope.SubnetMask
                    "Start Range"         = $Scope.StartRange
                    "End Range"           = $Scope.EndRange
                    "Dynamic DNS Updates" = (Get-DhcpServerv4DnsSetting -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId).DynamicUpdates
                    "Update Old Clients"  = (Get-DhcpServerv4DnsSetting -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId).UpdateDnsRRForOlderClients
                    "Delete DNS Expiry"   = (Get-DhcpServerv4DnsSetting -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId).DeleteDnsRROnLeaseExpiry
                }

                $EachScopeOptions = $null
                $EachScopeOptions = Get-DhcpServerv4OptionValue -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId -ErrorAction Stop

                ForEach($ScopeOption in $EachScopeOptions){
                    $ScopeOptionsSheet += [PSCustomObject]@{
                        "Server Name"  = $Server.DnsName
                        "Server IP"    = $Server.IPAddress
                        "Scope"        = $Scope.ScopeId
                        "Option Name"  = $ScopeOption.Name
                        "Option ID"    = $ScopeOption.OptionID
                        "Type"         = $ScopeOption.Type
                        "Value"        = $ScopeOption.Value -join ", "
                    }
                }
            }
#endregion

        }
        Catch{
            Write-Warning ("DHCP server " + $Server.DnsName + " is reporting an error.")
            Write-Warning ("Error: " + $_.Exception.Message)
            $Errors += [PSCustomObject]@{
                "Server Name" = $Server.DnsName
                "Error" = $_.Exception.Message
            }
        }
    }

#region Output to Excel

    # Server worksheet
    $ServerColumnCount = Get-ColumnName ($ServerSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ServerHeaderRow = "`$A`$1:`$$ServerColumnCount`$1"
    $ServerSheetLastRow = ($ServerSheet | Measure-Object).Count + 1

    If($ServerSheetLastRow -gt 1){
        $ServerSheetStyle = @()
        $ServerSheetStyle += New-ExcelStyle -Range "'Servers!'$ServerHeaderRow" -HorizontalAlignment Center

        $ServerSheetConditionalText = @()
        $ServerSheetConditionalText += New-ConditionalText -Range $ServerConflictDetectColumn -ConditionalType LessThan 3 -ConditionalTextColor Brown -BackgroundColor Wheat

        $ServerConflictDetectColumn = "'Servers'!`$F`$2:`$F`$$ServerSheetLastRow"
        $ServerSheet | Sort-Object ServerName | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Servers" -ConditionalText $ServerSheetConditionalText -Style $ServerSheetStyle
    }

    # AD List worksheet
    $ADListColumnCount = Get-ColumnName ($ADListSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ADListHeaderRow = "`$A`$1:`$$ADListColumnCount`$1"
    $ADListLastrow = ($ADListSheet | Measure-Object).Count + 1

    If($ADListLastrow -gt 1){
        $ADListSheetStyle = @()
        $ADListSheetStyle += New-ExcelStyle -Range "'AD List!'$ADListHeaderRow" -HorizontalAlignment Center

        $ADListSheet | Sort-Object DnsName | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "AD List" -Style $ADListSheetStyle
    }

    # Scope worksheet
    $ScopeColumnCount = Get-ColumnName ($ScopeSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ScopeHeaderRow = "`$A`$1:`$$ScopeColumnCount`$1"
    $ScopeSheetLastRow = ($ScopeSheet | Measure-Object).Count + 1

    If($ScopeSheetLastRow -gt 1){
        $ScopeSheetStyle = @()
        $ScopeSheetStyle += New-ExcelStyle -Range "'Scopes!'$ScopeHeaderRow" -HorizontalAlignment Center

        $ScopeSheet | Sort-Object ServerName,ScopeID | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Scopes" -Style $ScopeSheetStyle
    }

    # Server Options worksheet
    $ServerOptionsColumnCount = Get-ColumnName ($ServerOptionsSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ServerOptionsHeaderRow = "`$A`$1:`$$ServerOptionsColumnCount`$1"
    $ServerOptionsLastRow = ($ServerOptionsSheet | Measure-Object).Count + 1

    If($ServerOptionsLastRow -gt 1){
        $ServerOptionsStyle = @()
        $ServerOptionsStyle += New-ExcelStyle -Range "'Server Options!'$ServerOptionsHeaderRow" -HorizontalAlignment Center

        $ServerOptionsSheet | Sort-Object ServerName,OptionID | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Server Options" -Style $ServerOptionsStyle
    }

    # Scope Options worksheet
    $ScopeOptionsColumnCount = Get-ColumnName ($ScopeOptionsSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ScopeOptionsHeaderRow = "`$A`$1:`$$ScopeOptionsColumnCount`$1"
    $ScopeOptionsLastRow = ($ScopeOptionsSheet | Measure-Object).Count + 1

    If($ScopeOptionsLastRow -gt 1){
        $ScopeOptionsStyle = @()
        $ScopeOptionsStyle += New-ExcelStyle -Range "'Scope Options!'$ScopeOptionsHeaderRow" -HorizontalAlignment Center

        $ScopeOptionsSheet | Sort-Object ServerName,OptionID | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Scope Options" -Style $ScopeOptionsStyle
    }

    # SuperScope worksheet
    $SuperScopeColumnCount = Get-ColumnName ($SuperScopeSheet | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $SuperScopeHeaderRow = "`$A`$1:`$$SuperScopeColumnCount`$1"

    If($SuperScopeSheet){
        $SuperScopeStyle = @()
        $SuperScopeStyle += New-ExcelStyle -Range "'Super Scopes!'$SuperScopeHeaderRow"
        $SuperScopeSheet | Sort-Object DhcpServer,SuperScopeName,Member | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Super Scopes" -Style $SuperScopeStyle
    }

    # Error worksheet
    $ErrorColumnCount = Get-ColumnName ($Errors | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ErrorHeaderRow = "`$A`$1:`$$ErrorColumnCount`$1"

    If($Errors){
        $ErrorStyle = @()
        $ErrorStyle += New-ExcelStyle -Range "'Errors'$ErrorHeaderRow" -HorizontalAlignment Center

        $Errors | Sort-Object DnsName | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Errors" -Style $ErrorStyle
    }

#endregion
}