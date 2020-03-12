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
    $ServerSheetLastRow = ($ServerSheet | Measure-Object).Count + 1
    If($ServerSheetLastRow -gt 1){
        $ServerConflictDetectColumn = "'Servers'!`$F`$2:`$F`$$ServerSheetLastRow"
        $ServerSheet | Sort-Object "Server Name" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Servers" -ConditionalText $(
            New-ConditionalText -Range $ServerConflictDetectColumn -ConditionalType LessThan 3 -ConditionalTextColor Brown -BackgroundColor Wheat
        )
    }
    $ADListSheet | Sort-Object "DnsName" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "AD List"
    $ScopeSheet | Sort-Object "Server Name","ScopeID" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Scopes"
    $ServerOptionsSheet | Sort-Object "Server Name","OptionID" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Server Options"
    $ScopeOptionsSheet | Sort-Object "Server Name","OptionID" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Scope Options"
    $SuperScopeSheet | Sort-Object "DhcpServer","SuperScopeName","Member" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Super Scopes"
    If($Errors){
        $Errors | Sort-Object "DnsName" | Export-Excel -Path $Workbook -FreezeTopRow -BoldTopRow -AutoSize -WorksheetName "Errors"
    }

#endregion
}
