<#
Author: Stan Crider
Date: 11Mar2020
Crap:
Set ownership of home drives for enabled AD users in specified OU;
Sets ownership recursively from home directory to all child items
### Must have NTFSSecurity module installed!!!
### https://github.com/raandree/NTFSSecurity
#>

$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Temp\Set_Home_Ownership_Log_$Date.txt"
$ADLocation = "CN=Users,DC=acme,DC=com"

Write-Output ("Processing ownership script on $ADLocation at " + (Get-Date) + ".") | Out-File -FilePath $LogFile -Append
$UserProps = Get-ADUser -Filter * -SearchBase $ADLocation -Properties * -ErrorAction Stop | Sort-Object Name

ForEach($User in $UserProps){
    $HomeDirContents = $null
    If($null -ne $User.HomeDirectory){
        If((Test-Path $User.HomeDirectory) -and ($User.Enabled)){
            Write-Output "Processing $($User.Name)..." | Out-File -FilePath $LogFile -Append
            $HomeOwner = Get-NTFSOwner -Path $User.HomeDirectory
            If($HomeOwner.Account.Sid -ne $User.SID){
                Write-Output "Changing ownership of $($User.HomeDirectory) from $($HomeOwner.Account.AccountName) to $($User.Name)" | Out-File -FilePath $LogFile -Append
                Try{
                    Set-NTFSOwner -Path $User.HomeDirectory -Account $User.SID -ErrorAction Stop
                }
                Catch{
                    Write-Output $_.Exception.Message | Out-File -FilePath $LogFile -Append
                }
            }

            Try{
                $HomeDirContents = Get-ChildItem -Path $User.HomeDirectory -Recurse -ErrorAction Stop
            }
            Catch{
               Write-Output $_.Exception.Message | Out-File -FilePath $LogFile -Append
            }
            ForEach($FolderItem in $HomeDirContents){
                Try{
                    $FileOwner = Get-NTFSOwner -Path $FolderItem.FullName -ErrorAction Stop
                }
                Catch{
                    Write-Output $_.Exception.Message | Out-File -FilePath $LogFile -Append
                }
                If($FileOwner.Owner.Sid -ne $User.SID){
                    Try{
                        Set-NTFSOwner -Path $FolderItem.FullName -Account $User.SID -ErrorAction Stop
                    }
                    Catch{
                        Write-Output $_.Exception.Message | Out-File -FilePath $LogFile -Append
                    }
                }
            }
        }
    }
}
