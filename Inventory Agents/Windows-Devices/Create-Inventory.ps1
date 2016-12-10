function WMIInv ([string]$res, [string]$Name, [string]$query, [string]$namespace = "root\cimv2")
{
    $out = ",`n `"$($Name)`":" 
    $out += Get-WmiObject -Namespace $namespace -Query $query -ea SilentlyContinue| select [A-Z]* -ExcludeProperty Scope,Options,ClassPath,Properties,SystemProperties,Qualifiers,Site,Container,PSComputerName, Path | ConvertTo-Json
    return $out
}


$res = " { id: `"" + (Get-WmiObject Win32_ComputerSystemProduct uuid).uuid + "`","
$res += "`n hostname: `"" + $env:COMPUTERNAME + "`","
$res += "`n InventoryDate: `"" + $(Get-Date -format u) + "`""

$res += WMIInv $res "Battery" "Select * FROM Win32_Battery" "root\cimv2"
$res += WMIInv $res "BIOS" "Select * FROM Win32_Bios" "root\cimv2"
$res += WMIInv $res "CDROMDrive" "Select * FROM Win32_CDROMDrive" "root\cimv2"
$res += WMIInv $res "ComputerSystem" "Select * FROM Win32_ComputerSystem" "root\cimv2"
$res += WMIInv $res "ComputerSystemProduct" "Select * FROM Win32_ComputerSystemProduct" "root\cimv2"
$res += WMIInv $res "DiskDrive" "Select * FROM Win32_DiskDrive" "root\cimv2"
$res += WMIInv $res "DiskPartition" "Select * FROM Win32_DiskPartition" "root\cimv2"
$res += WMIInv $res "Environment" "Select * FROM Win32_Environment" "root\cimv2"
$res += WMIInv $res "IDEController" "Select * FROM Win32_IDEController" "root\cimv2"
$res += WMIInv $res "NetworkAdapter" "Select * FROM Win32_NetworkAdapter" "root\cimv2"
$res += WMIInv $res "NetworkAdapterConfiguration" "Select * FROM Win32_NetworkAdapterConfiguration" "root\cimv2"
#$res += WMIInv $res "NetrkClient" "Select * FROM Win32_NetworkClient" "root\cimv2"
#$res += WMIInv $res "MotherboardDevice" "Select * FROM Win32_MotherboardDevice" "root\cimv2"
$res += WMIInv $res "OperatingSystem" "Select * FROM Win32_OperatingSystem" "root\cimv2"
#$res += WMIInv $res "Process" "Select * FROM Win32_Process" "root\cimv2"
$res += WMIInv $res "PhysicalMemory" "Select * FROM Win32_PhysicalMemory" "root\cimv2"
$res += WMIInv $res "PnpEntity" "Select * FROM Win32_PnpEntity" "root\cimv2"
$res += WMIInv $res "QuickFixEngineering" "Select * FROM Win32_QuickFixEngineering" "root\cimv2"
$res += WMIInv $res "Share" "Select * FROM Win32_Share" "root\cimv2"
$res += WMIInv $res "SoundDevice" "Select * FROM Win32_SoundDevice" "root\cimv2"
$res += WMIInv $res "Service" "Select * FROM Win32_Service" "root\cimv2"
$res += WMIInv $res "SystemEnclosure" "Select * FROM Win32_SystemEnclosure" "root\cimv2"
$res += WMIInv $res "VideoController" "Select * FROM Win32_VideoController" "root\cimv2"
$res += WMIInv $res "Volume" "Select * FROM Win32_Volume" "root\cimv2"


$res += ",`n `"Software`":" 
$SW = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ea SilentlyContinue | ? { $_.DisplayName -ne $null -and $_.SystemComponent -ne 0x1 -and $_.ParentDisplayName -eq $null } | Select DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString
$SW += Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ea SilentlyContinue | ? { $_.DisplayName -ne $null -and $_.SystemComponent -ne 0x1 -and $_.ParentDisplayName -eq $null } | Select DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString
$res += $SW | ConvertTo-Json

$res += ",`n `"Windows Updates`":" 
$objSearcher = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher();$objResults = $objSearcher.Search('IsHidden=0');
$res += $objResults.Updates | Select-Object -Property @{n='IsInstalled';e={$_.IsInstalled}},@{n='KB';e={$_.KBArticleIDs}},@{n='Bulletin';e={$_.SecurityBulletinIDs.Item(0)}},@{n='Title';e={$_.Title}},@{n='UpdateID';e={$_.Identity.UpdateID}},@{n='Revision';e={$_.Identity.RevisionNumber}},@{n='LastChange';e={$_.LastDeploymentChangeTime}} | ConvertTo-Json 

$res += "`n } "

$res > "$((Get-Location).Path)\$($env:COMPUTERNAME).json"
