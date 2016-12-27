$accountName = "inventor"
$databaseName = "InventoryDB"
$collectionName = "Computers"
$connectionKey = "" #Place your Azure DocumentDB connection key here to automatically upload the data

#region Functions
function WMIInv ([string]$Name, [string]$query, [string]$namespace = "root\cimv2")
{
    $jsonout = ",`n `"$($Name)`":" 
    $val += Get-WmiObject -Namespace $namespace -Query $query -ea SilentlyContinue| select * -ExcludeProperty Scope,Options,ClassPath,Properties,SystemProperties,Qualifiers,Site,Container,PSComputerName, Path, __* | Sort | % {$_.} |ConvertTo-Json
    if($val -eq $null) { $val = "null" } 
    $jsonout += $val
    return $jsonout
}

#----
#---- from https://russellyoung.net/2016/06/18/managing-documentdb-with-powershell/
#----
    function GetKey([System.String]$Verb = '',[System.String]$ResourceId = '',
            [System.String]$ResourceType = '',[System.String]$Date = '',[System.String]$masterKey = '') {
        $keyBytes = [System.Convert]::FromBase64String($masterKey) 
        $text = @($Verb.ToLowerInvariant() + "`n" + $ResourceType.ToLowerInvariant() + "`n" + $ResourceId + "`n" + $Date.ToLowerInvariant() + "`n" + "`n")
        $body =[Text.Encoding]::UTF8.GetBytes($text)
        $hmacsha = new-object -TypeName System.Security.Cryptography.HMACSHA256 -ArgumentList (,$keyBytes) 
        $hash = $hmacsha.ComputeHash($body)
        $signature = [System.Convert]::ToBase64String($hash)
 
        [System.Web.HttpUtility]::UrlEncode($('type=master&ver=1.0&sig=' + $signature))
    }
 
    function GetUTDate() {
        $date = get-date
        $date = $date.ToUniversalTime();
        return $date.ToString("r", [System.Globalization.CultureInfo]::InvariantCulture);
    }
 
    function GetDatabases() {
        $uri = $rootUri + "/dbs"
 
        $hdr = BuildHeaders -resType dbs
 
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $hdr
        $response.Databases
 
        Write-Host ("Found " + $Response.Databases.Count + " Database(s)")
    }
 
    function GetCollections([string]$dbname){
        $uri = $rootUri + "/" + $dbname + "/colls"
        $headers = BuildHeaders -resType colls -resourceId $dbname
        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        $response.DocumentCollections
        Write-Host ("Found " + $Response.DocumentCollections.Count + " DocumentCollection(s)")
   }
 
    function BuildHeaders([string]$action = "get",[string]$resType, [string]$resourceId){
        $authz = GetKey -Verb $action -ResourceType $resType -ResourceId $resourceId -Date $apiDate -masterKey $connectionKey
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Authorization", $authz)
        $headers.Add("x-ms-version", '2015-12-16')
        $headers.Add("x-ms-date", $apiDate) 
        $headers
    }
 
    function PostDocument([string]$document, [string]$dbname, [string]$collection){
        $collName = "dbs/"+$dbname+"/colls/" + $collection
        $headers = BuildHeaders -action Post -resType docs -resourceId $collName
        $headers.Add("x-ms-documentdb-is-upsert", "true")
        $uri = $rootUri + "/" + $collName + "/docs"
     
        $response = Invoke-RestMethod $uri -Method Post -Body $document -ContentType 'application/json' -Headers $headers
        $response
    }

#----

#endregion

#region Inventory Classes
$json = " { `"id`": `"" + (Get-WmiObject Win32_ComputerSystemProduct uuid).uuid + "`","
$json += "`n `"hostname`": `"" + $env:COMPUTERNAME + "`","
$json += "`n `"InventoryDate`": `"" + $(Get-Date -format u) + "`""

$json += WMIInv "Battery" "Select * FROM Win32_Battery" "root\cimv2"
$json += WMIInv "BIOS" "Select * FROM Win32_Bios" "root\cimv2"
$json += WMIInv "CDROMDrive" "Select * FROM Win32_CDROMDrive" "root\cimv2"
$json += WMIInv "ComputerSystem" "Select * FROM Win32_ComputerSystem" "root\cimv2"
$json += WMIInv "ComputerSystemProduct" "Select * FROM Win32_ComputerSystemProduct" "root\cimv2"
$json += WMIInv "DiskDrive" "Select * FROM Win32_DiskDrive" "root\cimv2"
$json += WMIInv "DiskPartition" "Select * FROM Win32_DiskPartition" "root\cimv2"
$json += WMIInv "Environment" "Select * FROM Win32_Environment" "root\cimv2"
$json += WMIInv "IDEController" "Select * FROM Win32_IDEController" "root\cimv2"
$json += WMIInv "NetworkAdapter" "Select * FROM Win32_NetworkAdapter" "root\cimv2"
$json += WMIInv "NetworkAdapterConfiguration" "Select * FROM Win32_NetworkAdapterConfiguration" "root\cimv2"
#$json += WMIInv "NetrkClient" "Select * FROM Win32_NetworkClient" "root\cimv2"
#$json += WMIInv "MotherboardDevice" "Select * FROM Win32_MotherboardDevice" "root\cimv2"
$json += WMIInv "OperatingSystem" "Select * FROM Win32_OperatingSystem" "root\cimv2"
#$json += WMIInv "Process" "Select * FROM Win32_Process" "root\cimv2"
$json += WMIInv "PhysicalMemory" "Select * FROM Win32_PhysicalMemory" "root\cimv2"
$json += WMIInv "PnpEntity" "Select * FROM Win32_PnpEntity" "root\cimv2"
$json += WMIInv "QuickFixEngineering" "Select * FROM Win32_QuickFixEngineering" "root\cimv2"
$json += WMIInv "Share" "Select * FROM Win32_Share" "root\cimv2"
$json += WMIInv "SoundDevice" "Select * FROM Win32_SoundDevice" "root\cimv2"
$json += WMIInv "Service" "Select * FROM Win32_Service" "root\cimv2"
$json += WMIInv "SystemEnclosure" "Select * FROM Win32_SystemEnclosure" "root\cimv2"
$json += WMIInv "VideoController" "Select * FROM Win32_VideoController" "root\cimv2"
$json += WMIInv "Volume" "Select * FROM Win32_Volume" "root\cimv2"


$json += ",`n `"Software`":" 
$SW = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ea SilentlyContinue | ? { $_.DisplayName -ne $null -and $_.SystemComponent -ne 0x1 -and $_.ParentDisplayName -eq $null } | Select DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString
$SW += Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ea SilentlyContinue | ? { $_.DisplayName -ne $null -and $_.SystemComponent -ne 0x1 -and $_.ParentDisplayName -eq $null } | Select DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, UninstallString
$json += $SW | ConvertTo-Json

$json += ",`n `"Windows Updates`":"

$objSearcher = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher();
$objResults = $objSearcher.Search('IsHidden=0');
$upd += $objResults.Updates | Select-Object -Property @{n='IsInstalled';e={$_.IsInstalled}},@{n='KB';e={$_.KBArticleIDs}},@{n='Bulletin';e={$_.SecurityBulletinIDs.Item(0)}},@{n='Title';e={$_.Title}},@{n='UpdateID';e={$_.Identity.UpdateID}},@{n='Revision';e={$_.Identity.RevisionNumber}},@{n='LastChange';e={$_.LastDeploymentChangeTime}}
if($upd)
    {  $json += $upd  | ConvertTo-Json }
else { $json += "null"}

$json += "`n } "

#endregion

#region Upload to DocumentDB
#----
#---- from https://russellyoung.net/2016/06/18/managing-documentdb-with-powershell/
#----
if($connectionKey)
{
    $rootUri = "https://" + $accountName + ".documents.azure.com"
    write-host ("Root URI is " + $rootUri)
 
    #validate arguments
 
    $apiDate = GetUTDate
 
    $db = GetDatabases | where { $_.id -eq $databaseName }
 
    if ($db -eq $null) {
        write-error "Could not find database in account"
        return
    }

    $dbname = "dbs/" + $databaseName
    $collection = GetCollections -dbname $dbname | where { $_.id -eq $collectionName }
    
    if($collection -eq $null){
        write-error "Could not find collection in database"
        return
    }

    PostDocument -document $json -dbname $databaseName -collection $collectionName
}
else
{
    #Save as File if no connectionKey exists
    $json > "$((Get-Location).Path)\$($env:COMPUTERNAME).json"
}

#endregion