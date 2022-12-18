<#
.SYNOPSIS
  Backup your .fit-files from your Garmin watch.
.DESCRIPTION
  For information on how to use this see Get-Help -Full on this file.
.PARAMETER Help
  Show this help!
.PARAMETER BackupPath
  Specifies the path to where we will put the backup from your Garmin Watch.
.EXAMPLE
  C:\PS>
  <Description of example>
.NOTES
First enable Event logs for Microsoft-Windows-DriverFrameworks-UserMode/Operational, see
https://www.powershellmagazine.com/2013/07/15/pstip-how-to-enable-event-logs-using-windows-powershell/

In an Administrator Powershell:
$logName = 'Microsoft-Windows-DriverFrameworks-UserMode/Operational'
$log = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration $logName
$log.IsEnabled=$true
$log.SaveChanges()

Add a scheduled task from this XML:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2020-01-14T23:13:44.8701273</Date>
    <Author>simmel</Author>
    <URI>\Copy from Garmin</URI>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-DriverFrameworks-UserMode/Operational"&gt;&lt;Select Path="Microsoft-Windows-DriverFrameworks-UserMode/Operational"&gt;*[System[(EventID=2006)]] and *[UserData[UMDFHostAddDeviceEnd[InstanceId='USB\VID_091E&amp;amp;PID_4B48\0000EC9E2612']]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>lolguid</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass C:\BackupForerunner.ps1</Arguments>
    </Exec>
  </Actions>
</Task>
#>

Param(
  [Switch]$Help
  ,[Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String]$BackupPath
)

if ($Help) {
  Get-Help -Full "$(Get-Location)\$($MyInvocation.MyCommand)"
  Exit 255
}

# MTP code from https://github.com/nosalan/powershell-mtp-file-transfer/
# for an enhanced version that supports nested folders go to https://github.com/nosalan/powershell-mtp-file-transfer/blob/master/phone_backup_recursive.ps1

$ErrorActionPreference = "Stop"

function Create-Dir($path)
{
  if(! (Test-Path -Path $path))
  {
    Write-Host "Creating: $path"
    New-Item -Path $path -ItemType Directory
  }
  else
  {
    Write-Host "Path $path already exist"
  }
}


function Get-SubFolder($parentDir, $subPath)
{
  $result = $parentDir
  foreach($pathSegment in ($subPath -split "\\"))
  {
    $result = $result.GetFolder.Items() | Where-Object {$_.Name -eq $pathSegment} | select -First 1
    if($result -eq $null)
    {
      throw "Not found $subPath folder"
    }
  }
  return $result;
}


function Get-PhoneMainDir($phoneName)
{
  $o = New-Object -com Shell.Application
  $rootComputerDirectory = $o.NameSpace(0x11)
  $phoneDirectory = $rootComputerDirectory.Items() | Where-Object {$_.Name -eq $phoneName} | select -First 1

  if($phoneDirectory -eq $null)
  {
    throw "Not found '$phoneName' folder in This computer. Connect your phone."
  }

  return $phoneDirectory;
}


function Get-FullPathOfMtpDir($mtpDir)
{
 $fullDirPath = ""
 $directory = $mtpDir.GetFolder
 while($directory -ne $null)
 {
   $fullDirPath =  -join($directory.Title, '\', $fullDirPath)
   $directory = $directory.ParentFolder;
 }
 return $fullDirPath
}



function Copy-FromPhone-ToDestDir($sourceMtpDir, $destDirPath)
{
 Create-Dir $destDirPath
 $destDirShell = (new-object -com Shell.Application).NameSpace($destDirPath)
 $fullSourceDirPath = Get-FullPathOfMtpDir $sourceMtpDir

 Write-Host "Copying from: '" $fullSourceDirPath "' to '" $destDirPath "'"

 $copiedCount = 0;

 foreach ($item in $sourceMtpDir.GetFolder.Items())
  {
   $itemName = ($item.Name)
   $fullFilePath = Join-Path -Path $destDirPath -ChildPath $itemName
   if(Test-Path $fullFilePath)
   {
      Write-Host "Element '$itemName' already exists"
   }
   else
   {
     $copiedCount++;
     Write-Host ("Copying #{0}: {1}{2}" -f $copiedCount, $fullSourceDirPath, $item.Name)
     $destDirShell.CopyHere($item)
   }
  }
  Write-Host "Copied '$copiedCount' elements from '$fullSourceDirPath'"
}

# From https://karask.com/retry-powershell-invoke-webrequest/ but modified to use Invoke-RestMethod instead (since Invoke-WebRequest uses IE and I don't have it installed).
Function Req {
    Param(
        [Parameter(Mandatory=$True)]
        [hashtable]$Params,
        [int]$Retries = 3,
        [int]$SecondsDelay = 2
    )

    $method = $Params['Method']
    $url = $Params['Uri']

    $cmd = { Write-Host "$method $url... " -NoNewline; Invoke-RestMethod @Params }

    $retryCount = 0
    $completed = $false
    $response = $null

    while (-not $completed) {
        try {
            Invoke-Command $cmd -ArgumentList $Params
#            if ($response.StatusCode -ne 200) {
#                throw "Expecting reponse code 200, was: $($StatusCode)"
#            }
            $completed = $true
        } catch {
            if ($retrycount -ge $Retries) {
                Write-Warning "Request to $url failed the maximum number of $retryCount times."
                throw
            } else {
                Write-Warning "Request to $url failed. Retrying in $SecondsDelay seconds."
                Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__
                Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
                Start-Sleep $SecondsDelay
                $retrycount++
            }
        }
    }

    return $response
}


$phoneName = "Forerunner 645 Music" # MTP device name as it appears in This PC
$phoneRootDir = Get-PhoneMainDir $phoneName

$phoneCardPhotosSourceDir = Get-SubFolder $phoneRootDir "Primary\GARMIN\Activity"
Copy-FromPhone-ToDestDir $phoneCardPhotosSourceDir $BackupPath

pscp -r -p -batch -noagent -i C:\forerunner-backup.ppk -sftp C:\backup\ root@backup.domain.tld:/backup/forerunner/

Req -Params @{ 'Method'='GET';'TimeoutSec'='5';'Uri'='https://hc-ping.com/meowpew' }
