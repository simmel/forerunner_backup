<#
.SYNOPSIS
  Backup your .fit-files from your Garmin watch.
.DESCRIPTION
  For information on how to use this see Get-Help -Full on this file.
.PARAMETER Help
  Show this help!
.PARAMETER BackupPath
  Specifies the path to where we will put the backup from your Garmin Watch.
.PARAMETER DeviceName
  The MTP device name of your device e.g. "Forerunner 645 Music".
.EXAMPLE
  C:\PS> .\BackupForerunner.ps1 -BackupPath C:\backup -DeviceName "Forerunner 645 Music"
  Backup a Forerunner 645 Music to C:\backup
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
      <Arguments>-ExecutionPolicy Bypass -Command "C:\BackupForerunner.ps1 -BackupPath C:\backup -DeviceName 'Forerunner 645 Music' && pscp -r -p -batch -noagent -i C:\forerunner-backup.ppk -sftp C:\backup\ root@backup.domain.tld:/backup/forerunner/ && Invoke-RestMethod -TimeoutSec 5 -Uri https://hc-ping.com/meowpew"</Arguments>
    </Exec>
  </Actions>
</Task>
#>

Param(
  [Switch]$Help
  ,[Parameter()]
    [ValidateNotNullOrEmpty()]
    [String]$BackupPath
  ,[Parameter()]
    [ValidateNotNullOrEmpty()]
    [String]$DeviceName
)

if ($Help) {
  Get-Help -Full "$(Get-Location)\$($MyInvocation.MyCommand)"
  Exit 255
}

if (!$BackupPath -or !$DeviceName) {
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
    Write-Verbose "Creating: $path"
    New-Item -Path $path -ItemType Directory
  }
  else
  {
    Write-Verbose "Path $path already exist"
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

 Write-Host("Copying from '{0}' to '{1}'" -f $fullSourceDirPath, $destDirPath)

 $copiedCount = 0;

 foreach ($item in $sourceMtpDir.GetFolder.Items())
  {
   $itemName = ($item.Name)
   $fullFilePath = Join-Path -Path $destDirPath -ChildPath $itemName
   if(Test-Path $fullFilePath)
   {
      Write-Verbose "Element '$itemName' already exists"
   }
   else
   {
     $copiedCount++;
     Write-Verbose ("Copying #{0}: {1}{2}" -f $copiedCount, $fullSourceDirPath, $item.Name)
     $destDirShell.CopyHere($item, 1024)
   }
  }
  Write-Host "Copied '$copiedCount' elements from '$fullSourceDirPath'"
}

$phoneRootDir = Get-PhoneMainDir $DeviceName

$phoneCardPhotosSourceDir = Get-SubFolder $phoneRootDir "Primary\GARMIN\Activity"
Copy-FromPhone-ToDestDir $phoneCardPhotosSourceDir $BackupPath
