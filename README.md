# Backup Forerunner
A Powershell script to backup .fit-files from my Forerunner watch when it's attached

```
PS C:\> Get-Help .\BackupForerunner.ps1

NAME
    C:\BackupForerunner.ps1

SYNOPSIS
    Backup your .fit-files from your Garmin watch.


SYNTAX
    C:\BackupForerunner.ps1 [-Help] [[-BackupPath] <String>] [[-DeviceName] <String>] [<CommonParameter
    s>]


DESCRIPTION
    For information on how to use this see Get-Help -Full on this file.


RELATED LINKS

REMARKS
    To see the examples, type: "get-help C:\BackupForerunner.ps1 -examples".
    For more information, type: "get-help C:\BackupForerunner.ps1 -detailed".
    For technical information, type: "get-help C:\BackupForerunner.ps1 -full".
```

## Usage

`Get-Help -Full .\BackupForerunner.ps1` so see how to create a Scheduled task
which is triggered when the Garmin watch is connected to your computer.

If you're using Powershell >=7 you can also perform operations to e.g. copy the
backup to a NAS and signal to a monitoring system that you've done the backup:
```powershell
pwsh.exe -ExecutionPolicy Bypass -Command "C:\BackupForerunner.ps1 -BackupPath C:\backup -DeviceName 'Forerunner 645 Music' && pscp -r -p -batch -noagent -i C:\forerunner-backup.ppk -sftp C:\backup\ root@backup.domain.tld:/backup/forerunner/ && Invoke-RestMethod -TimeoutSec 5 -Uri https://hc-ping.com/meowpew"
```
