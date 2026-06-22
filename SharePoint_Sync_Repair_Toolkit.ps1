#requires -Version 5.1
<#
.SYNOPSIS
    Guarded SharePoint document library sync repair toolkit.
.DESCRIPTION
    Repairs common SharePoint library sync problems on Windows. SharePoint document
    libraries use the OneDrive sync engine, so the repairs reset or restart that
    engine, rebuild local Office cache data, refresh Explorer integration and repair
    supporting Windows services.
.NOTES
    Created by Dewald Pretorius - L2 IT Support Engineer.
    The tool does not delete cloud content or remove Microsoft 365 accounts. A
    OneDrive reset disconnects all sync connections and starts a full resynchronisation.
    Cache folders are moved into timestamped backups instead of being deleted.
#>

[CmdletBinding()]
param(
    [switch]$RepairAllSafe,
    [switch]$RestartOneDrive,
    [switch]$ResetOneDrive,
    [switch]$RebuildOfficeFileCache,
    [switch]$ClearTemporaryWebCache,
    [switch]$RestartWebClient,
    [switch]$RefreshExplorer,
    [switch]$FlushDns,
    [switch]$DryRun,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.1'
$RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'SharePoint_Sync_Repair_Logs'
}
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
$LogFile = Join-Path $OutputPath "SharePoint_Sync_Repair_$RunStamp.log"
$BackupRoot = Join-Path $OutputPath "Backup_$RunStamp"
New-Item -Path $BackupRoot -ItemType Directory -Force | Out-Null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DRYRUN')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN'    { Write-Host $Message -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        'DRYRUN'  { Write-Host "DRY RUN: $Message" -ForegroundColor Cyan }
        default   { Write-Host $Message }
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Confirm-Repair {
    param(
        [Parameter(Mandatory)][string]$Message,
        [switch]$HighImpact
    )
    if ($DryRun) { return $true }
    if ($HighImpact) {
        return (Read-Host "$Message Type REPAIR to continue") -eq 'REPAIR'
    }
    return (Read-Host "$Message Type YES to continue") -eq 'YES'
}

function Get-OneDriveExe {
    $paths = @(
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
        "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
    )
    foreach ($path in $paths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }
    return $null
}

function Stop-SyncApplications {
    $names = @('OneDrive','FileSyncHelper','OUTLOOK','WINWORD','EXCEL','POWERPNT','ONENOTE','MSACCESS','MSPUB')
    foreach ($name in $names) {
        $processes = @(Get-Process -Name $name -ErrorAction SilentlyContinue)
        foreach ($process in $processes) {
            if ($DryRun) {
                Write-Log "Would close $($process.ProcessName) process ID $($process.Id)." 'DRYRUN'
                continue
            }
            try { [void]$process.CloseMainWindow() } catch {}
        }
    }

    if (-not $DryRun) {
        Start-Sleep -Seconds 3
        foreach ($name in $names) {
            Get-Process -Name $name -ErrorAction SilentlyContinue |
                Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-SyncSnapshot {
    param([Parameter(Mandatory)][string]$Stage)

    $oneDriveExe = Get-OneDriveExe
    $accountsPath = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
    $officeCache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\OfficeFileCache'
    $syncRoots = @()

    $accountKeys = @(Get-ChildItem -LiteralPath $accountsPath -ErrorAction SilentlyContinue)
    foreach ($key in $accountKeys) {
        $item = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
        if ($item.UserFolder) {
            $syncRoots += [pscustomobject]@{
                Account = $key.PSChildName
                DisplayName = $item.DisplayName
                UserEmail = $item.UserEmail
                UserFolder = $item.UserFolder
                FolderExists = Test-Path -LiteralPath $item.UserFolder
            }
        }
    }

    $connectivity = foreach ($target in @('login.microsoftonline.com','www.office.com','sharepoint.com','onedrive.com')) {
        $dns = $false
        $https = $false
        try { [void][System.Net.Dns]::GetHostAddresses($target); $dns = $true } catch {}
        try { $https = Test-NetConnection -ComputerName $target -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue } catch {}
        [pscustomobject]@{ Target = $target; DNS = $dns; HTTPS443 = $https }
    }

    $version = $null
    if ($oneDriveExe) {
        try { $version = (Get-Item -LiteralPath $oneDriveExe).VersionInfo.FileVersion } catch {}
    }

    $snapshot = [ordered]@{
        Stage = $Stage
        Generated = (Get-Date).ToString('o')
        Computer = $env:COMPUTERNAME
        User = "$env:USERDOMAIN\$env:USERNAME"
        IsAdministrator = (Test-IsAdministrator)
        OneDrivePath = $oneDriveExe
        OneDriveVersion = $version
        OneDriveProcesses = @(
            Get-Process OneDrive, FileSyncHelper -ErrorAction SilentlyContinue |
                Select-Object Id, ProcessName, Path, StartTime
        )
        Accounts = $syncRoots
        OfficeFileCacheExists = (Test-Path -LiteralPath $officeCache)
        OfficeFileCachePath = $officeCache
        Services = @(
            Get-Service WebClient -ErrorAction SilentlyContinue |
                Select-Object Name, Status, StartType
        )
        Connectivity = $connectivity
    }

    $path = Join-Path $OutputPath "SharePoint_Sync_${Stage}_$RunStamp.json"
    $snapshot | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $path -Encoding UTF8
    Write-Log "Saved $Stage snapshot: $path" 'SUCCESS'
}

function Invoke-RestartOneDrive {
    $oneDriveExe = Get-OneDriveExe
    if (-not $oneDriveExe) {
        Write-Log 'OneDrive.exe was not found.' 'ERROR'
        return
    }
    if (-not (Confirm-Repair 'Restart the OneDrive sync engine used by SharePoint libraries?')) { return }

    if ($DryRun) {
        Write-Log "Would stop OneDrive and start $oneDriveExe." 'DRYRUN'
        return
    }

    Get-Process OneDrive, FileSyncHelper -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process -FilePath $oneDriveExe
    Write-Log 'OneDrive sync engine restarted.' 'SUCCESS'
}

function Invoke-ResetOneDrive {
    $oneDriveExe = Get-OneDriveExe
    if (-not $oneDriveExe) {
        Write-Log 'OneDrive.exe was not found.' 'ERROR'
        return
    }

    $warning = 'Reset the OneDrive sync engine? This disconnects every OneDrive and SharePoint sync connection, starts a full resynchronisation and may require folder selections to be configured again.'
    if (-not (Confirm-Repair $warning -HighImpact)) { return }

    if ($DryRun) {
        Write-Log "Would stop OneDrive, run $oneDriveExe /reset and restart OneDrive. All sync connections would be rebuilt." 'DRYRUN'
        return
    }

    Get-Process OneDrive, FileSyncHelper -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Process -FilePath $oneDriveExe -ArgumentList '/reset' -Wait
    Start-Sleep -Seconds 5
    if (-not (Get-Process OneDrive -ErrorAction SilentlyContinue)) {
        Start-Process -FilePath $oneDriveExe
    }
    Write-Log 'OneDrive reset completed. All OneDrive and SharePoint sync connections will perform a full resynchronisation.' 'SUCCESS'
}

function Invoke-RebuildOfficeFileCache {
    $cachePath = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\OfficeFileCache'
    if (-not (Test-Path -LiteralPath $cachePath)) {
        Write-Log 'The OfficeFileCache folder was not found. No cache rebuild was needed.' 'WARN'
        return
    }

    if (-not (Confirm-Repair 'Rebuild the local Office document cache? Save and synchronise all Office documents first.' -HighImpact)) { return }
    Stop-SyncApplications
    $destination = Join-Path $BackupRoot 'OfficeFileCache'

    if ($DryRun) {
        Write-Log "Would move $cachePath to $destination." 'DRYRUN'
        return
    }

    Move-Item -LiteralPath $cachePath -Destination $destination -Force
    Write-Log "Office document cache moved to $destination. Office will rebuild it." 'SUCCESS'

    $oneDriveExe = Get-OneDriveExe
    if ($oneDriveExe) {
        Start-Process -FilePath $oneDriveExe
        Write-Log 'OneDrive restarted after the Office cache rebuild.' 'SUCCESS'
    }
}

function Invoke-ClearTemporaryWebCache {
    if (-not (Confirm-Repair 'Clear temporary Windows internet files used by Microsoft 365 web authentication?')) { return }
    if ($DryRun) {
        Write-Log 'Would clear temporary WinINet internet files. Passwords and saved credentials would not be removed.' 'DRYRUN'
        return
    }

    Start-Process -FilePath rundll32.exe -ArgumentList 'InetCpl.cpl,ClearMyTracksByProcess 8' -Wait
    Write-Log 'Temporary Windows internet files cleared.' 'SUCCESS'
}

function Invoke-RestartWebClient {
    if (-not (Test-IsAdministrator)) {
        Write-Log 'Restarting the WebClient service requires Run as administrator.' 'WARN'
        return
    }
    if (-not (Confirm-Repair 'Restart the Windows WebClient service?')) { return }
    if ($DryRun) {
        Write-Log 'Would start or restart the WebClient service.' 'DRYRUN'
        return
    }

    $service = Get-Service WebClient -ErrorAction Stop
    Set-Service WebClient -StartupType Manual
    if ($service.Status -eq 'Running') {
        Restart-Service WebClient -Force
    } else {
        Start-Service WebClient
    }
    Write-Log 'Windows WebClient service is running.' 'SUCCESS'
}

function Invoke-RefreshExplorer {
    if (-not (Confirm-Repair 'Restart Windows Explorer to refresh SharePoint and OneDrive status icons?')) { return }
    if ($DryRun) {
        Write-Log 'Would restart Windows Explorer.' 'DRYRUN'
        return
    }

    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer.exe
    Write-Log 'Windows Explorer restarted.' 'SUCCESS'
}

function Invoke-FlushDns {
    if (-not (Confirm-Repair 'Flush the local DNS resolver cache?')) { return }
    if ($DryRun) {
        Write-Log 'Would run ipconfig.exe /flushdns.' 'DRYRUN'
        return
    }

    & ipconfig.exe /flushdns | Out-Null
    Write-Log 'DNS resolver cache flushed.' 'SUCCESS'
}

function Invoke-AllSafeRepairs {
    Write-Log 'Starting the safe SharePoint sync repair workflow.'
    Invoke-FlushDns
    Invoke-RestartOneDrive
    if (Test-IsAdministrator) {
        Invoke-RestartWebClient
    }
    Invoke-RefreshExplorer
}

function Show-Menu {
    do {
        Clear-Host
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-Host '  SHAREPOINT LIBRARY SYNC REPAIR TOOLKIT' -ForegroundColor Cyan
        Write-Host "  Version $ScriptVersion | Dewald Pretorius" -ForegroundColor DarkCyan
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-Host 'SharePoint libraries use the OneDrive sync engine.'
        Write-Host "Log: $LogFile"
        Write-Host "Dry run: $DryRun"
        Write-Host
        Write-Host ' 1. Run safe SharePoint sync repairs'
        Write-Host ' 2. Restart OneDrive sync engine'
        Write-Host ' 3. Reset OneDrive sync engine (disconnects and rebuilds all sync connections)'
        Write-Host ' 4. Rebuild Office document cache (backed up)'
        Write-Host ' 5. Clear temporary Microsoft 365 web cache'
        Write-Host ' 6. Restart Windows WebClient service'
        Write-Host ' 7. Refresh Windows Explorer sync integration'
        Write-Host ' 8. Flush DNS cache'
        Write-Host ' 0. Exit'
        $choice = Read-Host 'Select an option'

        try {
            switch ($choice) {
                '1' { Invoke-AllSafeRepairs }
                '2' { Invoke-RestartOneDrive }
                '3' { Invoke-ResetOneDrive }
                '4' { Invoke-RebuildOfficeFileCache }
                '5' { Invoke-ClearTemporaryWebCache }
                '6' { Invoke-RestartWebClient }
                '7' { Invoke-RefreshExplorer }
                '8' { Invoke-FlushDns }
                '0' { return }
                default { Write-Host 'Invalid selection.' -ForegroundColor Yellow }
            }
        } catch {
            Write-Log $_.Exception.Message 'ERROR'
        }

        if ($choice -ne '0') {
            Write-Host
            [void](Read-Host 'Press Enter to continue')
        }
    } while ($true)
}

Write-Log "SharePoint Sync Repair Toolkit $ScriptVersion started. DryRun=$DryRun"
Get-SyncSnapshot -Stage 'Before'

$repairSwitches = @(
    $RepairAllSafe, $RestartOneDrive, $ResetOneDrive, $RebuildOfficeFileCache,
    $ClearTemporaryWebCache, $RestartWebClient, $RefreshExplorer, $FlushDns
)

try {
    $hasSelectedRepair = @($repairSwitches | Where-Object { $_.IsPresent }).Count -gt 0
    if (-not $hasSelectedRepair) {
        Show-Menu
    } else {
        if ($RepairAllSafe)            { Invoke-AllSafeRepairs }
        if ($RestartOneDrive)          { Invoke-RestartOneDrive }
        if ($ResetOneDrive)            { Invoke-ResetOneDrive }
        if ($RebuildOfficeFileCache)   { Invoke-RebuildOfficeFileCache }
        if ($ClearTemporaryWebCache)   { Invoke-ClearTemporaryWebCache }
        if ($RestartWebClient)         { Invoke-RestartWebClient }
        if ($RefreshExplorer)          { Invoke-RefreshExplorer }
        if ($FlushDns)                 { Invoke-FlushDns }
    }
} catch {
    Write-Log $_.Exception.Message 'ERROR'
} finally {
    try { Get-SyncSnapshot -Stage 'After' } catch { Write-Log "Could not save final snapshot: $($_.Exception.Message)" 'WARN' }
    Write-Log "Repair workflow finished. Backups: $BackupRoot" 'SUCCESS'
    Write-Host "Logs and backups: $OutputPath" -ForegroundColor Green
}
