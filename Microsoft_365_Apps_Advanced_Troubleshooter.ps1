#requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365 Apps Advanced Troubleshooter.
.DESCRIPTION
    Menu-driven diagnostics and guided repair for Teams, OneDrive, SharePoint,
    Outlook, Microsoft 365 identity, Office Click-to-Run, Store apps and Windows.
.NOTES
    Created by Dewald Pretorius.
    Runs in standard-user mode unless manually started as Administrator.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ScriptVersion = '3.0.0'
$ScriptAuthor = 'Dewald Pretorius'

# Best-effort removal of Windows downloaded-file security markers.
try {
    $selfPath = $MyInvocation.MyCommand.Path
    $launcherPath = Join-Path (Split-Path -Parent $selfPath) 'Launch_Microsoft_365_Troubleshooter.bat'
    foreach ($target in @($selfPath, $launcherPath)) {
        if ($target -and (Test-Path -LiteralPath $target)) {
            Unblock-File -LiteralPath $target -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $target -Stream Zone.Identifier -ErrorAction SilentlyContinue
        }
    }
} catch {}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$Desktop = [Environment]::GetFolderPath('Desktop')
$LogRoot = Join-Path $Desktop 'Microsoft_365_Troubleshooter_Logs'
New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
$LogFile = Join-Path $LogRoot ("Troubleshooter_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level='INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN' { Write-Host $Message -ForegroundColor Yellow }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message }
    }
}

function Pause-Menu { Write-Host; [void](Read-Host 'Press Enter to continue') }
function Confirm-Action { param([string]$Message); (Read-Host "$Message Type YES to continue") -eq 'YES' }

function Show-Header {
    Clear-Host
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '  Microsoft 365 Apps Advanced Troubleshooter' -ForegroundColor Cyan
    Write-Host "  Created by $ScriptAuthor" -ForegroundColor DarkCyan
    Write-Host "  Version $ScriptVersion" -ForegroundColor DarkCyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ("Computer : {0}" -f $env:COMPUTERNAME)
    Write-Host ("User     : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Host ("Admin    : {0}" -f (Test-IsAdministrator))
    Write-Host ("Log      : {0}" -f $LogFile)
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host
}

function Export-Report {
    param([string]$Name,[scriptblock]$ScriptBlock)
    $path = Join-Path $LogRoot ("{0}_{1}.txt" -f $Name,(Get-Date -Format 'yyyyMMdd_HHmmss'))
    & $ScriptBlock 2>&1 | Out-String -Width 300 | Set-Content -Path $path -Encoding UTF8
    Write-Log "Report saved: $path" 'SUCCESS'
}

function Get-OneDriveExe {
    $paths = @(
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
        "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
    )
    $paths | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}

function Get-OfficeAppPath {
    param([string]$Exe)
    $command = Get-Command $Exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    $roots = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16",
        "$env:ProgramFiles\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16"
    )
    foreach ($root in $roots) {
        if ($root) {
            $candidate = Join-Path $root $Exe
            if (Test-Path $candidate) { return $candidate }
        }
    }
}

function Test-M365Connectivity {
    Show-Header
    $targets = @(
        'login.microsoftonline.com','www.office.com','graph.microsoft.com',
        'outlook.office.com','teams.microsoft.com','sharepoint.com','onedrive.live.com'
    )
    $results = foreach ($target in $targets) {
        $dns = $false; $https = $false
        try { $dns = [bool](Resolve-DnsName $target -ErrorAction Stop | Select-Object -First 1) } catch {}
        try { $https = Test-NetConnection $target -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue } catch {}
        [pscustomobject]@{Target=$target;DNS=$dns;HTTPS443=$https}
    }
    $results | Format-Table -AutoSize
    $results | Export-Csv (Join-Path $LogRoot ("Connectivity_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))) -NoTypeInformation
    Pause-Menu
}

function Get-GeneralDiagnostics {
    Show-Header
    Export-Report -Name 'General_Diagnostics' -ScriptBlock {
        'MICROSOFT 365 GENERAL DIAGNOSTICS'
        "Generated: $(Get-Date)"
        "Created by: $ScriptAuthor"
        ''
        'OPERATING SYSTEM'
        Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,BuildNumber,LastBootUpTime | Format-List
        'DISK'
        Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Select-Object DeviceID,@{n='FreeGB';e={[math]::Round($_.FreeSpace/1GB,2)}},@{n='SizeGB';e={[math]::Round($_.Size/1GB,2)}} | Format-Table -AutoSize
        'OFFICE CLICK-TO-RUN'
        Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue | Select-Object ProductReleaseIds,VersionToReport,Platform,UpdateChannel,CDNBaseUrl | Format-List
        'SERVICES'
        Get-Service ClickToRunSvc,wuauserv,bits,TokenBroker -ErrorAction SilentlyContinue | Select-Object Name,Status,StartType | Format-Table -AutoSize
        'DEVICE REGISTRATION'
        dsregcmd.exe /status
        'PROXY'
        netsh winhttp show proxy
        'RECENT APPLICATION ERRORS'
        Get-WinEvent -FilterHashtable @{LogName='Application';Level=2;StartTime=(Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue | Select-Object -First 40 TimeCreated,Id,ProviderName,Message | Format-List
    }
    Pause-Menu
}

function Stop-M365Processes {
    Show-Header
    $names = 'OUTLOOK','WINWORD','EXCEL','POWERPNT','ONENOTE','Teams','ms-teams','OneDrive'
    foreach ($name in $names) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force; Write-Log "Stopped $($_.ProcessName)." 'SUCCESS' }
            catch { Write-Log "Could not stop $($_.ProcessName): $($_.Exception.Message)" 'WARN' }
        }
    }
    Pause-Menu
}

function Reset-OneDrive {
    Show-Header
    $exe = Get-OneDriveExe
    if (-not $exe) { Write-Log 'OneDrive executable not found.' 'ERROR'; Pause-Menu; return }
    if (-not (Confirm-Action 'Reset the local OneDrive client?')) { return }
    Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
    Start-Process $exe -ArgumentList '/reset' -Wait
    Start-Sleep 3
    Start-Process $exe
    Write-Log 'OneDrive reset command completed.' 'SUCCESS'
    Pause-Menu
}

function Test-OneDriveSharePoint {
    Show-Header
    Export-Report -Name 'OneDrive_SharePoint_Diagnostics' -ScriptBlock {
        'ONEDRIVE AND SHAREPOINT DIAGNOSTICS'
        "Created by: $ScriptAuthor"
        ''
        'ONEDRIVE PROCESS'
        Get-Process OneDrive -ErrorAction SilentlyContinue | Select-Object Id,Path,StartTime,CPU | Format-Table -AutoSize
        'ONEDRIVE SETTINGS'
        Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue | Select-Object PSChildName | Format-Table -AutoSize
        'KNOWN FOLDER PATHS'
        Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction SilentlyContinue | Format-List
        'SYNC ROOTS'
        Get-ChildItem $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object Name -match 'OneDrive|SharePoint' | Select-Object FullName | Format-Table -AutoSize
        'INVALID OR LONG PATH CANDIDATES'
        Get-ChildItem $env:USERPROFILE -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName.Length -gt 240 -or $_.Name -match '["*:<>?/\\|]' } | Select-Object -First 100 FullName,@{n='Length';e={$_.FullName.Length}} | Format-Table -AutoSize
        'CONNECTIVITY'
        foreach($hostName in 'www.office.com','login.microsoftonline.com','sharepoint.com','graph.microsoft.com') {
            [pscustomobject]@{Target=$hostName;HTTPS443=(Test-NetConnection $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue)}
        }
    }
    Pause-Menu
}

function Clear-TeamsCache {
    Show-Header
    if (-not (Confirm-Action 'Close Teams and clear local Teams cache?')) { return }
    Stop-Process -Name Teams,ms-teams -Force -ErrorAction SilentlyContinue
    $paths = @(
        "$env:APPDATA\Microsoft\Teams",
        "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"
    )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Get-ChildItem $path -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Cleared Teams cache: $path" 'SUCCESS'
        }
    }
    Pause-Menu
}

function Test-Teams {
    Show-Header
    Export-Report -Name 'Teams_Diagnostics' -ScriptBlock {
        'MICROSOFT TEAMS DIAGNOSTICS'
        "Created by: $ScriptAuthor"
        'PROCESS'
        Get-Process Teams,ms-teams -ErrorAction SilentlyContinue | Select-Object Id,ProcessName,Path,StartTime,CPU | Format-Table -AutoSize
        'APP PACKAGE'
        Get-AppxPackage -Name MSTeams -ErrorAction SilentlyContinue | Select-Object Name,Version,InstallLocation,Status | Format-List
        'AUDIO DEVICES'
        Get-CimInstance Win32_SoundDevice -ErrorAction SilentlyContinue | Select-Object Name,Status,Manufacturer | Format-Table -AutoSize
        'CAMERAS'
        Get-PnpDevice -Class Camera -ErrorAction SilentlyContinue | Select-Object FriendlyName,Status,InstanceId | Format-Table -AutoSize
        'PRIVACY SETTINGS'
        Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone' -ErrorAction SilentlyContinue | Format-List
        Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam' -ErrorAction SilentlyContinue | Format-List
        'CONNECTIVITY'
        foreach($hostName in 'teams.microsoft.com','login.microsoftonline.com','statics.teams.cdn.office.net') {
            [pscustomobject]@{Target=$hostName;HTTPS443=(Test-NetConnection $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue)}
        }
        'RECENT TEAMS EVENTS'
        Get-WinEvent -FilterHashtable @{LogName='Application';StartTime=(Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue | Where-Object Message -match 'Teams|MSTeams|WebView2' | Select-Object -First 50 TimeCreated,Id,ProviderName,LevelDisplayName,Message | Format-List
    }
    Pause-Menu
}

function Test-Outlook {
    Show-Header
    Export-Report -Name 'Outlook_Diagnostics' -ScriptBlock {
        'OUTLOOK DIAGNOSTICS'
        "Created by: $ScriptAuthor"
        'PROCESS'
        Get-Process OUTLOOK -ErrorAction SilentlyContinue | Select-Object Id,Path,StartTime,CPU | Format-Table -AutoSize
        'PROFILES'
        Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles' -ErrorAction SilentlyContinue | Select-Object PSChildName | Format-Table -AutoSize
        'OST/PST FILES'
        Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Outlook" -File -ErrorAction SilentlyContinue | Select-Object Name,Length,LastWriteTime | Format-Table -AutoSize
        'ADD-INS'
        Get-ChildItem 'HKCU:\Software\Microsoft\Office\Outlook\Addins' -ErrorAction SilentlyContinue | Select-Object PSChildName | Format-Table -AutoSize
        'AUTODISCOVER'
        Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover' -ErrorAction SilentlyContinue | Format-List
        'CONNECTIVITY'
        foreach($hostName in 'outlook.office.com','autodiscover-s.outlook.com','login.microsoftonline.com') {
            [pscustomobject]@{Target=$hostName;HTTPS443=(Test-NetConnection $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue)}
        }
        'RECENT OUTLOOK EVENTS'
        Get-WinEvent -FilterHashtable @{LogName='Application';StartTime=(Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue | Where-Object Message -match 'OUTLOOK|Office|MAPI' | Select-Object -First 50 TimeCreated,Id,ProviderName,LevelDisplayName,Message | Format-List
    }
    Pause-Menu
}

function Start-OutlookSafeMode {
    $path = Get-OfficeAppPath 'OUTLOOK.EXE'
    if ($path) { Start-Process $path -ArgumentList '/safe' } else { Write-Log 'Outlook executable not found.' 'ERROR' }
    Pause-Menu
}

function Test-Identity {
    Show-Header
    Export-Report -Name 'Identity_Diagnostics' -ScriptBlock {
        'MICROSOFT 365 IDENTITY DIAGNOSTICS'
        "Created by: $ScriptAuthor"
        'DEVICE REGISTRATION'
        dsregcmd.exe /status
        'CREDENTIAL MANAGER'
        cmdkey.exe /list
        'OFFICE IDENTITY'
        Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity' -ErrorAction SilentlyContinue | Format-List
        'TIME SERVICE'
        w32tm.exe /query /status
        'WORK OR SCHOOL ACCOUNT'
        Get-ItemProperty 'HKCU:\Software\Microsoft\IdentityCRL\StoredIdentities\*' -ErrorAction SilentlyContinue | Select-Object PSChildName | Format-Table -AutoSize
        'CONNECTIVITY'
        foreach($hostName in 'login.microsoftonline.com','device.login.microsoftonline.com','graph.microsoft.com') {
            [pscustomobject]@{Target=$hostName;HTTPS443=(Test-NetConnection $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue)}
        }
    }
    Pause-Menu
}

function Restart-ClickToRun {
    Show-Header
    try {
        Restart-Service ClickToRunSvc -Force -ErrorAction Stop
        Write-Log 'Office Click-to-Run service restarted.' 'SUCCESS'
    } catch { Write-Log $_.Exception.Message 'ERROR' }
    Pause-Menu
}

function Open-OfficeRepair { Start-Process 'appwiz.cpl'; Pause-Menu }
function Open-InstalledApps { Start-Process 'ms-settings:appsfeatures'; Pause-Menu }
function Open-WorkAccount { Start-Process 'ms-settings:workplace'; Pause-Menu }
function Open-SoundSettings { Start-Process 'ms-settings:sound'; Pause-Menu }
function Open-PrivacyMicrophone { Start-Process 'ms-settings:privacy-microphone'; Pause-Menu }
function Open-PrivacyCamera { Start-Process 'ms-settings:privacy-webcam'; Pause-Menu }

function Run-SFC {
    Show-Header
    if (-not (Test-IsAdministrator)) { Write-Log 'SFC requires Run as administrator.' 'WARN'; Pause-Menu; return }
    sfc.exe /scannow
    Pause-Menu
}

function Run-DISM {
    Show-Header
    if (-not (Test-IsAdministrator)) { Write-Log 'DISM requires Run as administrator.' 'WARN'; Pause-Menu; return }
    DISM.exe /Online /Cleanup-Image /RestoreHealth
    Pause-Menu
}

function Show-OneDriveMenu {
    do {
        Show-Header
        Write-Host 'ONEDRIVE / SHAREPOINT SYNC SCENARIOS' -ForegroundColor Cyan
        Write-Host ' 1. OneDrive not starting'
        Write-Host ' 2. Sign-in loop or wrong account'
        Write-Host ' 3. Red X / stuck syncing / processing changes'
        Write-Host ' 4. Duplicate OneDrive folders'
        Write-Host ' 5. Invalid file names or long paths'
        Write-Host ' 6. Known Folder Backup problem'
        Write-Host ' 7. SharePoint library not syncing'
        Write-Host ' 8. Reset OneDrive'
        Write-Host ' 9. Collect diagnostics'
        Write-Host ' 0. Back'
        switch (Read-Host 'Select') {
            '1' { $exe=Get-OneDriveExe; if($exe){Start-Process $exe}else{Write-Log 'OneDrive not found.' 'ERROR'}; Pause-Menu }
            '2' { Test-Identity }
            '3' { Test-OneDriveSharePoint }
            '4' { Test-OneDriveSharePoint }
            '5' { Test-OneDriveSharePoint }
            '6' { Test-OneDriveSharePoint }
            '7' { Test-OneDriveSharePoint }
            '8' { Reset-OneDrive }
            '9' { Test-OneDriveSharePoint }
            '0' { return }
        }
    } while ($true)
}

function Show-TeamsMenu {
    do {
        Show-Header
        Write-Host 'MICROSOFT TEAMS SCENARIOS' -ForegroundColor Cyan
        Write-Host ' 1. Teams will not open or crashes'
        Write-Host ' 2. Sign-in loop or blank window'
        Write-Host ' 3. Microphone not detected'
        Write-Host ' 4. Camera not detected'
        Write-Host ' 5. No audio or wrong device'
        Write-Host ' 6. Poor call quality'
        Write-Host ' 7. Meetings, presence or notifications issue'
        Write-Host ' 8. Clear Teams cache'
        Write-Host ' 9. Collect diagnostics'
        Write-Host ' 0. Back'
        switch (Read-Host 'Select') {
            '1' { Test-Teams }
            '2' { Test-Identity }
            '3' { Open-PrivacyMicrophone }
            '4' { Open-PrivacyCamera }
            '5' { Open-SoundSettings }
            '6' { Test-M365Connectivity }
            '7' { Test-Teams }
            '8' { Clear-TeamsCache }
            '9' { Test-Teams }
            '0' { return }
        }
    } while ($true)
}

function Show-OutlookMenu {
    do {
        Show-Header
        Write-Host 'MICROSOFT OUTLOOK SCENARIOS' -ForegroundColor Cyan
        Write-Host ' 1. Outlook will not start or crashes'
        Write-Host ' 2. Repeated password prompts'
        Write-Host ' 3. Disconnected / Trying to connect'
        Write-Host ' 4. Mail not sending or receiving'
        Write-Host ' 5. Search not working'
        Write-Host ' 6. Shared mailbox or calendar issue'
        Write-Host ' 7. Add-in problem'
        Write-Host ' 8. OST or profile issue'
        Write-Host ' 9. Start Outlook Safe Mode'
        Write-Host '10. Collect diagnostics'
        Write-Host ' 0. Back'
        switch (Read-Host 'Select') {
            '1' { Start-OutlookSafeMode }
            '2' { Test-Identity }
            '3' { Test-Outlook }
            '4' { Test-Outlook }
            '5' { Start-Process 'ms-settings:search'; Pause-Menu }
            '6' { Test-Outlook }
            '7' { Start-OutlookSafeMode }
            '8' { Start-Process control.exe -ArgumentList 'mlcfg32.cpl'; Pause-Menu }
            '9' { Start-OutlookSafeMode }
            '10' { Test-Outlook }
            '0' { return }
        }
    } while ($true)
}

function Show-SharePointMenu {
    do {
        Show-Header
        Write-Host 'SHAREPOINT SCENARIOS' -ForegroundColor Cyan
        Write-Host ' 1. Site will not open'
        Write-Host ' 2. Access denied or permissions issue'
        Write-Host ' 3. Document library will not sync'
        Write-Host ' 4. File read-only or locked'
        Write-Host ' 5. Upload fails'
        Write-Host ' 6. Office web apps not loading'
        Write-Host ' 7. Authentication loop or wrong account'
        Write-Host ' 8. Collect diagnostics'
        Write-Host ' 0. Back'
        switch (Read-Host 'Select') {
            '1' { Test-M365Connectivity }
            '2' { Test-Identity }
            '3' { Show-OneDriveMenu }
            '4' { Test-OneDriveSharePoint }
            '5' { Test-OneDriveSharePoint }
            '6' { Test-M365Connectivity }
            '7' { Test-Identity }
            '8' { Test-OneDriveSharePoint }
            '0' { return }
        }
    } while ($true)
}

Write-Log "Troubleshooter $ScriptVersion started. Created by $ScriptAuthor."

do {
    Show-Header
    Write-Host 'MAIN MENU' -ForegroundColor Cyan
    Write-Host ' 1. General Microsoft 365 diagnostics'
    Write-Host ' 2. OneDrive and SharePoint sync scenarios'
    Write-Host ' 3. Microsoft Teams scenarios'
    Write-Host ' 4. Microsoft Outlook scenarios'
    Write-Host ' 5. SharePoint scenarios'
    Write-Host ' 6. Microsoft 365 identity diagnostics'
    Write-Host ' 7. Microsoft 365 connectivity test'
    Write-Host ' 8. Close Microsoft 365 app processes'
    Write-Host ' 9. Restart Office Click-to-Run service'
    Write-Host '10. Open Office repair'
    Write-Host '11. Open Installed Apps settings'
    Write-Host '12. Open Access work or school'
    Write-Host '13. Run SFC'
    Write-Host '14. Run DISM RestoreHealth'
    Write-Host ' 0. Exit'
    $choice = Read-Host 'Select an option'
    switch ($choice) {
        '1' { Get-GeneralDiagnostics }
        '2' { Show-OneDriveMenu }
        '3' { Show-TeamsMenu }
        '4' { Show-OutlookMenu }
        '5' { Show-SharePointMenu }
        '6' { Test-Identity }
        '7' { Test-M365Connectivity }
        '8' { Stop-M365Processes }
        '9' { Restart-ClickToRun }
        '10' { Open-OfficeRepair }
        '11' { Open-InstalledApps }
        '12' { Open-WorkAccount }
        '13' { Run-SFC }
        '14' { Run-DISM }
        '0' { Write-Log 'Troubleshooter closed by the user.' }
        default { Write-Host 'Invalid selection.' -ForegroundColor Yellow; Start-Sleep 1 }
    }
} while ($choice -ne '0')
