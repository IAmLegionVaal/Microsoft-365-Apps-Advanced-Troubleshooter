# Microsoft 365 Apps Advanced Troubleshooter

Created by **Dewald Pretorius**.

A menu-driven PowerShell troubleshooting toolkit for common Microsoft 365 desktop and cloud-connected application issues, with guarded repair workflows.

## Coverage

- Microsoft Teams
- OneDrive
- SharePoint
- Outlook
- Microsoft 365 identity and sign-in
- Office Click-to-Run
- Microsoft Store apps
- Windows system health
- Microsoft 365 connectivity

## Included scenarios

### OneDrive and SharePoint sync

- OneDrive not starting
- Sign-in loops and wrong-account issues
- Red X, stuck syncing and processing changes
- Duplicate OneDrive folders
- Invalid file names and long paths
- Known Folder Backup problems
- SharePoint library sync failures
- OneDrive reset

### Microsoft Teams

- Teams will not open or crashes
- Sign-in loop or blank authentication window
- Microphone not detected
- Camera not detected
- No audio or wrong device
- Poor call quality
- Meeting, presence and notification issues
- Teams cache reset

### Outlook

- Outlook will not start or crashes
- Repeated password prompts
- Disconnected or trying to connect
- Mail not sending or receiving
- Search failures
- Shared mailbox and calendar issues
- Add-in problems
- OST and profile problems
- Outlook Safe Mode

### SharePoint

- Site will not open
- Access denied and permissions issues
- Document library will not sync
- Locked or read-only files
- Upload failures
- Office web apps not loading
- Authentication loops and wrong-account issues

## Main troubleshooting toolkit

Run:

```text
Launch_Microsoft_365_Troubleshooter.bat
```

Or launch manually:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\Microsoft_365_Apps_Advanced_Troubleshooter.ps1"
```

## Standalone SharePoint library sync repair

`SharePoint_Sync_Repair_Toolkit.ps1` performs actual client-side repairs for SharePoint document libraries. SharePoint libraries use the OneDrive sync engine on Windows, so the tool repairs both the sync engine and its supporting Office and Windows components.

Repair options include:

- Restart the OneDrive sync engine
- Reset the OneDrive sync engine using the Microsoft-supported `/reset` operation
- Rebuild the local Office document cache by moving it to a timestamped backup
- Clear temporary Microsoft 365 web cache files
- Start or restart the Windows WebClient service
- Restart Windows Explorer to refresh sync status icons and shell integration
- Flush the DNS resolver cache
- Run a combined safe-repair workflow

Run the repair menu:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SharePoint_Sync_Repair_Toolkit.ps1
```

Preview repairs without changing the workstation:

```powershell
.\SharePoint_Sync_Repair_Toolkit.ps1 -RepairAllSafe -DryRun
```

Run selected repairs:

```powershell
.\SharePoint_Sync_Repair_Toolkit.ps1 -RestartOneDrive -FlushDns -RefreshExplorer
.\SharePoint_Sync_Repair_Toolkit.ps1 -ResetOneDrive
.\SharePoint_Sync_Repair_Toolkit.ps1 -RebuildOfficeFileCache
```

A double-click launcher is included:

```text
Launch_SharePoint_Sync_Repair_Toolkit.bat
```

## Features

- Self-unblocking BAT launchers
- Self-unblocking PowerShell scripts
- No automatic administrator elevation
- Timestamped diagnostic reports
- Before-and-after JSON repair snapshots
- Recoverable timestamped cache backups
- Connectivity testing
- Guided least-disruptive troubleshooting
- Explicit confirmation before repairs
- `-DryRun` support for the standalone SharePoint sync repair tool
- SFC and DISM options when manually elevated
- Clear attribution to Dewald Pretorius

## Logs

The main toolkit writes reports to:

```text
Desktop\Microsoft_365_Troubleshooter_Logs
```

The standalone SharePoint sync repair tool writes logs and backups to:

```text
Desktop\SharePoint_Sync_Repair_Logs
```

## Safety

- The main toolkit runs in standard-user mode by default.
- Repairs that require elevation state that the launcher must be started manually with **Run as administrator**.
- SharePoint and Office cache repairs move local cache folders into timestamped backups instead of deleting them.
- The SharePoint repair tool does not delete cloud content or remove Microsoft 365 accounts.
- Microsoft states that resetting OneDrive disconnects all sync connections and starts a full synchronisation. Folder selections may need to be configured again after the reset.
- OneDrive reset therefore requires the technician to type `REPAIR`.
- Office document cache rebuild requires all Office files to be saved and synchronised first.

## Author

Dewald Pretorius — L2 IT Support Engineer
