# Microsoft 365 Apps Advanced Troubleshooter

Created by **Dewald Pretorius**.

A menu-driven PowerShell troubleshooting toolkit for common Microsoft 365 desktop and cloud-connected application issues.

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

## Features

- Self-unblocking BAT launcher
- Self-unblocking PowerShell script
- No automatic administrator elevation
- Timestamped diagnostic reports
- Connectivity testing
- Guided least-disruptive troubleshooting
- SFC and DISM options when manually elevated
- Clear attribution to Dewald Pretorius

## Run

1. Download or clone the repository.
2. Keep the BAT and PS1 files in the same folder.
3. Double-click:

```text
Launch_Microsoft_365_Troubleshooter.bat
```

The first launch of a downloaded BAT file can still display one Windows security warning because Windows evaluates the file before its internal unblock command runs. After approving it once, the launcher removes the downloaded-file marker from itself and the PowerShell script for future launches from that folder.

## Manual PowerShell launch

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\Microsoft_365_Apps_Advanced_Troubleshooter.ps1"
```

## Logs

Reports are written to:

```text
Desktop\Microsoft_365_Troubleshooter_Logs
```

## Safety

The toolkit runs in standard-user mode by default. Repairs that require elevation clearly state that the launcher must be started manually with **Run as administrator**.

Review prompts before clearing caches, resetting OneDrive or running system repairs.
