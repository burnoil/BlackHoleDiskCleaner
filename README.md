# BlackHoleDiskCleaner
Comprehensive Powershell Disk Cleaner

# BlackHoleDiskCleaner.ps1 - Skip Parameters Quick Reference

## Overview
You can selectively disable cleanup operations using skip switches. This is useful when you want to run only specific cleanup tasks or avoid certain operations.

## Available Skip Parameters

### `-SkipTempFiles`
Skips cleaning temporary files and folders in:
- C:\Windows\Temp\*
- C:\Users\*\Documents\*tmp
- C:\Users\*\Appdata\Local\Temp\*
- C:\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*
- C:\Users\*\AppData\Roaming\Microsoft\Windows\Cookies\*

### `-SkipDiskCleanup`
Skips running the Windows Disk Cleanup utility (CleanMgr.exe)

### `-SkipDISM`
Skips running DISM to clean old service pack files

### `-SkipRecycleBin`
Skips emptying the Recycle Bin

## Usage Examples

### Skip Recycle Bin Only
```powershell
.\Clean.ps1 -LocalRun -SkipRecycleBin
```
Runs all cleanup operations except Recycle Bin

### Skip Multiple Operations
```powershell
.\Clean.ps1 -LocalRun -SkipRecycleBin -SkipDISM
```
Runs cleanup but skips both Recycle Bin and DISM

### Skip Recycle Bin in Silent Mode
```powershell
.\Clean.ps1 -LocalRun -Silent -SkipRecycleBin
```
Runs silently and skips Recycle Bin

### Only Clean Temp Files (Skip Everything Else)
```powershell
.\Clean.ps1 -LocalRun -SkipDiskCleanup -SkipDISM -SkipRecycleBin
```
Only cleans temp files, skips all other operations

### Only Run DISM (Skip Everything Else)
```powershell
.\Clean.ps1 -LocalRun -SkipTempFiles -SkipDiskCleanup -SkipRecycleBin
```
Only runs DISM cleanup

## Common Scenarios

### Conservative Cleanup (Leave Recycle Bin Alone)
```powershell
.\Clean.ps1 -LocalRun -SkipRecycleBin
```

### Quick Cleanup (Temp Files Only)
```powershell
.\Clean.ps1 -LocalRun -SkipDiskCleanup -SkipDISM -SkipRecycleBin
```

### Deep Cleanup (Everything But Recycle Bin)
```powershell
.\Clean.ps1 -LocalRun -RepairWMI -SkipRecycleBin
```

### Scheduled Maintenance (Silent, Skip DISM)
```powershell
.\Clean.ps1 -LocalRun -Silent -SkipDISM
```
DISM can be slow, so skip it for quick scheduled tasks

## Combining with Other Parameters

All skip parameters work with any combination of:
- `-LocalRun` or `-ComputerName`
- `-Silent`
- `-RepairWMI`
- `-EnableVerbose`
- `-RecycleBinRetentionDays`
- `-TargetDrive`

### Example: Remote, Silent, Custom Drive, Skip Recycle Bin
```powershell
.\Clean.ps1 -ComputerName "PC01" -Silent -TargetDrive "D:" -SkipRecycleBin
```

## What Gets Skipped

When you use a skip parameter:
- The operation is completely bypassed
- No files are touched for that operation
- A message is logged in the transcript: "Skipping [operation] (Skip[Parameter] specified)"
- In silent mode, no console output but transcript still logs the skip

## Exit Behavior

Skip parameters don't affect exit codes:
- **0**: Success (operations that ran completed successfully)
- **1**: Fatal error (connection or space check failed)

Skipped operations are not considered failures.

## Tips

1. **Test First**: Run with skips first to see what will be cleaned
2. **Recycle Bin**: Most conservative skip - users expect recycle bin to persist
3. **DISM**: Can be slow, consider skipping for quick cleanups
4. **Temp Files**: Usually safe to always clean, rarely skip this
5. **Disk Cleanup**: Safe but can be slow, skip if time is critical

## BigFix/SCCM Deployment Examples

### Conservative Automated Cleanup
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\Clean.ps1' -LocalRun -Silent -SkipRecycleBin"
```

### Quick Temp Cleanup Only
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\Clean.ps1' -LocalRun -Silent -SkipDiskCleanup -SkipDISM -SkipRecycleBin"
```

### Full Cleanup Including WMI
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\Clean.ps1' -LocalRun -Silent -RepairWMI"
```
