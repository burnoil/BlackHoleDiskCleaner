# BlackHoleDiskCleaner.ps1 - Enhanced Features Guide

## New Cleanup Locations

### 1. Windows Update Cache (`-SkipWindowsUpdate` to disable)
**What it cleans:**
- `C:\Windows\SoftwareDistribution\Download\*` - Downloaded update files
- `C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*` - Delivery Optimization cache

**Why it matters:**
- Can recover 1-10 GB depending on pending updates
- Safe to delete - Windows will re-download if needed
- Service is stopped/restarted safely during cleanup

**When to use:**
- Machines with disk space issues
- After major Windows updates
- Regular maintenance

**When to skip:**
- Machines with slow/metered internet connections
- If updates are pending installation
- During update troubleshooting

### 2. Browser Caches (`-SkipBrowserCache` to disable)
**What it cleans:**
- Chrome: Cache, Code Cache, GPU Cache for all user profiles
- Edge: Cache, Code Cache, GPU Cache for all user profiles  
- Firefox: Cache2 folders for all profiles

**Why it matters:**
- Can recover 500 MB - 5 GB per user
- Safe to delete - browsers rebuild caches automatically
- Doesn't delete bookmarks, passwords, or history

**When to use:**
- Users reporting browser slowness
- Disk space recovery
- Before browser troubleshooting

**When to skip:**
- Users are actively browsing (files may be locked)
- If you want to preserve "instant page load" performance temporarily

### 3. System Logs (`-SkipSystemLogs` to disable)
**What it cleans:**
- CBS (Component-Based Servicing) logs older than specified days (default 30)
- Windows Update .etl logs older than specified days
- DISM logs older than specified days
- **Keeps current/active log files**

**Why it matters:**
- Can recover 100 MB - 2 GB on machines with update history
- Safe to delete old logs - only removes aged logs
- Preserves recent logs for troubleshooting

**When to use:**
- Machines with long update history
- Regular maintenance (monthly/quarterly)
- Compliance with log retention policies

**When to skip:**
- During active troubleshooting of Windows updates
- If you need historical log data for analysis
- First 30 days after major Windows upgrade

## New Advanced Features

### 1. Aggressive DISM (`-AggressiveDISM`)
**What it does:**
```powershell
# Standard DISM (default):
dism.exe /online /cleanup-Image /spsuperseded

# Aggressive DISM:
dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase
```

**Differences:**
- **Standard**: Only removes old service pack files (mostly irrelevant on Win10/11)
- **Aggressive**: Removes superseded Windows components, can recover 2-15 GB

**⚠️ IMPORTANT WARNING:**
- `/ResetBase` **prevents uninstalling recent Windows updates**
- Cannot roll back recent feature updates after this runs
- Use only on stable systems past their 30-day update rollback window

**When to use:**
- Machines with limited disk space
- Systems stable for 30+ days after updates
- Before major deployments needing disk space

**When to avoid:**
- First 30 days after major Windows updates
- Troubleshooting unstable systems
- Systems that might need update rollback

**Example:**
```powershell
# Standard cleanup (safe)
.\BlackHoleDiskCleaner.ps1 -LocalRun

# Aggressive cleanup (maximum space recovery)
.\BlackHoleDiskCleaner.ps1 -LocalRun -AggressiveDISM
```

### 2. Dry Run Mode (`-DryRun`)
**What it does:**
- Runs the entire script without actually deleting files
- Shows what would be cleaned
- Logs operations to transcript
- Calculates potential space recovery

**Use cases:**
- Testing before deployment
- Auditing what will be cleaned
- Training/documentation
- Compliance verification

**Example:**
```powershell
# Preview cleanup without deleting
.\BlackHoleDiskCleaner.ps1 -LocalRun -DryRun

# Preview with all operations
.\BlackHoleDiskCleaner.ps1 -LocalRun -DryRun -AggressiveDISM
```

### 3. Log Retention Control (`-LogRetentionDays`)
**What it does:**
- Controls how old system logs must be before deletion
- Default: 30 days
- Range: 1-365 days

**Examples:**
```powershell
# Keep only 7 days of logs (aggressive)
.\BlackHoleDiskCleaner.ps1 -LocalRun -LogRetentionDays 7

# Keep 90 days of logs (conservative)
.\BlackHoleDiskCleaner.ps1 -LocalRun -LogRetentionDays 90

# Keep 1 year of logs (very conservative)
.\BlackHoleDiskCleaner.ps1 -LocalRun -LogRetentionDays 365
```

## Recommended Usage Scenarios

### Scenario 1: Conservative Weekly Maintenance
**Goal**: Safe, routine cleanup without disrupting users
```powershell
.\BlackHoleDiskCleaner.ps1 -LocalRun -Silent -SkipRecycleBin -SkipDISM -SkipBrowserCache
```
**What it does**: Only cleans temp files, Windows Update cache, and old logs

### Scenario 2: Aggressive Space Recovery
**Goal**: Maximum disk space recovery on stable systems
```powershell
.\BlackHoleDiskCleaner.ps1 -LocalRun -AggressiveDISM -RepairWMI
```
**What it does**: All cleanup operations including aggressive DISM

### Scenario 3: User-Friendly Cleanup
**Goal**: Clean system files but leave user-visible items alone
```powershell
.\BlackHoleDiskCleaner.ps1 -LocalRun -SkipRecycleBin -SkipBrowserCache
```
**What it does**: System cleanup without touching Recycle Bin or browser caches

### Scenario 4: Pre-Deployment Test
**Goal**: Verify what will be cleaned before deploying to 1000+ machines
```powershell
.\BlackHoleDiskCleaner.ps1 -LocalRun -DryRun -EnableVerbose
```
**What it does**: Shows everything that would be cleaned with full details

### Scenario 5: Browser Performance Reset
**Goal**: Clear browser caches across all users
```powershell
.\BlackHoleDiskCleaner.ps1 -LocalRun -SkipTempFiles -SkipDiskCleanup -SkipDISM -SkipRecycleBin -SkipWindowsUpdate -SkipSystemLogs
```
**What it does**: Only cleans browser caches

### Scenario 6: Update Troubleshooting Cleanup
**Goal**: Clean update-related files to resolve update issues
```powershell
.\BlackHoleDiskCleaner.ps1 -LocalRun -SkipTempFiles -SkipRecycleBin -SkipBrowserCache -SkipSystemLogs
```
**What it does**: Only Windows Update cache and DISM

## Expected Space Recovery

| Operation | Typical Recovery | Notes |
|-----------|------------------|-------|
| Temp Files | 500 MB - 2 GB | Varies by user activity |
| Windows Update | 1 GB - 10 GB | Depends on pending updates |
| Browser Caches | 500 MB - 5 GB | Per-user, depends on browsing |
| System Logs | 100 MB - 2 GB | Depends on system age |
| Disk Cleanup | 500 MB - 5 GB | Varies by selected items |
| DISM (Standard) | 0 - 500 MB | Legacy, minimal on Win10/11 |
| DISM (Aggressive) | 2 GB - 15 GB | Can be substantial |
| Recycle Bin | 0 - 20 GB | User-dependent |

**Total potential**: 5 GB - 60 GB depending on system state and options

## Safety Considerations

### ✅ Always Safe
- Standard temp files cleanup
- Browser cache cleanup (browsers rebuild)
- Old system logs (older than retention period)
- Recycle Bin (with retention period)

### ⚠️ Use with Caution
- Windows Update cache (requires re-download on slow connections)
- Aggressive DISM (prevents update rollback)

### ❌ Never Touches
- User documents and files
- Program Files or installed applications
- Windows system files actively in use
- Active log files
- Browser bookmarks/passwords/history

## BigFix/SCCM Deployment Examples

### Conservative Automated Cleanup (Recommended)
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\BlackHoleDiskCleaner.ps1' -LocalRun -Silent -SkipRecycleBin -SkipDISM"
```

### Aggressive Disk Space Recovery
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\BlackHoleDiskCleaner.ps1' -LocalRun -Silent -AggressiveDISM -RepairWMI"
```

### Targeted Browser Cache Cleanup
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\BlackHoleDiskCleaner.ps1' -LocalRun -Silent -SkipTempFiles -SkipDiskCleanup -SkipDISM -SkipRecycleBin -SkipWindowsUpdate -SkipSystemLogs"
```

### Full Cleanup with All Options
```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\BlackHoleDiskCleaner.ps1' -LocalRun -Silent -AggressiveDISM -RepairWMI"
```

## Performance Considerations

### Fast Operations (< 5 minutes)
- Temp files cleanup
- Browser cache cleanup
- Windows Update cache

### Medium Operations (5-15 minutes)
- Disk Cleanup utility
- System logs cleanup
- Standard DISM

### Slow Operations (15-60 minutes)
- Aggressive DISM
- WMI repository repair
- Recycle Bin with thousands of files

**Tip**: For scheduled tasks during business hours, consider skipping DISM operations

## Troubleshooting

### "No space recovered"
**Possible causes:**
1. Storage Sense already cleaned recently
2. All skip parameters were used
3. Disk is actually full with legitimate files

**Solutions:**
- Try `-AggressiveDISM` for more thorough cleanup
- Check what's actually using space with WinDirStat or similar
- Review skip parameters

### "Windows Update cache cleanup failed"
**Possible causes:**
1. Windows Update service couldn't be stopped
2. Files locked by pending update installation

**Solutions:**
- Run after Windows Updates complete
- Reboot and try again
- Skip with `-SkipWindowsUpdate`

### "Browser cache cleanup incomplete"
**Possible causes:**
1. Browsers are currently running (files locked)
2. User profiles not accessible

**Solutions:**
- Run when users are logged off
- Close all browser windows
- Skip with `-SkipBrowserCache`

## Best Practices

1. **Test First**: Always run with `-DryRun` on a test machine first
2. **Schedule Wisely**: Run during maintenance windows or off-hours
3. **Log Retention**: Keep at least 30 days of system logs
4. **Update Rollback**: Wait 30 days after major updates before using `-AggressiveDISM`
5. **User Communication**: Notify users if cleaning browser caches or Recycle Bin
6. **Monitor Results**: Review transcript logs to verify expected cleanup
7. **Incremental Deployment**: Start with conservative options, add aggressive cleanup after validation

## Version Compatibility

All features tested and confirmed working on:
- ✅ Windows 10 (1809+)
- ✅ Windows 11 (21H2+)
- ✅ Windows Server 2019
- ✅ Windows Server 2022

Legacy features maintained for:
- ⚠️ Windows 7/8.1 (limited testing, basic features work)
- ⚠️ Windows Server 2016 (basic features work)
