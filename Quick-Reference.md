# Clean.ps1 - Quick Reference Card

## Most Common Commands

### Standard Weekly Maintenance
```powershell
.\Clean.ps1 -LocalRun -Silent
```
Safe, automated, no user disruption

### Conservative Cleanup (Leave User Items Alone)
```powershell
.\Clean.ps1 -LocalRun -SkipRecycleBin -SkipBrowserCache
```
Cleans system files only

### Maximum Space Recovery
```powershell
.\Clean.ps1 -LocalRun -AggressiveDISM -RepairWMI
```
⚠️ Use on stable systems (30+ days since updates)

### Test Before Deploying
```powershell
.\Clean.ps1 -LocalRun -DryRun -EnableVerbose
```
Preview what will be cleaned

### Silent Scheduled Task
```powershell
.\Clean.ps1 -LocalRun -Silent -SkipRecycleBin -SkipDISM
```
Perfect for automation

### Browser Performance Fix
```powershell
.\Clean.ps1 -LocalRun -SkipTempFiles -SkipDiskCleanup -SkipDISM -SkipRecycleBin -SkipWindowsUpdate -SkipSystemLogs
```
Cleans browser caches only

---

## Parameter Quick Reference

| Want to... | Use this parameter |
|------------|-------------------|
| Run silently | `-Silent` |
| Skip Recycle Bin | `-SkipRecycleBin` |
| Skip browser caches | `-SkipBrowserCache` |
| Skip Windows Update cache | `-SkipWindowsUpdate` |
| Skip DISM | `-SkipDISM` |
| Maximum space recovery | `-AggressiveDISM` |
| Test first | `-DryRun` |
| See details | `-EnableVerbose` |
| Repair WMI | `-RepairWMI` |
| Target different drive | `-TargetDrive "D:"` |
| Keep more logs | `-LogRetentionDays 90` |

---

## What Gets Cleaned (By Default)

✅ Temp files (all user profiles)
✅ Windows Update cache
✅ Browser caches (Chrome, Edge, Firefox)
✅ Old system logs (30+ days)
✅ Disk Cleanup utility items
✅ DISM cleanup (standard)
✅ Recycle Bin (3+ days old)

---

## Safety Quick Facts

✅ Safe on Windows 10/11
✅ Never touches user documents
✅ Never deletes program files
✅ Skips locked/in-use files
✅ Full logging maintained
✅ Dry run mode available

⚠️ Use `-AggressiveDISM` only on stable systems
⚠️ Close browsers before browser cache cleanup

---

## BigFix One-Liner (Most Common)

```
waithidden PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\Clean.ps1' -LocalRun -Silent -SkipRecycleBin -SkipDISM"
```

---

## Typical Space Recovery

| Cleanup Type | Space Recovered |
|--------------|-----------------|
| Light (weekly) | 500 MB - 2 GB |
| Standard | 2 GB - 10 GB |
| Aggressive | 5 GB - 30 GB |
| Maximum | 10 GB - 60 GB |

---

## Logs Location

`C:\temp\<ComputerName>-CleanupLogs_<Timestamp>.txt`

Check this file for details, even in Silent mode

---

## Troubleshooting

**No space recovered?**
→ Try `-AggressiveDISM`

**Browsers slow after cleanup?**
→ Normal, caches rebuild on use

**Want to undo?**
→ Can't undo, but `-DryRun` tests first

**Need more details?**
→ Use `-EnableVerbose`

---

## Pro Tips

1. Always test with `-DryRun` first
2. Use `-Silent` for scheduled tasks
3. Skip DISM on fast cleanups
4. Wait 30 days after updates before `-AggressiveDISM`
5. Review transcript logs periodically
6. Combine parameters for custom scenarios

---

## Get Help

```powershell
Get-Help .\Clean.ps1 -Full
```

Or read the detailed guides:
- Silent-Mode-Guide.md
- Skip-Parameters-Guide.md
- Enhanced-Features-Guide.md
- Improvements-Summary.md
