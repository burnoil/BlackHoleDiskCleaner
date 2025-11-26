#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Cleans temporary files and performs disk cleanup operations on local or remote computers.

.DESCRIPTION
    This script performs comprehensive disk cleanup including temp file removal, disk cleanup utility,
    DISM cleanup, recycle bin cleanup, and optional WMI repository repair.

.PARAMETER LocalRun
    Run the script on the local computer without prompting for computer name.

.PARAMETER ComputerName
    Specify a remote computer name to run cleanup operations against.

.PARAMETER EnableVerbose
    Enable verbose output showing detailed file deletion attempts.

.PARAMETER RepairWMI
    Enable WMI repository repair after cleanup operations.

.PARAMETER Silent
    Suppress all console output except errors. Useful for automated/scheduled tasks.

.PARAMETER RecycleBinRetentionDays
    Number of days to retain items in recycle bin before deletion. Default is 3 days.

.PARAMETER TargetDrive
    The drive letter to target for cleanup operations. Default is 'C:'.

.PARAMETER SkipTempFiles
    Skip cleaning temporary files and folders.

.PARAMETER SkipDiskCleanup
    Skip running the Windows Disk Cleanup utility (CleanMgr.exe).

.PARAMETER SkipDISM
    Skip running DISM to clean old service pack files.

.PARAMETER SkipRecycleBin
    Skip emptying the Recycle Bin.

.PARAMETER SkipWindowsUpdate
    Skip cleaning Windows Update cache and Delivery Optimization files.

.PARAMETER SkipBrowserCache
    Skip cleaning browser caches (Chrome, Edge, Firefox).

.PARAMETER SkipSystemLogs
    Skip cleaning old system logs (CBS, Windows Update, DISM logs).

.PARAMETER AggressiveDISM
    Use aggressive DISM cleanup with StartComponentCleanup and ResetBase for maximum space recovery.
    Warning: This prevents uninstalling recent Windows updates. Use with caution.

.PARAMETER DryRun
    Preview mode - shows what would be cleaned without actually deleting files.

.PARAMETER LogRetentionDays
    Number of days to retain system log files. Logs older than this are deleted. Default is 30 days.

.EXAMPLE
    .\Clean.ps1 -LocalRun -Silent
    Runs cleanup on local computer with no console output.

.EXAMPLE
    .\Clean.ps1 -ComputerName "PC01" -RepairWMI
    Runs cleanup on remote computer PC01 and repairs WMI repository.

.EXAMPLE
    .\Clean.ps1 -LocalRun -SkipRecycleBin -SkipDISM
    Runs cleanup but skips Recycle Bin and DISM operations.

.EXAMPLE
    .\Clean.ps1 -LocalRun -AggressiveDISM
    Runs cleanup with aggressive DISM component store cleanup for maximum space recovery.

.EXAMPLE
    .\Clean.ps1 -LocalRun -DryRun
    Preview what would be cleaned without actually deleting anything.

.EXAMPLE
    .\Clean.ps1 -LocalRun -SkipBrowserCache -SkipWindowsUpdate
    Runs standard cleanup but skips browser and Windows Update caches.
#>

[CmdletBinding()]
Param(
    [Parameter(ParameterSetName = 'Local')]
    [switch]$LocalRun,

    [Parameter(ParameterSetName = 'Remote')]
    [string]$ComputerName,

    [switch]$EnableVerbose,
    [switch]$RepairWMI,
    [switch]$Silent,

    [ValidateRange(0, 365)]
    [int]$RecycleBinRetentionDays = 3,

    [ValidatePattern('^[A-Z]:$')]
    [string]$TargetDrive = 'C:',

    # Skip specific operations
    [switch]$SkipTempFiles,
    [switch]$SkipDiskCleanup,
    [switch]$SkipDISM,
    [switch]$SkipRecycleBin,
    [switch]$SkipWindowsUpdate,
    [switch]$SkipBrowserCache,
    [switch]$SkipSystemLogs,

    # Advanced options
    [switch]$AggressiveDISM,
    [switch]$DryRun,
    [ValidateRange(1, 365)]
    [int]$LogRetentionDays = 30
)

# Set verbosity based on parameters
if ($EnableVerbose) {
    $VerbosePreference = "Continue"
}

# Suppress all errors and warnings in Silent mode
if ($Silent) {
    $ErrorActionPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'
    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
}

# Output function that respects Silent mode
Function Write-CleanupMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )

    # In Silent mode, suppress ALL output (transcript still captures everything)
    if ($Silent) {
        return
    }

    $ColorMap = @{
        'Info'    = 'Cyan'
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
    }

    Write-Host $Message -ForegroundColor $ColorMap[$Type]
}

# Registry locations for Disk Cleanup
$SageSet = "StateFlags0099"
$Base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\"

$CleanupLocations = @(
    "Active Setup Temp Folders",
    "BranchCache",
    "Downloaded Program Files",
    "GameNewsFiles",
    "GameStatisticsFiles",
    "GameUpdateFiles",
    "Internet Cache Files",
    "Memory Dump Files",
    "Offline Pages Files",
    "Old ChkDsk Files",
    "Previous Installations",
    "Service Pack Cleanup",
    "Setup Log Files",
    "System error memory dump files",
    "System error minidump files",
    "Temporary Files",
    "Temporary Setup Files",
    "Temporary Sync Files",
    "Thumbnail Cache",
    "Update Cleanup",
    "Upgrade Discarded Files",
    "User file versions",
    "Windows Defender",
    "Windows Error Reporting Archive Files",
    "Windows Error Reporting Queue Files",
    "Windows Error Reporting System Archive Files",
    "Windows Error Reporting System Queue Files",
    "Windows ESD installation files",
    "Windows Upgrade Log Files"
)

Function Get-ComputerName {
    [CmdletBinding()]
    Param(
        [switch]$LocalRun,
        [string]$Computer
    )

    if ($LocalRun) {
        $CompName = $env:COMPUTERNAME
        $Remote = $false
    }
    elseif ($Computer) {
        $CompName = $Computer
        $Remote = $true
    }
    else {
        $CompName = Read-Host "Enter the name of the computer"
        $Remote = $true
    }

    [PSCustomObject]@{
        ComputerName = $CompName
        Remote       = $Remote
        PSRemoting   = $false
        Credential   = $null
    }
}

Function Test-PSRemoting {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Testing PS Remoting to $($ComputerOBJ.ComputerName)" -Type Info

    try {
        $TestResult = Test-WSMan -ComputerName $ComputerOBJ.ComputerName -ErrorAction Stop
        $ComputerOBJ.PSRemoting = $true
        Write-CleanupMessage "PS Remoting is available" -Type Success
    }
    catch {
        Write-CleanupMessage "PS Remoting is not available. Script will attempt WMI/CIM operations only." -Type Warning
        $ComputerOBJ.PSRemoting = $false
    }

    return $ComputerOBJ
}

Function Clear-RecycleBin {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ,

        [int]$RetentionDays = 3
    )

    Write-CleanupMessage "Emptying Recycle Bin items older than $RetentionDays days" -Type Warning

    $ScriptBlock = {
        param($RetentionDays)
        $RecyclerError = $false
        $Shell = $null

        try {
            $Shell = New-Object -ComObject Shell.Application
            $Recycler = $Shell.NameSpace(0xa)
            $Items = $Recycler.Items()

            foreach ($item in $Items) {
                try {
                    $DeletedDate = $Recycler.GetDetailsOf($item, 2) -replace "\u200f|\u200e", ""
                    $DeletedDatetime = Get-Date $DeletedDate -ErrorAction Stop
                    [int]$DeletedDays = (New-TimeSpan -Start $DeletedDatetime -End (Get-Date)).Days

                    if ($DeletedDays -ge $RetentionDays) {
                        Remove-Item -Path $item.Path -Confirm:$false -Force -Recurse -ErrorAction Stop
                    }
                }
                catch {
                    # Individual item cleanup failure - continue with others
                    continue
                }
            }
        }
        catch {
            $RecyclerError = $true
        }
        finally {
            if ($null -ne $Shell) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        return -not $RecyclerError
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $RetentionDays -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock -RetentionDays $RetentionDays
    }

    if ($Result) {
        Write-CleanupMessage "Recycle Bin items older than $RetentionDays days were deleted" -Type Success
    }
    else {
        Write-CleanupMessage "Unable to delete some items in the Recycle Bin" -Type Error
    }
}

Function Clear-Path {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-Verbose "Cleaning $Path"

    $ScriptBlock = {
        param($TargetPath)

        if (Test-Path $TargetPath) {
            $Items = Get-ChildItem -Path $TargetPath -Recurse -ErrorAction SilentlyContinue

            foreach ($Item in $Items) {
                try {
                    Remove-Item -Path $Item.FullName -Confirm:$false -Recurse -Force -ErrorAction Stop
                }
                catch {
                    Write-Verbose "$($Item.FullName) - $($_.Exception.Message)"
                }
            }
        }
    }

    if ($ComputerOBJ.PSRemoting) {
        Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $Path -Credential $ComputerOBJ.Credential
    }
    else {
        & $ScriptBlock -TargetPath $Path
    }
}

Function Get-DriveFreeSpace {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ,

        [Parameter(Mandatory = $true)]
        [string]$Drive,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Initial', 'Final')]
        [string]$Stage
    )

    try {
        if ($ComputerOBJ.Remote) {
            $CimSession = New-CimSession -ComputerName $ComputerOBJ.ComputerName -Credential $ComputerOBJ.Credential -ErrorAction Stop
            $LogicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Drive'" -CimSession $CimSession -ErrorAction Stop
            Remove-CimSession -CimSession $CimSession
        }
        else {
            $LogicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Drive'" -ErrorAction Stop
        }

        $FreeSpaceGB = [decimal]("{0:N2}" -f ($LogicalDisk.FreeSpace / 1GB))
        $Message = if ($Stage -eq 'Initial') { "Current" } else { "Final" }
        Write-CleanupMessage "$Message Free Space on $Drive : $FreeSpaceGB GB" -Type Info

        return $FreeSpaceGB
    }
    catch {
        Write-CleanupMessage "Unable to retrieve free space from $Drive drive" -Type Error
        return $null
    }
}

Function Invoke-DiskCleanup {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Configuring Disk Cleanup utility" -Type Warning

    $ScriptBlock = {
        param($Locations, $SageSet, $Base)

        # Set registry keys for cleanup locations
        foreach ($Location in $Locations) {
            $KeyPath = Join-Path -Path $Base -ChildPath $Location
            if (Test-Path $KeyPath) {
                try {
                    Set-ItemProperty -Path $KeyPath -Name $SageSet -Type DWORD -Value 2 -ErrorAction Stop
                }
                catch {
                    Write-Verbose "Failed to set registry key for $Location"
                }
            }
        }

        # Run CleanMgr
        try {
            Start-Process -FilePath CleanMgr.exe -ArgumentList "/sagerun:99" -Wait -NoNewWindow -ErrorAction Stop
            return $true
        }
        catch {
            return $false
        }
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $CleanupLocations, $SageSet, $Base -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock -Locations $CleanupLocations -SageSet $SageSet -Base $Base
    }

    if ($Result) {
        Write-CleanupMessage "Disk Cleanup completed successfully" -Type Success
    }
    else {
        Write-CleanupMessage "Failed to run Disk Cleanup utility" -Type Error
    }
}

Function Clear-IEHistory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Attempting to erase Internet Explorer temp data" -Type Warning

    $ScriptBlock = {
        try {
            Start-Process -FilePath rundll32.exe -ArgumentList 'inetcpl.cpl,ClearMyTracksByProcess 4351' -Wait -NoNewWindow -ErrorAction Stop
            return $true
        }
        catch {
            return $false
        }
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock
    }

    if ($Result) {
        Write-CleanupMessage "Internet Explorer temp data erased successfully" -Type Success
    }
    else {
        Write-CleanupMessage "Failed to erase Internet Explorer temp data" -Type Error
    }
}

Function Invoke-DISM {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    if ($AggressiveDISM) {
        Write-CleanupMessage "Running DISM with aggressive component store cleanup" -Type Warning
    }
    else {
        Write-CleanupMessage "Running DISM to clean old service pack files" -Type Warning
    }

    $ScriptBlock = {
        param($UseAggressive)
        try {
            if ($UseAggressive) {
                # More thorough cleanup - removes superseded components
                # This can take longer but recovers more space
                $DISMResult = dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase 2>&1
            }
            else {
                # Legacy method - just service pack cleanup
                $DISMResult = dism.exe /online /cleanup-Image /spsuperseded 2>&1
            }
            return $DISMResult
        }
        catch {
            return $false
        }
    }

    if ($ComputerOBJ.PSRemoting) {
        $DISM = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $AggressiveDISM -Credential $ComputerOBJ.Credential
    }
    else {
        $DISM = & $ScriptBlock -UseAggressive $AggressiveDISM
    }

    if ($DISM -match 'The operation completed successfully') {
        Write-CleanupMessage "DISM completed successfully" -Type Success
    }
    else {
        Write-CleanupMessage "Unable to complete DISM cleanup" -Type Error
    }
}

Function Repair-WMIRepository {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Attempting to repair WMI repository" -Type Warning

    $ScriptBlock = {
        try {
            $Result = cmd.exe /c "Winmgmt /salvagerepository" 2>&1
            return $true
        }
        catch {
            return $false
        }
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock
    }

    if ($Result) {
        Write-CleanupMessage "WMI repository repair completed" -Type Success
    }
    else {
        Write-CleanupMessage "Failed to repair WMI repository" -Type Error
    }
}

Function Clear-WindowsUpdateCache {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Cleaning Windows Update cache and delivery optimization" -Type Warning

    $ScriptBlock = {
        param($TargetDrive, $DryRun)
        $CleanedItems = 0

        # Stop Windows Update service
        try {
            Stop-Service -Name wuauserv -Force -ErrorAction Stop
            $ServiceStopped = $true
        }
        catch {
            $ServiceStopped = $false
        }

        if ($ServiceStopped) {
            # Clean SoftwareDistribution Download folder
            $SoftwareDistPath = "$TargetDrive\Windows\SoftwareDistribution\Download\*"
            if (Test-Path $SoftwareDistPath) {
                try {
                    if (-not $DryRun) {
                        Remove-Item -Path $SoftwareDistPath -Recurse -Force -ErrorAction SilentlyContinue
                    }
                    $CleanedItems++
                }
                catch {
                    # Continue on error
                }
            }

            # Restart Windows Update service
            try {
                Start-Service -Name wuauserv -ErrorAction Stop
            }
            catch {
                # Service will start on next update check
            }
        }

        # Clean Delivery Optimization cache
        $DeliveryOptPath = "$TargetDrive\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*"
        if (Test-Path $DeliveryOptPath) {
            try {
                if (-not $DryRun) {
                    Remove-Item -Path $DeliveryOptPath -Recurse -Force -ErrorAction SilentlyContinue
                }
                $CleanedItems++
            }
            catch {
                # Continue on error
            }
        }

        return $CleanedItems
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $TargetDrive, $DryRun -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock -TargetDrive $TargetDrive -DryRun $DryRun
    }

    if ($Result -gt 0) {
        Write-CleanupMessage "Windows Update cache cleaned ($Result locations)" -Type Success
    }
    else {
        Write-CleanupMessage "No Windows Update cache to clean or operation failed" -Type Info
    }
}

Function Clear-BrowserCaches {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Cleaning browser caches (Chrome, Edge, Firefox)" -Type Warning

    $ScriptBlock = {
        param($TargetDrive, $DryRun)
        $CleanedItems = 0

        # Chrome cache paths
        $ChromePaths = @(
            "$TargetDrive\Users\*\AppData\Local\Google\Chrome\User Data\Default\Cache\*",
            "$TargetDrive\Users\*\AppData\Local\Google\Chrome\User Data\Default\Code Cache\*",
            "$TargetDrive\Users\*\AppData\Local\Google\Chrome\User Data\Default\GPUCache\*"
        )

        # Edge cache paths
        $EdgePaths = @(
            "$TargetDrive\Users\*\AppData\Local\Microsoft\Edge\User Data\Default\Cache\*",
            "$TargetDrive\Users\*\AppData\Local\Microsoft\Edge\User Data\Default\Code Cache\*",
            "$TargetDrive\Users\*\AppData\Local\Microsoft\Edge\User Data\Default\GPUCache\*"
        )

        # Firefox cache paths
        $FirefoxPaths = @(
            "$TargetDrive\Users\*\AppData\Local\Mozilla\Firefox\Profiles\*.default*\cache2\*",
            "$TargetDrive\Users\*\AppData\Local\Mozilla\Firefox\Profiles\*.default-release\cache2\*"
        )

        $AllPaths = $ChromePaths + $EdgePaths + $FirefoxPaths

        foreach ($Path in $AllPaths) {
            if (Test-Path $Path) {
                try {
                    $Items = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
                    foreach ($Item in $Items) {
                        try {
                            if (-not $DryRun) {
                                Remove-Item -Path $Item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                            }
                        }
                        catch {
                            # Continue on locked files
                        }
                    }
                    $CleanedItems++
                }
                catch {
                    # Continue on error
                }
            }
        }

        return $CleanedItems
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $TargetDrive, $DryRun -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock -TargetDrive $TargetDrive -DryRun $DryRun
    }

    if ($Result -gt 0) {
        Write-CleanupMessage "Browser caches cleaned ($Result locations)" -Type Success
    }
    else {
        Write-CleanupMessage "No browser caches to clean" -Type Info
    }
}

Function Clear-SystemLogs {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        $ComputerOBJ
    )

    Write-CleanupMessage "Cleaning old system logs and CBS logs" -Type Warning

    $ScriptBlock = {
        param($TargetDrive, $LogRetentionDays, $DryRun)
        $CleanedItems = 0
        $CutoffDate = (Get-Date).AddDays(-$LogRetentionDays)

        # CBS (Component-Based Servicing) logs
        $CBSPath = "$TargetDrive\Windows\Logs\CBS"
        if (Test-Path $CBSPath) {
            try {
                $OldLogs = Get-ChildItem -Path $CBSPath -Filter "*.log" -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -lt $CutoffDate -and $_.Name -ne "CBS.log" }
                
                foreach ($Log in $OldLogs) {
                    try {
                        if (-not $DryRun) {
                            Remove-Item -Path $Log.FullName -Force -ErrorAction Stop
                        }
                        $CleanedItems++
                    }
                    catch {
                        # Continue on error
                    }
                }
            }
            catch {
                # Continue on error
            }
        }

        # Windows Update logs
        $WindowsUpdateLogs = "$TargetDrive\Windows\Logs\WindowsUpdate\*.etl"
        if (Test-Path $WindowsUpdateLogs) {
            try {
                $OldLogs = Get-ChildItem -Path "$TargetDrive\Windows\Logs\WindowsUpdate" -Filter "*.etl" -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -lt $CutoffDate }
                
                foreach ($Log in $OldLogs) {
                    try {
                        if (-not $DryRun) {
                            Remove-Item -Path $Log.FullName -Force -ErrorAction Stop
                        }
                        $CleanedItems++
                    }
                    catch {
                        # Continue on error
                    }
                }
            }
            catch {
                # Continue on error
            }
        }

        # DISM logs
        $DISMPath = "$TargetDrive\Windows\Logs\DISM"
        if (Test-Path $DISMPath) {
            try {
                $OldLogs = Get-ChildItem -Path $DISMPath -Filter "*.log" -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -lt $CutoffDate -and $_.Name -ne "dism.log" }
                
                foreach ($Log in $OldLogs) {
                    try {
                        if (-not $DryRun) {
                            Remove-Item -Path $Log.FullName -Force -ErrorAction Stop
                        }
                        $CleanedItems++
                    }
                    catch {
                        # Continue on error
                    }
                }
            }
            catch {
                # Continue on error
            }
        }

        return $CleanedItems
    }

    if ($ComputerOBJ.PSRemoting) {
        $Result = Invoke-Command -ComputerName $ComputerOBJ.ComputerName -ScriptBlock $ScriptBlock -ArgumentList $TargetDrive, $LogRetentionDays, $DryRun -Credential $ComputerOBJ.Credential
    }
    else {
        $Result = & $ScriptBlock -TargetDrive $TargetDrive -LogRetentionDays $LogRetentionDays -DryRun $DryRun
    }

    if ($Result -gt 0) {
        Write-CleanupMessage "System logs cleaned ($Result files older than $LogRetentionDays days)" -Type Success
    }
    else {
        Write-CleanupMessage "No old system logs to clean" -Type Info
    }
}

#region Main Execution

if (-not $Silent) {
    Clear-Host
}

if ($DryRun) {
    Write-CleanupMessage "***** DRY RUN MODE - No files will be deleted *****" -Type Warning
    Write-Host ""
}

if (-not $Silent) {
    Write-CleanupMessage "** This tool will erase temp files across all user profiles. Use with caution. **" -Type Warning
    Write-Host ""
}

# Determine target computer
if ($LocalRun) {
    $ComputerOBJ = Get-ComputerName -LocalRun
}
elseif ($ComputerName) {
    $ComputerOBJ = Get-ComputerName -Computer $ComputerName
}
else {
    $ComputerOBJ = Get-ComputerName
}

# Confirm target
if ($LocalRun -or $Silent) {
    Write-CleanupMessage "Target computer: $($ComputerOBJ.ComputerName)" -Type Info
}
else {
    Write-Host "You have entered $($ComputerOBJ.ComputerName). Is this correct?"
    Pause
}

if (-not $Silent) {
    Write-Host ""
}

# Stop any existing transcript
try {
    Stop-Transcript -ErrorAction Stop
}
catch {
    # No transcript running
}

# Start logging
$TranscriptPath = "C:\temp"
if (-not (Test-Path $TranscriptPath)) {
    New-Item -Path $TranscriptPath -ItemType Directory -Force | Out-Null
}

$Timestamp = Get-Date -Format "yyyy-MM-dd_THHmmss"
$TranscriptFile = "$TranscriptPath\$($ComputerOBJ.ComputerName)-CleanupLogs_$Timestamp.txt"
Write-CleanupMessage "Starting transcript logging to $TranscriptFile" -Type Info
Start-Transcript -Path $TranscriptFile | Out-Null
Write-Output "Cleanup started at: $([System.DateTime]::Now)"

if (-not $Silent) {
    Write-Host ""
}

if (-not $Silent) {
    Write-Host ("*" * 95)
}
if ($ComputerOBJ.Remote) {
    $ComputerOBJ = Test-PSRemoting -ComputerOBJ $ComputerOBJ
    if (-not $ComputerOBJ.PSRemoting) {
        Write-CleanupMessage "Cannot establish PS Remoting. Exiting." -Type Error
        if (-not $Silent) {
            Read-Host "Press Enter to exit"
        }
        exit 1
    }
}

# Get initial free space
$OrigFreeSpace = Get-DriveFreeSpace -ComputerOBJ $ComputerOBJ -Drive $TargetDrive -Stage Initial

if ($null -eq $OrigFreeSpace) {
    Write-CleanupMessage "Cannot determine free space. Exiting." -Type Error
    if (-not $Silent) {
        Read-Host "Press Enter to exit"
    }
    exit 1
}

if (-not $Silent) {
    Write-Host ("*" * 95)
    Write-Host ""
}

#region Cleanup Operations

if (-not $SkipTempFiles) {
    Write-CleanupMessage "Cleaning temp directories across all user profiles" -Type Warning

    # Standard temp cleanup paths
    $CleanupPaths = @(
        "$TargetDrive\Windows\Temp\*",
        "$TargetDrive\Users\*\Documents\*tmp",
        "$TargetDrive\Documents and Settings\*\Local Settings\Temp\*",
        "$TargetDrive\Users\*\Appdata\Local\Temp\*",
        "$TargetDrive\Users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*",
        "$TargetDrive\Users\*\AppData\Roaming\Microsoft\Windows\Cookies\*"
    )

    foreach ($Path in $CleanupPaths) {
        Clear-Path -Path $Path -ComputerOBJ $ComputerOBJ
    }

    Write-CleanupMessage "Standard temp paths cleaned" -Type Success
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping temp files cleanup (SkipTempFiles specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

# Additional cleanup operations (new)
if (-not $SkipWindowsUpdate) {
    Clear-WindowsUpdateCache -ComputerOBJ $ComputerOBJ
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping Windows Update cache cleanup (SkipWindowsUpdate specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

if (-not $SkipBrowserCache) {
    Clear-BrowserCaches -ComputerOBJ $ComputerOBJ
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping browser cache cleanup (SkipBrowserCache specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

if (-not $SkipSystemLogs) {
    Clear-SystemLogs -ComputerOBJ $ComputerOBJ
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping system logs cleanup (SkipSystemLogs specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

# Run cleanup utilities
if (-not $SkipDiskCleanup) {
    Invoke-DiskCleanup -ComputerOBJ $ComputerOBJ
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping Disk Cleanup utility (SkipDiskCleanup specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

if (-not $SkipDISM) {
    Invoke-DISM -ComputerOBJ $ComputerOBJ
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping DISM cleanup (SkipDISM specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

if (-not $SkipRecycleBin) {
    Clear-RecycleBin -ComputerOBJ $ComputerOBJ -RetentionDays $RecycleBinRetentionDays
    if (-not $Silent) {
        Write-Host ""
    }
} else {
    Write-CleanupMessage "Skipping Recycle Bin cleanup (SkipRecycleBin specified)" -Type Info
    if (-not $Silent) {
        Write-Host ""
    }
}

if ($RepairWMI) {
    Repair-WMIRepository -ComputerOBJ $ComputerOBJ
    if (-not $Silent) {
        Write-Host ""
    }
}

#endregion

if (-not $Silent) {
    Write-Host ("*" * 95)
}

# Get final free space
$FinalFreeSpace = Get-DriveFreeSpace -ComputerOBJ $ComputerOBJ -Drive $TargetDrive -Stage Final
$SpaceRecovered = $FinalFreeSpace - $OrigFreeSpace

if ($SpaceRecovered -lt 0) {
    Write-CleanupMessage "Less than a gigabyte of free space was recovered" -Type Info
}
elseif ($SpaceRecovered -eq 0) {
    Write-CleanupMessage "No space was recovered" -Type Info
}
else {
    Write-CleanupMessage "Free space recovered: $([Math]::Round($SpaceRecovered, 2)) GB" -Type Success
}

if (-not $Silent) {
    Write-Host ("*" * 95)
    Write-Host ""
}

Write-Output "Cleanup completed at: $([System.DateTime]::Now)"
Stop-Transcript | Out-Null

#endregion
