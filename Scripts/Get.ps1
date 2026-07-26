param (
    [switch]$Verbose,
    [switch]$WhatIf,
    [switch]$Dev,
    [switch]$CLI,
    [switch]$Silent,
    [switch]$Sysprep,
    [string]$LogPath,
    [string]$User,
    [Alias('NoRestartExplorer')]
    [switch]$SkipExplorerRestart,
    [switch]$CreateRestorePoint,
    [switch]$SkipRegistryBackup,
    [switch]$RunDefaults,
    [switch]$RunDefaultsLite,
    [switch]$RunSavedSettings,
    [string]$Config,
    [string]$Apps,
    [string]$AppRemovalTarget,
    [switch]$RemoveApps,
    [switch]$RemoveGamingApps,
    [switch]$RemoveHPApps,
    [switch]$ForceRemoveEdge,
    [switch]$DisableDVR,
    [switch]$DisableGameBarIntegration,
    [switch]$EnableWindowsSandbox,
    [switch]$EnableWindowsSubsystemForLinux,
    [switch]$DisableTelemetry,
    [switch]$DisableSearchHistory,
    [switch]$DisableFastStartup,
    [switch]$DisableBitlockerAutoEncryption,
    [switch]$DisableModernStandbyNetworking,
    [switch]$DisableNotifications,
    [switch]$DisableStorageSense,
    [switch]$DisableUpdateASAP,
    [switch]$PreventUpdateAutoReboot,
    [switch]$DisableDeliveryOptimization,
    [switch]$DisableDeviceAutoAppDownload,
    [switch]$DisableBing,
    [switch]$DisableStoreSearchSuggestions,
    [switch]$DisableDesktopSpotlight,
    [switch]$DisableLockscreenTips,
    [switch]$DisableSuggestions,
    [switch]$DisableLocationServices,
    [switch]$DisableFindMyDevice,
    [switch]$DisableEdgeAds,
    [switch]$DisableBraveBloat,
    [switch]$DisableSettings365Ads,
    [switch]$DisableSettingsHome,
    [switch]$ShowHiddenFolders,
    [switch]$ShowKnownFileExt,
    [switch]$HideDupliDrive,
    [switch]$EnableDarkMode,
    [switch]$DisableTransparency,
    [switch]$DisableAnimations,
    [switch]$TaskbarAlignLeft,
    [switch]$CombineTaskbarAlways, [switch]$CombineTaskbarWhenFull, [switch]$CombineTaskbarNever,
    [switch]$CombineMMTaskbarAlways, [switch]$CombineMMTaskbarWhenFull, [switch]$CombineMMTaskbarNever,
    [switch]$MMTaskbarModeAll, [switch]$MMTaskbarModeMainActive, [switch]$MMTaskbarModeActive,
    [switch]$HideSearchTb, [switch]$ShowSearchIconTb, [switch]$ShowSearchLabelTb, [switch]$ShowSearchBoxTb,
    [switch]$HideTaskview,
    [switch]$DisableStartRecommended,
    [switch]$DisableStartAllApps, [switch]$StartAllAppsCategory, [switch]$StartAllAppsGrid, [switch]$StartAllAppsList,
    [switch]$DisableStartPhoneLink,
    [switch]$DisableCopilot,
    [switch]$DisableRecall,
    [switch]$DisableClickToDo,
    [switch]$DisableAISvcAutoStart,
    [switch]$DisablePaintAI,
    [switch]$DisableNotepadAI,
    [switch]$DisableEdgeAI,
    [switch]$DisableSearchHighlights,
    [switch]$DisableWidgets,
    [switch]$HideChat,
    [switch]$EnableEndTask,
    [switch]$EnableLastActiveClick,
    [switch]$ClearStart,
    [string]$ReplaceStart,
    [switch]$ClearStartAllUsers,
    [string]$ReplaceStartAllUsers,
    [switch]$RevertContextMenu,
    [switch]$DisableDragTray,
    [switch]$DisableMouseAcceleration,
    [switch]$DisableStickyKeys,
    [switch]$DisableWindowSnapping,
    [switch]$DisableSnapAssist,
    [switch]$DisableSnapLayouts,
    [switch]$HideTabsInAltTab, [switch]$Show3TabsInAltTab, [switch]$Show5TabsInAltTab, [switch]$Show20TabsInAltTab,
    [switch]$HideHome,
    [switch]$HideGallery,
    [switch]$ExplorerToHome,
    [switch]$ExplorerToThisPC,
    [switch]$ExplorerToDownloads,
    [switch]$ExplorerToOneDrive,
    [switch]$AddFoldersToThisPC,
    [switch]$HideOnedrive,
    [switch]$Hide3dObjects,
    [switch]$HideMusic,
    [switch]$HideIncludeInLibrary,
    [switch]$HideGiveAccessTo,
    [switch]$HideShare,
    [switch]$ShowDriveLettersFirst,
    [switch]$ShowDriveLettersLast,
    [switch]$ShowNetworkDriveLettersFirst,
    [switch]$HideDriveLetters
)

# Show error if current powershell environment does not have LanguageMode set to FullLanguage 
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
   Write-Host "错误：Win11Debloat 无法在您的系统上运行。PowerShell 执行受到安全策略限制" -ForegroundColor Red
   Write-Output ""
   Write-Output "按 Enter 键退出..."
   Read-Host | Out-Null
   Exit
}

Clear-Host
Write-Output "-------------------------------------------------------------------------------------------"
Write-Output " Win11Debloat 脚本 - Get"
Write-Output "-------------------------------------------------------------------------------------------"

$tempRootPath = $env:TEMP
$tempWorkPath = Join-Path $tempRootPath 'Win11Debloat'
$tempArchivePath = Join-Path $tempRootPath 'win11debloat.zip'

# Download Win11Debloat from GitHub as a zip archive.
try {
    if ($Dev) {
        Write-Output "> 正在下载 Win11Debloat 开发版..."
        $sourceUri = "https://github.com/Raphire/Win11Debloat/archive/refs/heads/master.zip"
    } else {
        Write-Output "> 正在下载 Win11Debloat..."
        $sourceUri = (Invoke-RestMethod https://api.github.com/repos/Raphire/Win11Debloat/releases/latest).zipball_url
    }
    Invoke-RestMethod $sourceUri -OutFile $tempArchivePath
}
catch {
    Write-Host "错误：无法从 GitHub 获取最新版本。请检查您的网络连接后重试。" -ForegroundColor Red
    Write-Output ""
    Write-Output "按 Enter 键退出..."
    Read-Host | Out-Null
    Exit
}

# Remove old script folder if it exists, but keep configs, logs and backups
if (Test-Path $tempWorkPath) {
    Write-Output ""
    Write-Output "> 正在清理旧的 Win11Debloat 文件夹..."

    Get-ChildItem -Path $tempWorkPath -Exclude Config,Logs,Backups | Remove-Item -Recurse -Force
}

$configDir = Join-Path $tempWorkPath 'Config'
$backupDir = Join-Path $tempWorkPath 'ConfigOld'

# Temporarily move existing config files if they exist to prevent them from being overwritten by the new script files, will be moved back after the new script is unpacked
if (Test-Path "$configDir") {
    Write-Output ""
    Write-Output "> 正在备份现有配置文件..."

    New-Item -ItemType Directory -Path "$backupDir" -Force | Out-Null

    $filesToKeep = @(
        'LastUsedSettings.json'
    )

    Get-ChildItem -Path "$configDir" -Recurse | Where-Object { $_.Name -in $filesToKeep } | Move-Item -Destination "$backupDir"

    Remove-Item "$configDir" -Recurse -Force
}

Write-Output ""
Write-Output "> 正在解压..."

# Unzip archive to Win11Debloat folder
Expand-Archive $tempArchivePath $tempWorkPath

# Remove archive
Remove-Item $tempArchivePath

# Move files
Get-ChildItem -Path (Join-Path $tempWorkPath '*Win11Debloat-*') -Recurse | Move-Item -Destination $tempWorkPath

# Add existing config files back to Config folder
if (Test-Path "$backupDir") {
    if (-not (Test-Path "$configDir")) {
        New-Item -ItemType Directory -Path "$configDir" -Force | Out-Null
    }

    Write-Output ""
    Write-Output "> 正在恢复现有配置文件..."

    Get-ChildItem -Path "$backupDir" -Recurse | Move-Item -Destination "$configDir"
    Remove-Item "$backupDir" -Recurse -Force
}

# Make list of arguments to pass on to the script (exclude the -Dev switch, which only affects this launcher)
$arguments = $($PSBoundParameters.GetEnumerator() | Where-Object { $_.Key -ne 'Dev' } | ForEach-Object {
    if ($_.Value -eq $true) {
        "-$($_.Key)"
    } 
    else {
         "-$($_.Key) ""$($_.Value)"""
    }
})

Write-Output ""
Write-Output "> 正在启动 Win11Debloat..."

# Minimize the powershell window when no parameters are provided
if ($arguments.Count -eq 0) {
    $windowStyle = "Minimized"
}
else {
    $windowStyle = "Normal"
}

# Remove Powershell 7 modules from path to prevent module loading issues in the script
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $NewPSModulePath = $env:PSModulePath -split ';' | Where-Object -FilterScript { $_ -like '*WindowsPowerShell*' }
    $env:PSModulePath = $NewPSModulePath -join ';'
}

# Run Win11Debloat script with the provided arguments
$debloatScriptPath = Join-Path $tempWorkPath 'Win11Debloat.ps1'
$debloatProcess = Start-Process powershell.exe -WindowStyle $windowStyle -PassThru -ArgumentList "-executionpolicy bypass -File `"$debloatScriptPath`" $arguments" -Verb RunAs

# Wait for the process to finish before continuing
if ($null -ne $debloatProcess) {
    $debloatProcess.WaitForExit()
}

# Remove all remaining script files, except for configs, logs and backups
if (Test-Path $tempWorkPath) {
    Write-Output ""
    Write-Output "> 正在清理..."

    # Cleanup, remove Win11Debloat directory
    Get-ChildItem -Path $tempWorkPath -Exclude Config,Logs,Backups | Remove-Item -Recurse -Force
}

Write-Output ""
