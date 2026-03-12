# Restart the Windows Explorer process
function RestartExplorer {
    Write-Host "> 正在尝试重启 Windows 资源管理器进程以应用所有更改..."
    
    if ($script:Params.ContainsKey("Sysprep") -or $script:Params.ContainsKey("User") -or $script:Params.ContainsKey("NoRestartExplorer")) {
        Write-Host "已跳过资源管理器进程重启，请手动重启电脑以应用所有更改" -ForegroundColor Yellow
        return
    }

    foreach ($paramKey in $script:Params.Keys) {
        if ($script:Features.ContainsKey($paramKey) -and $script:Features[$paramKey].RequiresReboot -eq $true) {
            $feature = $script:Features[$paramKey]
            Write-Host "警告：'$($feature.Action) $($feature.Label)' 需要重启才能完全生效" -ForegroundColor Yellow
        }
    }

    # Only restart if the powershell process matches the OS architecture.
    # Restarting explorer from a 32bit PowerShell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem) {
        Write-Host "正在重启 Windows 资源管理器进程...（屏幕可能会闪烁）"
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Host "无法重启 Windows 资源管理器进程，请手动重启电脑以应用所有更改" -ForegroundColor Yellow
    }
}