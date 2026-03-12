function CreateSystemRestorePoint {
    $SysRestore = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval"
    $failed = $false

    if ($SysRestore.RPSessionInterval -eq 0) {
        # In GUI mode, skip the prompt and just try to enable it
        if ($script:GuiWindow -or $Silent -or $( Read-Host -Prompt "系统还原已禁用，是否启用并创建还原点？(y/n)") -eq 'y') {
            $enableSystemRestoreJob = Start-Job {
                try {
                    Enable-ComputerRestore -Drive "$env:SystemDrive"
                }
                catch {
                    return "错误：无法启用系统还原：$_"
                }
                return $null
            }

            $enableSystemRestoreJobDone = $enableSystemRestoreJob | Wait-Job -TimeOut 20

            if (-not $enableSystemRestoreJobDone) {
                Remove-Job -Job $enableSystemRestoreJob -Force -ErrorAction SilentlyContinue
                Write-Host "错误：无法启用系统还原并创建还原点，操作超时" -ForegroundColor Red
                $failed = $true
            }
            else {
                $result = Receive-Job $enableSystemRestoreJob
                Remove-Job -Job $enableSystemRestoreJob -ErrorAction SilentlyContinue
                if ($result) {
                    Write-Host $result -ForegroundColor Red
                    $failed = $true
                }
            }
        }
        else {
            Write-Host ""
            $failed = $true
        }
    }

    if (-not $failed) {
        $createRestorePointJob = Start-Job {
            # Find existing restore points that are less than 24 hours old
            try {
                $recentRestorePoints = Get-ComputerRestorePoint | Where-Object { (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime) -le (New-TimeSpan -Hours 24) }
            }
            catch {
                return @{ Success = $false; Message = "错误：无法获取现有还原点：$_" }
            }

            if ($recentRestorePoints.Count -eq 0) {
                try {
                    Checkpoint-Computer -Description "Restore point created by Win11Debloat" -RestorePointType "MODIFY_SETTINGS"
                    return @{ Success = $true; Message = "系统还原点创建成功" }
                }
                catch {
                    return @{ Success = $false; Message = "错误：无法创建还原点：$_" }
                }
            }
            else {
                return @{ Success = $true; Message = "已存在最近的还原点，未创建新的还原点" }
            }
        }

        $createRestorePointJobDone = $createRestorePointJob | Wait-Job -TimeOut 20

        if (-not $createRestorePointJobDone) {
            Remove-Job -Job $createRestorePointJob -Force -ErrorAction SilentlyContinue
            Write-Host "错误：无法创建系统还原点，操作超时" -ForegroundColor Red
            $failed = $true
        }
        else {
            $result = Receive-Job $createRestorePointJob
            Remove-Job -Job $createRestorePointJob -ErrorAction SilentlyContinue
            if ($result.Success) {
                Write-Host $result.Message
            }
            else {
                Write-Host $result.Message -ForegroundColor Red
                $failed = $true
            }
        }
    }

    # Ensure that the user is aware if creating a restore point failed, and give them the option to continue without a restore point or cancel the script
    if ($failed) {
        if ($script:GuiWindow) {
            $result = Show-MessageBox "无法创建系统还原点。是否在没有还原点的情况下继续？" "还原点创建失败" "YesNo" "Warning"

            if ($result -ne "Yes") {
                $script:CancelRequested = $true
                return
            }
        }
        elseif (-not $Silent) {
            Write-Host "无法创建系统还原点。是否在没有还原点的情况下继续？(y/n)" -ForegroundColor Yellow
            if ($( Read-Host ) -ne 'y') {
                $script:CancelRequested = $true
                return
            }
        }

        Write-Host "警告：在没有还原点的情况下继续" -ForegroundColor Yellow
    }
}