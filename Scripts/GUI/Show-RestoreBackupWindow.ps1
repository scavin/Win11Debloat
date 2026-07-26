<#
    .SYNOPSIS
        Shows the backup-restore dialog and performs the selected restore.
#>
function Show-RestoreBackupWindow {
    param(
        [System.Windows.Window]$Owner = $null
    )

    try {
        Write-Host '正在打开恢复备份对话框。'

        $restoreResult = [PSCustomObject]@{
            RestoredRegistry = $false
            RestoredStartMenu = $false
        }

        $dialogResult = Show-RestoreBackupDialog -Owner $Owner
        if (-not $dialogResult -or $dialogResult.Result -eq 'Cancel') {
            Write-Host '用户已取消恢复。'
            return $restoreResult
        }

        $successMessage = $null
        $warningMessage = $null

        if ($dialogResult.Result -eq 'RestoreRegistry') {
            $backup = $dialogResult.Backup
            if (-not $backup) {
                throw '请求恢复注册表备份但未选择备份。'
            }

            Write-Host "用户确认恢复 $($backup.Target) 的注册表。"
            $restoreOpResult = Restore-RegistryBackupState -Backup $backup
            if ($restoreOpResult -and $restoreOpResult.Result) {
                $restoreResult.RestoredRegistry = $true
                if ($script:Params.ContainsKey("WhatIf")) {
                    $successMessage = '[WhatIf] 注册表备份将被恢复（未做任何更改）。'
                }
                else {
                    $successMessage = '注册表备份恢复成功。部分更改可能需要重启才能生效。'
                }
            }
        }
        elseif ($dialogResult.Result -eq 'Restore-StartMenu') {
            $scope = $dialogResult.StartMenuScope
            $useManualBackupFile = ($dialogResult.UseManualBackupFile -eq $true)
            $backupFilePath = $null
            if ($dialogResult -is [hashtable] -and $dialogResult.ContainsKey('BackupFilePath')) {
                $backupFilePath = $dialogResult['BackupFilePath']
            }
            elseif ($dialogResult.PSObject.Properties.Match('BackupFilePath').Count -gt 0) {
                $backupFilePath = $dialogResult.BackupFilePath
            }

            if ($useManualBackupFile -and [string]::IsNullOrWhiteSpace($backupFilePath)) {
                throw '开始菜单恢复已取消：未选择备份文件。'
            }

            $result = if ($scope -eq 'AllUsers') {
                Restore-StartMenuForAllUsers -BackupFilePath $backupFilePath
            }
            else {
                Restore-StartMenu -BackupFilePath $backupFilePath
            }

            $resultEntries = @($result)
            $successCount = @($resultEntries | Where-Object { $_.Result -eq $true }).Count
            $failedEntries = @($resultEntries | Where-Object { $_.Result -ne $true })

            if ($successCount -eq 0) {
                $errorSummary = ($resultEntries | ForEach-Object { $_.Message }) -join [Environment]::NewLine
                throw "恢复开始菜单备份失败。`n$errorSummary"
            }

            if ($failedEntries.Count -gt 0) {
                $failureSummary = ($failedEntries | ForEach-Object { $_.Message }) -join [Environment]::NewLine
                $warningMessage = "已成功为 $successCount 个用户恢复开始菜单备份。`n部分用户无法恢复：`n$failureSummary"
            }
            else {
                if ($script:Params.ContainsKey("WhatIf")) {
                    $successMessage = '[WhatIf] 开始菜单备份将被恢复（未做任何更改）。'
                }
                elseif ($scope -eq 'AllUsers') {
                    $successMessage = "已成功为所有用户恢复开始菜单备份。更改将在用户下次登录时生效。"
                }
                else {
                    $successMessage = "已成功为当前用户恢复开始菜单备份。更改将在您下次登录时生效。"
                }
            }

            $restoreResult.RestoredStartMenu = $true
        }

        if ($warningMessage) {
            Write-Host "$warningMessage"
            Show-MessageBox -Title '备份已恢复' -Message $warningMessage -Icon Warning
        }
        elseif ($successMessage) {
            Write-Host "$successMessage"
            Show-MessageBox -Title '备份已恢复' -Message $successMessage -Icon Success
        }

        return $restoreResult
    }
    catch {
        $errorMessage = if ($_.Exception.Message) { $_.Exception.Message } else { '发生意外错误。' }
        Write-Error "恢复操作失败：$errorMessage"
        Show-MessageBox -Title '错误' -Message "恢复失败：$errorMessage" -Icon Error
        return [PSCustomObject]@{
            RestoredRegistry = $false
            RestoredStartMenu = $false
        }
    }
}
