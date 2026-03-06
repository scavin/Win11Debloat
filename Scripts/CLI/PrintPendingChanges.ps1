# Prints all pending changes that will be made by the script
function PrintPendingChanges {
    Write-Output "Win11Debloat 将进行以下更改："

    if ($script:Params['CreateRestorePoint']) {
        Write-Output "- $($script:Features['CreateRestorePoint'].Label)"
    }
    foreach ($parameterName in $script:Params.Keys) {
        if ($script:ControlParams -contains $parameterName) {
            continue
        }

        # Print parameter description
        switch ($parameterName) {
            'Apps' {
                continue
            }
            'CreateRestorePoint' {
                continue
            }
            'RemoveApps' {
                $appsList = GenerateAppsList

                if ($appsList.Count -eq 0) {
                    Write-Host "未选择任何有效的应用进行移除" -ForegroundColor Yellow
                    Write-Output ""
                    continue
                }

                Write-Output "- 移除 $($appsList.Count) 个应用："
                Write-Host $appsList -ForegroundColor DarkGray
                continue
            }
            'RemoveAppsCustom' {
                $appsList = LoadAppsFromFile $script:CustomAppsListFilePath

                if ($appsList.Count -eq 0) {
                    Write-Host "未选择任何有效的应用进行移除" -ForegroundColor Yellow
                    Write-Output ""
                    continue
                }

                Write-Output "- 移除 $($appsList.Count) 个应用："
                Write-Host $appsList -ForegroundColor DarkGray
                continue
            }
            default {
                if ($script:Features -and $script:Features.ContainsKey($parameterName)) {
                    $action = $script:Features[$parameterName].Action
                    $message = $script:Features[$parameterName].Label
                    Write-Output "- $action $message"
                }
                else {
                    # Fallback: show the parameter name if no feature description is available
                    Write-Output "- $parameterName"
                }
                continue
            }
        }
    }

    Write-Output ""
    Write-Output ""
    Write-Output "按 Enter 键执行脚本，或按 CTRL+C 退出..."
    Read-Host | Out-Null
}