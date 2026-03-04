# Shows the CLI default mode app removal options. Loops until a valid option is selected.
function ShowCLIDefaultModeAppRemovalOptions {
    PrintHeader 'Default Mode'

    Write-Host "请注意：默认选择的应用包括 Microsoft Teams、Spotify、便笺等。选择选项 2 可查看和更改脚本将移除的应用" -ForegroundColor DarkGray
    Write-Host ""

    Do {
        Write-Host "选项：" -ForegroundColor Yellow
        Write-Host " (n) 不移除任何应用" -ForegroundColor Yellow
        Write-Host " (1) 仅移除默认选择的应用" -ForegroundColor Yellow
        Write-Host " (2) 手动选择要移除的应用" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host "是否要移除应用？应用将为所有用户移除 (n/1/2)"

        # Show app selection form if user entered option 3
        if ($RemoveAppsInput -eq '2') {
            $result = Show-AppSelectionWindow

            if ($result -ne $true) {
                # User cancelled or closed app selection, change RemoveAppsInput so the menu will be shown again
                Write-Host ""
                Write-Host "已取消应用选择，请重试" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Host ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2')

    return $RemoveAppsInput
}