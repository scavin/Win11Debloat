# Removes apps specified during function call based on the target scope.
function RemoveApps {
    param (
        $appslist
    )

    # Determine target from script-level params, defaulting to AllUsers
    $targetUser = GetTargetUserForAppRemoval

    $appIndex = 0
    $appCount = @($appsList).Count

    Foreach ($app in $appsList) {
        if ($script:CancelRequested) {
            return
        }

        $appIndex++

        # Update step name and sub-progress to show which app is being removed (only for bulk removal)
        if ($script:ApplySubStepCallback -and $appCount -gt 1) {
            & $script:ApplySubStepCallback "正在卸载应用 ($appIndex/$appCount)" $appIndex $appCount
        }

        Write-Host "正在尝试卸载 $app..."

        # Use WinGet only to remove OneDrive and Edge
        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            if ($script:WingetInstalled -eq $false) {
                Write-Host "WinGet 未安装或版本过旧，无法卸载 $app" -ForegroundColor Red
                continue
            }

            $appName = $app -replace '\.', '_'

            # Uninstall app via WinGet, or create a scheduled task to uninstall it later
            if ($script:Params.ContainsKey("User")) {
                ImportRegistryFile "正在添加计划任务以卸载用户 $(GetUserName) 的 $app..." "Uninstall_$($appName).reg"
            }
            elseif ($script:Params.ContainsKey("Sysprep")) {
                ImportRegistryFile "正在添加计划任务以在新用户登录后卸载 $app..." "Uninstall_$($appName).reg"
            }
            else {
                # Uninstall app via WinGet
                $wingetOutput = Invoke-NonBlocking -ScriptBlock {
                    param($appId)
                    winget uninstall --accept-source-agreements --disable-interactivity --id $appId
                } -ArgumentList $app

                If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "Uninstall failed with exit code")) {
                    Write-Host "无法通过 WinGet 卸载 Microsoft Edge" -ForegroundColor Red

                    if ($script:GuiWindow) {
                        $result = Show-MessageBox -Message '无法通过 WinGet 卸载 Microsoft Edge。是否要强制卸载？不推荐！' -Title '强制卸载 Microsoft Edge？' -Button 'YesNo' -Icon 'Warning'

                        if ($result -eq 'Yes') {
                            Write-Host ""
                            ForceRemoveEdge
                        }
                    }
                    elseif ($( Read-Host -Prompt "是否要强制卸载 Microsoft Edge？不推荐！(y/n)" ) -eq 'y') {
                        Write-Host ""
                        ForceRemoveEdge
                    }
                }
            }

            continue
        }

        # Use Remove-AppxPackage to remove all other apps
        $appPattern = '*' + $app + '*'

        try {
            switch ($targetUser) {
                "AllUsers" {
                    # Remove installed app for all existing users, and from OS image
                    Invoke-NonBlocking -ScriptBlock {
                        param($pattern)
                        Get-AppxPackage -Name $pattern -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue
                        Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $pattern } | ForEach-Object { Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName }
                    } -ArgumentList $appPattern
                }
                "CurrentUser" {
                    # Remove installed app for current user only
                    Invoke-NonBlocking -ScriptBlock {
                        param($pattern)
                        Get-AppxPackage -Name $pattern | Remove-AppxPackage -ErrorAction Continue
                    } -ArgumentList $appPattern
                }
                default {
                    # Target is a specific username - remove app for that user only
                    Invoke-NonBlocking -ScriptBlock {
                        param($pattern, $user)
                        $userAccount = New-Object System.Security.Principal.NTAccount($user)
                        $userSid = $userAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                        Get-AppxPackage -Name $pattern -User $userSid | Remove-AppxPackage -User $userSid -ErrorAction Continue
                    } -ArgumentList @($appPattern, $targetUser)
                }
            }
        }
        catch {
            if ($DebugPreference -ne "SilentlyContinue") {
                Write-Host "尝试卸载 $app 时出现问题" -ForegroundColor Yellow
                Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
            }
        }
    }

    Write-Host ""
}