#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$Silent,
    [switch]$Sysprep,
    [string]$LogPath,
    [string]$User,
    [switch]$CreateRestorePoint,
    [switch]$RunAppsListGenerator, [switch]$RunAppConfigurator,
    [switch]$RunDefaults, [switch]$RunWin11Defaults,
    [switch]$RunSavedSettings,
    [switch]$RemoveApps, 
    [switch]$RemoveAppsCustom,
    [switch]$RemoveGamingApps,
    [switch]$RemoveCommApps,
    [switch]$RemoveDevApps,
    [switch]$RemoveHPApps,
    [switch]$RemoveW11Outlook,
    [switch]$ForceRemoveEdge,
    [switch]$DisableDVR,
    [switch]$DisableTelemetry,
    [switch]$DisableFastStartup,
    [switch]$DisableBingSearches, [switch]$DisableBing,
    [switch]$DisableDesktopSpotlight,
    [switch]$DisableLockscrTips, [switch]$DisableLockscreenTips,
    [switch]$DisableWindowsSuggestions, [switch]$DisableSuggestions,
    [switch]$DisableSettings365Ads,
    [switch]$DisableSettingsHome,
    [switch]$ShowHiddenFolders,
    [switch]$ShowKnownFileExt,
    [switch]$HideDupliDrive,
    [switch]$EnableDarkMode,
    [switch]$DisableTransparency,
    [switch]$DisableAnimations,
    [switch]$TaskbarAlignLeft,
    [switch]$HideSearchTb, [switch]$ShowSearchIconTb, [switch]$ShowSearchLabelTb, [switch]$ShowSearchBoxTb,
    [switch]$HideTaskview,
    [switch]$DisableStartRecommended,
    [switch]$DisableStartPhoneLink,
    [switch]$DisableCopilot,
    [switch]$DisableRecall,
    [switch]$DisableWidgets, [switch]$HideWidgets,
    [switch]$DisableChat, [switch]$HideChat,
    [switch]$EnableEndTask,
    [switch]$ClearStart,
    [string]$ReplaceStart,
    [switch]$ClearStartAllUsers,
    [string]$ReplaceStartAllUsers,
    [switch]$RevertContextMenu,
    [switch]$DisableMouseAcceleration,
    [switch]$DisableStickyKeys,
    [switch]$HideHome,
    [switch]$HideGallery,
    [switch]$ExplorerToHome,
    [switch]$ExplorerToThisPC,
    [switch]$ExplorerToDownloads,
    [switch]$ExplorerToOneDrive,
    [switch]$DisableOnedrive, [switch]$HideOnedrive,
    [switch]$Disable3dObjects, [switch]$Hide3dObjects,
    [switch]$DisableMusic, [switch]$HideMusic,
    [switch]$DisableIncludeInLibrary, [switch]$HideIncludeInLibrary,
    [switch]$DisableGiveAccessTo, [switch]$HideGiveAccessTo,
    [switch]$DisableShare, [switch]$HideShare
)


# Show error if current powershell environment is limited by security policies
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
    Write-Host "错误：Win11Debloat 无法在您的系统上运行，PowerShell 执行受到安全策略限制" -ForegroundColor Red
    AwaitKeyToExit
}

# Log script output to 'Win11Debloat.log' at the specified path
if ($LogPath -and (Test-Path $LogPath)) {
    Start-Transcript -Path "$LogPath/Win11Debloat.log" -Append -IncludeInvocationHeader -Force | Out-Null
}
else {
    Start-Transcript -Path "$PSScriptRoot/Win11Debloat.log" -Append -IncludeInvocationHeader -Force | Out-Null
}

# Shows application selection form that allows the user to select what apps they want to remove or keep
function ShowAppSelectionForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    # Initialise form objects
    $form = New-Object System.Windows.Forms.Form
    $label = New-Object System.Windows.Forms.Label
    $button1 = New-Object System.Windows.Forms.Button
    $button2 = New-Object System.Windows.Forms.Button
    $selectionBox = New-Object System.Windows.Forms.CheckedListBox 
    $loadingLabel = New-Object System.Windows.Forms.Label
    $onlyInstalledCheckBox = New-Object System.Windows.Forms.CheckBox
    $checkUncheckCheckBox = New-Object System.Windows.Forms.CheckBox
    $initialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $script:selectionBoxIndex = -1

    # saveButton eventHandler
    $handler_saveButton_Click= 
    {
        if ($selectionBox.CheckedItems -contains "Microsoft.WindowsStore" -and -not $Silent) {
            $warningSelection = [System.Windows.Forms.Messagebox]::Show('Are you sure you wish to uninstall the Microsoft Store? This app cannot easily be reinstalled.', 'Are you sure?', 'YesNo', 'Warning')
        
            if ($warningSelection -eq 'No') {
                return
            }
        }

        $script:SelectedApps = $selectionBox.CheckedItems

        # Create file that stores selected apps if it doesn't exist
        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
            $null = New-Item "$PSScriptRoot/CustomAppsList"
        } 

        Set-Content -Path "$PSScriptRoot/CustomAppsList" -Value $script:SelectedApps

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    # cancelButton eventHandler
    $handler_cancelButton_Click= 
    {
        $form.Close()
    }

    $selectionBox_SelectedIndexChanged= 
    {
        $script:selectionBoxIndex = $selectionBox.SelectedIndex
    }

    $selectionBox_MouseDown=
    {
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
                if ($script:selectionBoxIndex -ne -1) {
                    $topIndex = $script:selectionBoxIndex

                    if ($selectionBox.SelectedIndex -gt $topIndex) {
                        for (($i = ($topIndex)); $i -le $selectionBox.SelectedIndex; $i++) {
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                    elseif ($topIndex -gt $selectionBox.SelectedIndex) {
                        for (($i = ($selectionBox.SelectedIndex)); $i -le $topIndex; $i++) {
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                }
            }
            elseif ($script:selectionBoxIndex -ne $selectionBox.SelectedIndex) {
                $selectionBox.SetItemChecked($selectionBox.SelectedIndex, -not $selectionBox.GetItemChecked($selectionBox.SelectedIndex))
            }
        }
    }

    $check_All=
    {
        for (($i = 0); $i -lt $selectionBox.Items.Count; $i++) {
            $selectionBox.SetItemChecked($i, $checkUncheckCheckBox.Checked)
        }
    }

    $load_Apps=
    {
        # Correct the initial state of the form to prevent the .Net maximized form issue
        $form.WindowState = $initialFormWindowState

        # Reset state to default before loading appslist again
        $script:selectionBoxIndex = -1
        $checkUncheckCheckBox.Checked = $False

        # Show loading indicator
        $loadingLabel.Visible = $true
        $form.Refresh()

        # Clear selectionBox before adding any new items
        $selectionBox.Items.Clear()

        # Set filePath where Appslist can be found
        $appsFile = "$PSScriptRoot/Appslist.txt"
        $listOfApps = ""

        if ($onlyInstalledCheckBox.Checked -and ($script:wingetInstalled -eq $true)) {
            # Attempt to get a list of installed apps via winget, times out after 10 seconds
            $job = Start-Job { return winget list --accept-source-agreements --disable-interactivity }
            $jobDone = $job | Wait-Job -TimeOut 10

            if (-not $jobDone) {
                # Show error that the script was unable to get list of apps from winget
                [System.Windows.MessageBox]::Show('无法通过 winget 加载已安装应用列表，某些应用可能不会显示在列表中。', '错误', 'Ok', 'Error')
            }
            else {
                # Add output of job (list of apps) to $listOfApps
                $listOfApps = Receive-Job -Job $job
            }
        }

        # Go through appslist and add items one by one to the selectionBox
        Foreach ($app in (Get-Content -Path $appsFile | Where-Object { $_ -notmatch '^\s*$' -and $_ -notmatch '^#  .*' -and $_ -notmatch '^# -* #' } )) { 
            $appChecked = $true

            # Remove first # if it exists and set appChecked to false
            if ($app.StartsWith('#')) {
                $app = $app.TrimStart("#")
                $appChecked = $false
            }

            # Remove any comments from the Appname
            if (-not ($app.IndexOf('#') -eq -1)) {
                $app = $app.Substring(0, $app.IndexOf('#'))
            }
            
            # Remove leading and trailing spaces and `*` characters from Appname
            $app = $app.Trim()
            $appString = $app.Trim('*')

            # Make sure appString is not empty
            if ($appString.length -gt 0) {
                if ($onlyInstalledCheckBox.Checked) {
                    # onlyInstalledCheckBox is checked, check if app is installed before adding it to selectionBox
                    if (-not ($listOfApps -like ("*$appString*")) -and -not (Get-AppxPackage -Name $app)) {
                        # App is not installed, continue with next item
                        continue
                    }
                    if (($appString -eq "Microsoft.Edge") -and -not ($listOfApps -like "* Microsoft.Edge *")) {
                        # App is not installed, continue with next item
                        continue
                    }
                }

                # Add the app to the selectionBox and set it's checked status
                $selectionBox.Items.Add($appString, $appChecked) | Out-Null
            }
        }
        
        # Hide loading indicator
        $loadingLabel.Visible = $False

        # Sort selectionBox alphabetically
        $selectionBox.Sorted = $True
    }

    $form.Text = "Win11Debloat 应用程序选择"
    $form.Name = "appSelectionForm"
    $form.DataBindings.DefaultDataSourceUpdateMode = 0
    $form.ClientSize = New-Object System.Drawing.Size(400,502)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $False

    $button1.TabIndex = 4
    $button1.Name = "saveButton"
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "确认"
    $button1.Location = New-Object System.Drawing.Point(27,472)
    $button1.Size = New-Object System.Drawing.Size(75,23)
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    $button1.add_Click($handler_saveButton_Click)

    $form.Controls.Add($button1)

    $button2.TabIndex = 5
    $button2.Name = "cancelButton"
    $button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "取消"
    $button2.Location = New-Object System.Drawing.Point(129,472)
    $button2.Size = New-Object System.Drawing.Size(75,23)
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $button2.add_Click($handler_cancelButton_Click)

    $form.Controls.Add($button2)

    $label.Location = New-Object System.Drawing.Point(13,5)
    $label.Size = New-Object System.Drawing.Size(400,14)
    $Label.Font = 'Microsoft Sans Serif,8'
    $label.Text = '勾选您要移除的应用，取消勾选您要保留的应用'

    $form.Controls.Add($label)

    $loadingLabel.Location = New-Object System.Drawing.Point(16,46)
    $loadingLabel.Size = New-Object System.Drawing.Size(300,418)
    $loadingLabel.Text = '正在加载应用...'
    $loadingLabel.BackColor = "White"
    $loadingLabel.Visible = $false

    $form.Controls.Add($loadingLabel)

    $onlyInstalledCheckBox.TabIndex = 6
    $onlyInstalledCheckBox.Location = New-Object System.Drawing.Point(230,474)
    $onlyInstalledCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $onlyInstalledCheckBox.Text = '仅显示已安装应用'
    $onlyInstalledCheckBox.add_CheckedChanged($load_Apps)

    $form.Controls.Add($onlyInstalledCheckBox)

    $checkUncheckCheckBox.TabIndex = 7
    $checkUncheckCheckBox.Location = New-Object System.Drawing.Point(16,22)
    $checkUncheckCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $checkUncheckCheckBox.Text = '全选/全不选'
    $checkUncheckCheckBox.add_CheckedChanged($check_All)

    $form.Controls.Add($checkUncheckCheckBox)

    $selectionBox.FormattingEnabled = $True
    $selectionBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $selectionBox.Name = "selectionBox"
    $selectionBox.Location = New-Object System.Drawing.Point(13,43)
    $selectionBox.Size = New-Object System.Drawing.Size(374,424)
    $selectionBox.TabIndex = 3
    $selectionBox.add_SelectedIndexChanged($selectionBox_SelectedIndexChanged)
    $selectionBox.add_Click($selectionBox_MouseDown)

    $form.Controls.Add($selectionBox)

    # Save the initial state of the form
    $initialFormWindowState = $form.WindowState

    # Load apps into selectionBox
    $form.add_Load($load_Apps)

    # Focus selectionBox when form opens
    $form.Add_Shown({$form.Activate(); $selectionBox.Focus()})

    # Show the Form
    return $form.ShowDialog()
}


# Returns list of apps from the specified file, it trims the app names and removes any comments
function ReadAppslistFromFile {
    param (
        $appsFilePath
    )

    $appsList = @()

    # Get list of apps from file at the path provided, and remove them one by one
    Foreach ($app in (Get-Content -Path $appsFilePath | Where-Object { $_ -notmatch '^#.*' -and $_ -notmatch '^\s*$' } )) { 
        # Remove any comments from the Appname
        if (-not ($app.IndexOf('#') -eq -1)) {
            $app = $app.Substring(0, $app.IndexOf('#'))
        }

        # Remove any spaces before and after the Appname
        $app = $app.Trim()
        
        $appString = $app.Trim('*')
        $appsList += $appString
    }

    return $appsList
}


# Removes apps specified during function call from all user accounts and from the OS image.
function RemoveApps {
    param (
        $appslist
    )

    Foreach ($app in $appsList) { 
        Write-Output "正在尝试移除 $app..."

        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            # Use winget to remove OneDrive and Edge
            if ($script:wingetInstalled -eq $false) {
                Write-Host "错误：WinGet 未安装或版本过旧，无法移除 $app" -ForegroundColor Red
            }
            else {
                # Uninstall app via winget
                Strip-Progress -ScriptBlock { winget uninstall --accept-source-agreements --disable-interactivity --id $app } | Tee-Object -Variable wingetOutput 

                If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "Uninstall failed with exit code")) {
                    Write-Host "无法通过 Winget 卸载 Microsoft Edge" -ForegroundColor Red
                    Write-Output ""

                    if ($( Read-Host -Prompt "您是否要强制卸载 Edge？不推荐这样做！(y/n)" ) -eq 'y') {
                        Write-Output ""
                        ForceRemoveEdge
                    }
                }
            }
        }
        else {
            # Use Remove-AppxPackage to remove all other apps
            $app = '*' + $app + '*'

            # Remove installed app for all existing users
            if ($WinVersion -ge 22000) {
                # Windows 11 build 22000 or later
                try {
                    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue

                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "已为所有用户移除 $app" -ForegroundColor DarkGray
                    }
                }
                catch {
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为所有用户移除 $app" -ForegroundColor Yellow
                        Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                    }
                }
            }
            else {
                # Windows 10
                try {
                    Get-AppxPackage -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
                    
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "已为当前用户移除 $app" -ForegroundColor DarkGray
                    }
                }
                catch {
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为当前用户移除 $app" -ForegroundColor Yellow
                        Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                    }
                }
                
                try {
                    Get-AppxPackage -Name $app -PackageTypeFilter Main, Bundle, Resource -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                    
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "已为所有用户移除 $app" -ForegroundColor DarkGray
                    }
                }
                catch {
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为所有用户移除 $app" -ForegroundColor Yellow
                        Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                    }
                }
            }

            # Remove provisioned app from OS image, so the app won't be installed for any new users
            try {
                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $app } | ForEach-Object { Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName }
            }
            catch {
                Write-Host "无法从 Windows 镜像中移除 $app" -ForegroundColor Yellow
                Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
            }
        }
    }
            
    Write-Output ""
}


# Forcefully removes Microsoft Edge using it's uninstaller
function ForceRemoveEdge {
    # Based on work from loadstring1 & ave9858
    Write-Output "> 正在强制卸载 Microsoft Edge..."

    $regView = [Microsoft.Win32.RegistryView]::Registry32
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
    $hklm.CreateSubKey('SOFTWARE\Microsoft\EdgeUpdateDev').SetValue('AllowUninstall', '')

    # Create stub (Creating this somehow allows uninstalling edge)
    $edgeStub = "$env:SystemRoot\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe"
    New-Item $edgeStub -ItemType Directory | Out-Null
    New-Item "$edgeStub\MicrosoftEdge.exe" | Out-Null

    # Remove edge
    $uninstallRegKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge')
    if ($null -ne $uninstallRegKey) {
        Write-Output "正在运行卸载程序..."
        $uninstallString = $uninstallRegKey.GetValue('UninstallString') + ' --force-uninstall'
        Start-Process cmd.exe "/c $uninstallString" -WindowStyle Hidden -Wait

        Write-Output "正在移除残留文件..."

        $edgePaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Tombstones\Microsoft Edge.lnk",
            "$env:PUBLIC\Desktop\Microsoft Edge.lnk",
            "$env:USERPROFILE\Desktop\Microsoft Edge.lnk",
            "$edgeStub"
        )

        foreach ($path in $edgePaths) {
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force -Recurse -ErrorAction SilentlyContinue
                Write-Host "  已移除 $path" -ForegroundColor DarkGray
            }
        }

        Write-Output "正在清理注册表..."

        # Remove ms edge from autostart
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "Microsoft Edge Update" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "Microsoft Edge Update" /f *>$null

        Write-Output "Microsoft Edge 已被卸载"
    }
    else {
        Write-Output ""
        Write-Host "错误：无法强制卸载 Microsoft Edge，找不到卸载程序" -ForegroundColor Red
    }
    
    Write-Output ""
}


# Execute provided command and strips progress spinners/bars from console output
function Strip-Progress {
    param(
        [ScriptBlock]$ScriptBlock
    )

    # Regex pattern to match spinner characters and progress bar patterns
    $progressPattern = 'Γû[Æê]|^\s+[-\\|/]\s+$'

    # Corrected regex pattern for size formatting, ensuring proper capture groups are utilized  
    $sizePattern = '(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB) /\s+(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB)'

    & $ScriptBlock 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            "错误: $($_.Exception.Message)"
        } else {
            $line = $_ -replace $progressPattern, '' -replace $sizePattern, ''
            if (-not ([string]::IsNullOrWhiteSpace($line)) -and -not ($line.StartsWith('  '))) {
                $line
            }
        }
    }
}


# Import & execute regfile
function RegImport {
    param (
        $message,
        $path
    )

    Write-Output $message

    if ($script:Params.ContainsKey("Sysprep")) {
        $defaultUserPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\NTUSER.DAT'
        
        reg load "HKU\Default" $defaultUserPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
    }
    elseif ($script:Params.ContainsKey("User")) {
        $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$($script:Params.Item("User"))\NTUSER.DAT"
        
        reg load "HKU\Default" $userPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
        
    }
    else {
        reg import "$PSScriptRoot\Regfiles\$path"  
    }

    Write-Output ""
}


# Restart the Windows Explorer process
function RestartExplorer {
    if ($script:Params.ContainsKey("Sysprep") -or $script:Params.ContainsKey("User")) {
        return
    }

    Write-Output "> 正在重启 Windows 资源管理器进程以应用所有更改...（这可能会导致屏幕闪烁）"

    if ($script:Params.ContainsKey("DisableMouseAcceleration")) {
        Write-Host "警告：增强指针精度设置的更改只会在重启后生效" -ForegroundColor Yellow
    }

    if ($script:Params.ContainsKey("DisableStickyKeys")) {
        Write-Host "警告：粘滞键设置的更改只会在重启后生效" -ForegroundColor Yellow
    }

    if ($script:Params.ContainsKey("DisableAnimations")) {
        Write-Host "警告：动画效果只会在重启后禁用" -ForegroundColor Yellow
    }

    # Only restart if the powershell process matches the OS architecture.
    # Restarting explorer from a 32bit PowerShell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem) {
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Warning "无法重启 Windows 资源管理器进程，请手动重启电脑以应用所有更改。"
    }
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenuForAllUsers {
    param (
        $startMenuTemplate = "$PSScriptRoot/Assets/Start/start2.bin"
    )

    Write-Output "> 正在为所有用户移除开始菜单中的所有固定应用..."

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "错误：无法清空开始菜单，脚本文件夹中缺少 start2.bin 文件" -ForegroundColor Red
        Write-Output ""
        return
    }

    # Get path to start menu file for all users
    $userPathString = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\*\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"
    $usersStartMenuPaths = get-childitem -path $userPathString

    # Go through all users and replace the start menu file
    ForEach ($startMenuPath in $usersStartMenuPaths) {
        ReplaceStartMenu $startMenuTemplate "$($startMenuPath.Fullname)\start2.bin"
    }

    # Also replace the start menu file for the default user profile
    $defaultStartMenuPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState'

    # Create folder if it doesn't exist
    if (-not (Test-Path $defaultStartMenuPath)) {
        new-item $defaultStartMenuPath -ItemType Directory -Force | Out-Null
        Write-Output "已为默认用户配置文件创建 LocalState 文件夹"
    }

    # Copy template to default profile
    Copy-Item -Path $startMenuTemplate -Destination $defaultStartMenuPath -Force
    Write-Output "已替换默认用户配置文件的开始菜单"
    Write-Output ""
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenu {
    param (
        $startMenuTemplate = "$PSScriptRoot/Assets/Start/start2.bin",
        $startMenuBinFile = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"
    )

    # Change path to correct user if a user was specified
    if ($script:Params.ContainsKey("User")) {
        $startMenuBinFile = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$(GetUserName)\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"
    }

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "错误：无法替换开始菜单，找不到模板文件" -ForegroundColor Red
        return
    }

    if ([IO.Path]::GetExtension($startMenuTemplate) -ne ".bin" ) {
        Write-Host "错误：无法替换开始菜单，模板文件不是有效的 .bin 文件" -ForegroundColor Red
        return
    }

    # Check if bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuBinFile)) {
        Write-Host "错误：无法替换用户 $(GetUserName) 的开始菜单，找不到原始 start2.bin 文件" -ForegroundColor Red
        return
    }

    $backupBinFile = $startMenuBinFile + ".bak"

    # Backup current start menu file
    Move-Item -Path $startMenuBinFile -Destination $backupBinFile -Force

    # Copy template file
    Copy-Item -Path $startMenuTemplate -Destination $startMenuBinFile -Force

    Write-Output "已替换用户 $(GetUserName) 的开始菜单"
}


# Add parameter to script and write to file
function AddParameter {
    param (
        $parameterName,
        $message
    )

    # Add key if it doesn't already exist
    if (-not $script:Params.ContainsKey($parameterName)) {
        $script:Params.Add($parameterName, $true)
    }

    # Create or clear file that stores last used settings
    if (-not (Test-Path "$PSScriptRoot/SavedSettings")) {
        $null = New-Item "$PSScriptRoot/SavedSettings"
    } 
    elseif ($script:FirstSelection) {
        $null = Clear-Content "$PSScriptRoot/SavedSettings"
    }
    
    $script:FirstSelection = $false

    # Create entry and add it to the file
    $entry = "$parameterName#- $message"
    Add-Content -Path "$PSScriptRoot/SavedSettings" -Value $entry
}


function PrintHeader {
    param (
        $title
    )

    $fullTitle = " Win11Debloat 脚本 - $title"

    if ($script:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle (Sysprep 模式)"
    }
    else {
        $fullTitle = "$fullTitle (用户: $(GetUserName))"
    }

    Clear-Host
    Write-Output "-------------------------------------------------------------------------------------------"
    Write-Output $fullTitle
    Write-Output "-------------------------------------------------------------------------------------------"
}


function PrintFromFile {
    param (
        $path,
        $title
    )

    Clear-Host

    PrintHeader $title

    # Get & print script menu from file
    Foreach ($line in (Get-Content -Path $path )) {   
        Write-Output $line
    }
}


function AwaitKeyToExit {
    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output "按任意键退出..."
        $null = [System.Console]::ReadKey()
    }

    Stop-Transcript
    Exit
}


function GetUserName {
    if ($script:Params.ContainsKey("User")) { 
        return $script:Params.Item("User") 
    }
    
    return $env:USERNAME
}


function CreateSystemRestorePoint {
    Write-Output "> 正在尝试创建系统还原点..."

    $SysRestore = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval"

    if ($SysRestore.RPSessionInterval -eq 0) {
        if ($Silent -or $( Read-Host -Prompt "系统还原已禁用，您是否要启用并创建一个还原点？(y/n)") -eq 'y') {
            try {
                Enable-ComputerRestore -Drive "$env:SystemDrive"
            } catch {
                Write-Host "错误：启用系统还原失败: $_" -ForegroundColor Red
                Write-Output ""
                return
            }
        } else {
            Write-Output ""
            return
        }
    }

    # Find existing restore points that are less than 24 hours old
    try {
        $recentRestorePoints = Get-ComputerRestorePoint | Where-Object { (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime) -le (New-TimeSpan -Hours 24) }
    } catch {
        Write-Host "错误：无法检索现有还原点: $_" -ForegroundColor Red
        Write-Output ""
        return
    }

    if ($recentRestorePoints.Count -eq 0) {
        try {
            Checkpoint-Computer -Description "Win11Debloat 创建的还原点" -RestorePointType "MODIFY_SETTINGS"
            Write-Output "系统还原点创建成功"
        } catch {
            Write-Host "错误：无法创建还原点: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "最近的还原点已存在，未创建新的还原点。" -ForegroundColor Yellow
    }

    Write-Output ""
}


function DisplayCustomModeOptions {
    # Get current Windows build version to compare against features
    $WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild
            
    PrintHeader '自定义模式'

    AddParameter 'CreateRestorePoint' '创建系统还原点'

    # Show options for removing apps, only continue on valid input
    Do {
        Write-Host "选项:" -ForegroundColor Yellow
        Write-Host " (n) 不移除任何应用" -ForegroundColor Yellow
        Write-Host " (1) 仅移除 'Appslist.txt' 中的默认选择的臃肿软件应用" -ForegroundColor Yellow
        Write-Host " (2) 移除默认选择的臃肿软件应用，以及邮件和日历应用、开发者应用和游戏应用"  -ForegroundColor Yellow
        Write-Host " (3) 手动选择要移除的应用" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host "您是否要移除任何应用？应用将为所有用户移除 (n/1/2/3)"

        # Show app selection form if user entered option 3
        if ($RemoveAppsInput -eq '3') {
            $result = ShowAppSelectionForm

            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                Write-Output ""
                Write-Host "已取消应用程序选择，请重试" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Output ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2' -and $RemoveAppsInput -ne '3') 

    # Select correct option based on user input
    switch ($RemoveAppsInput) {
        '1' {
            AddParameter 'RemoveApps' '移除默认选择的臃肿软件应用'
        }
        '2' {
            AddParameter 'RemoveApps' '移除默认选择的臃肿软件应用'
            AddParameter 'RemoveCommApps' '移除邮件、日历和联系人应用'
            AddParameter 'RemoveW11Outlook' '移除新的 Outlook for Windows 应用'
            AddParameter 'RemoveDevApps' '移除开发者相关应用'
            AddParameter 'RemoveGamingApps' '移除 Xbox 应用和 Xbox 游戏栏'
            AddParameter 'DisableDVR' '禁用 Xbox 游戏/屏幕录制'
        }
        '3' {
            Write-Output "您已选择移除 $($script:SelectedApps.Count) 个应用"

            AddParameter 'RemoveAppsCustom' "移除 $($script:SelectedApps.Count) 个应用:"

            Write-Output ""

            if ($( Read-Host -Prompt "禁用 Xbox 游戏/屏幕录制？这也会停止游戏覆盖弹窗 (y/n)" ) -eq 'y') {
                AddParameter 'DisableDVR' '禁用 Xbox 游戏/屏幕录制'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用遥测、诊断数据、活动历史记录、应用启动跟踪和定向广告？(y/n)" ) -eq 'y') {
        AddParameter 'DisableTelemetry' '禁用遥测、诊断数据、活动历史记录、应用启动跟踪和定向广告'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "在开始菜单、设置、通知、资源管理器和锁屏中禁用提示、技巧、建议和广告？(y/n)" ) -eq 'y') {
        AddParameter 'DisableSuggestions' '在开始菜单、设置、通知和文件资源管理器中禁用提示、技巧、建议和广告'
        AddParameter 'DisableSettings365Ads' '在设置主页中禁用 Microsoft 365 广告'
        AddParameter 'DisableLockscreenTips' '在锁屏上禁用提示和技巧'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "从 Windows 搜索中禁用并移除 Bing 网页搜索、Bing AI 和 Cortana？(y/n)" ) -eq 'y') {
        AddParameter 'DisableBing' '从 Windows 搜索中禁用并移除 Bing 网页搜索、Bing AI 和 Cortana'
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621) {
        Write-Output ""

        if ($( Read-Host -Prompt "禁用并移除 Microsoft Copilot 和 Windows Recall 快照？这适用于所有用户 (y/n)" ) -eq 'y') {
            AddParameter 'DisableCopilot' '禁用并移除 Microsoft Copilot'
            AddParameter 'DisableRecall' '禁用并移除 Windows Recall 快照'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用桌面上的 Windows 聚焦背景？(y/n)" ) -eq 'y') {
        AddParameter 'DisableDesktopSpotlight' '禁用 Windows 聚焦桌面背景选项'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "为系统和应用启用深色模式？(y/n)" ) -eq 'y') {
        AddParameter 'EnableDarkMode' '为系统和应用启用深色模式'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用透明度、动画和视觉效果？(y/n)" ) -eq 'y') {
        AddParameter 'DisableTransparency' '禁用透明度效果'
        AddParameter 'DisableAnimations' '禁用动画和视觉效果'
    }

    # Only show this option for Windows 11 users running build 22000 or later
    if ($WinVersion -ge 22000) {
        Write-Output ""

        if ($( Read-Host -Prompt "恢复旧的 Windows 10 样式上下文菜单？(y/n)" ) -eq 'y') {
            AddParameter 'RevertContextMenu' '恢复旧的 Windows 10 样式上下文菜单'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "关闭增强指针精度，也称为鼠标加速？(y/n)" ) -eq 'y') {
        AddParameter 'DisableMouseAcceleration' '关闭增强指针精度（鼠标加速）'
    }

    # Only show this option for Windows 11 users running build 26100 or later
    if ($WinVersion -ge 26100) {
        Write-Output ""

        if ($( Read-Host -Prompt "禁用粘滞键键盘快捷键？(y/n)" ) -eq 'y') {
            AddParameter 'DisableStickyKeys' '禁用粘滞键键盘快捷键'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "禁用快速启动？这适用于所有用户 (y/n)" ) -eq 'y') {
        AddParameter 'DisableFastStartup' '禁用快速启动'
    }

    # Only show option for disabling context menu items for Windows 10 users or if the user opted to restore the Windows 10 context menu
    if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -or $script:Params.ContainsKey('RevertContextMenu')) {
        Write-Output ""

        if ($( Read-Host -Prompt "您是否要禁用任何上下文菜单选项？(y/n)" ) -eq 'y') {
            Write-Output ""

            if ($( Read-Host -Prompt "   在上下文菜单中隐藏'包含在库中'选项？(y/n)" ) -eq 'y') {
                AddParameter 'HideIncludeInLibrary' "在上下文菜单中隐藏'包含在库中'选项"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   在上下文菜单中隐藏'授予访问权限'选项？(y/n)" ) -eq 'y') {
                AddParameter 'HideGiveAccessTo' "在上下文菜单中隐藏'授予访问权限'选项"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   在上下文菜单中隐藏'共享'选项？(y/n)" ) -eq 'y') {
                AddParameter 'HideShare' "在上下文菜单中隐藏'共享'选项"
            }
        }
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621) {
        Write-Output ""

        if ($( Read-Host -Prompt "您是否要对开始菜单进行任何更改？(y/n)" ) -eq 'y') {
            Write-Output ""

            if ($script:Params.ContainsKey("Sysprep")) {
                if ($( Read-Host -Prompt "从所有现有用户和新用户的开始菜单中移除所有固定应用？(y/n)" ) -eq 'y') {
                    AddParameter 'ClearStartAllUsers' '从现有用户和新用户的开始菜单中移除所有固定应用'
                }
            }
            else {
                Do {
                    Write-Host "   选项:" -ForegroundColor Yellow
                    Write-Host "    (n) 不从开始菜单中移除任何固定应用" -ForegroundColor Yellow
                    Write-Host "    (1) 仅从此用户 ($(GetUserName)) 的开始菜单中移除所有固定应用" -ForegroundColor Yellow
                    Write-Host "    (2) 从所有现有用户和新用户的开始菜单中移除所有固定应用"  -ForegroundColor Yellow
                    $ClearStartInput = Read-Host "   从开始菜单中移除所有固定应用？(n/1/2)" 
                }
                while ($ClearStartInput -ne 'n' -and $ClearStartInput -ne '0' -and $ClearStartInput -ne '1' -and $ClearStartInput -ne '2') 

                # Select correct option based on user input
                switch ($ClearStartInput) {
                    '1' {
                        AddParameter 'ClearStart' "仅从此用户的开始菜单中移除所有固定应用"
                    }
                    '2' {
                        AddParameter 'ClearStartAllUsers' "从所有现有用户和新用户的开始菜单中移除所有固定应用"
                    }
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   在开始菜单中禁用并隐藏推荐部分？这适用于所有用户 (y/n)" ) -eq 'y') {
                AddParameter 'DisableStartRecommended' '在开始菜单中禁用并隐藏推荐部分'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   在开始菜单中禁用 Phone Link 移动设备集成？(y/n)" ) -eq 'y') {
                AddParameter 'DisableStartPhoneLink' '在开始菜单中禁用 Phone Link 移动设备集成'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "您是否要对任务栏和相关服务进行任何更改？(y/n)" ) -eq 'y') {
        # Only show these specific options for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000) {
            Write-Output ""

            if ($( Read-Host -Prompt "   将任务栏按钮对齐到左侧？(y/n)" ) -eq 'y') {
                AddParameter 'TaskbarAlignLeft' '将任务栏图标对齐到左侧'
            }

            # Show options for search icon on taskbar, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   选项:" -ForegroundColor Yellow
                Write-Host "    (n) 无更改" -ForegroundColor Yellow
                Write-Host "    (1) 从任务栏隐藏搜索图标" -ForegroundColor Yellow
                Write-Host "    (2) 在任务栏上显示搜索图标" -ForegroundColor Yellow
                Write-Host "    (3) 在任务栏上显示带标签的搜索图标" -ForegroundColor Yellow
                Write-Host "    (4) 在任务栏上显示搜索框" -ForegroundColor Yellow
                $TbSearchInput = Read-Host "   隐藏或更改任务栏上的搜索图标？(n/1/2/3/4)" 
            }
            while ($TbSearchInput -ne 'n' -and $TbSearchInput -ne '0' -and $TbSearchInput -ne '1' -and $TbSearchInput -ne '2' -and $TbSearchInput -ne '3' -and $TbSearchInput -ne '4') 

            # Select correct taskbar search option based on user input
            switch ($TbSearchInput) {
                '1' {
                    AddParameter 'HideSearchTb' '从任务栏隐藏搜索图标'
                }
                '2' {
                    AddParameter 'ShowSearchIconTb' '在任务栏上显示搜索图标'
                }
                '3' {
                    AddParameter 'ShowSearchLabelTb' '在任务栏上显示带标签的搜索图标'
                }
                '4' {
                    AddParameter 'ShowSearchBoxTb' '在任务栏上显示搜索框'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   从任务栏隐藏任务视图按钮？(y/n)" ) -eq 'y') {
                AddParameter 'HideTaskview' '从任务栏隐藏任务视图按钮'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   禁用小组件服务并从任务栏隐藏图标？(y/n)" ) -eq 'y') {
            AddParameter 'DisableWidgets' '禁用小组件服务并从任务栏隐藏小组件（新闻和兴趣）图标'
        }

        # Only show this options for Windows users running build 22621 or earlier
        if ($WinVersion -le 22621) {
            Write-Output ""

            if ($( Read-Host -Prompt "   从任务栏隐藏聊天（立即开会）图标？(y/n)" ) -eq 'y') {
                AddParameter 'HideChat' '从任务栏隐藏聊天（立即开会）图标'
            }
        }
        
        # Only show this options for Windows users running build 22631 or later
        if ($WinVersion -ge 22631) {
            Write-Output ""

            if ($( Read-Host -Prompt "   在任务栏右键菜单中启用'结束任务'选项？(y/n)" ) -eq 'y') {
                AddParameter 'EnableEndTask' "在任务栏右键菜单中启用'结束任务'选项"
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "您是否要对文件资源管理器进行任何更改？(y/n)" ) -eq 'y') {
        # Show options for changing the File Explorer default location
        Do {
            Write-Output ""
            Write-Host "   选项:" -ForegroundColor Yellow
            Write-Host "    (n) 无更改" -ForegroundColor Yellow
            Write-Host "    (1) 将文件资源管理器打开到'主页'" -ForegroundColor Yellow
            Write-Host "    (2) 将文件资源管理器打开到'此电脑'" -ForegroundColor Yellow
            Write-Host "    (3) 将文件资源管理器打开到'下载'" -ForegroundColor Yellow
            Write-Host "    (4) 将文件资源管理器打开到'OneDrive'" -ForegroundColor Yellow
            $ExplSearchInput = Read-Host "   更改文件资源管理器打开的默认位置？(n/1/2/3/4)" 
        }
        while ($ExplSearchInput -ne 'n' -and $ExplSearchInput -ne '0' -and $ExplSearchInput -ne '1' -and $ExplSearchInput -ne '2' -and $ExplSearchInput -ne '3' -and $ExplSearchInput -ne '4') 

        # Select correct taskbar search option based on user input
        switch ($ExplSearchInput) {
            '1' {
                AddParameter 'ExplorerToHome' "将文件资源管理器打开的默认位置更改为'主页'"
            }
            '2' {
                AddParameter 'ExplorerToThisPC' "将文件资源管理器打开的默认位置更改为'此电脑'"
            }
            '3' {
                AddParameter 'ExplorerToDownloads' "将文件资源管理器打开的默认位置更改为'下载'"
            }
            '4' {
                AddParameter 'ExplorerToOneDrive' "将文件资源管理器打开的默认位置更改为'OneDrive'"
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   显示隐藏的文件、文件夹和驱动器？(y/n)" ) -eq 'y') {
            AddParameter 'ShowHiddenFolders' '显示隐藏的文件、文件夹和驱动器'
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   为已知文件类型显示文件扩展名？(y/n)" ) -eq 'y') {
            AddParameter 'ShowKnownFileExt' '为已知文件类型显示文件扩展名'
        }

        # Only show this option for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000) {
            Write-Output ""

            if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏主页部分？(y/n)" ) -eq 'y') {
                AddParameter 'HideHome' '从文件资源管理器侧面板中隐藏主页部分'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏库部分？(y/n)" ) -eq 'y') {
                AddParameter 'HideGallery' '从文件资源管理器侧面板中隐藏库部分'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏重复的可移动驱动器条目，使它们仅显示在此电脑下？(y/n)" ) -eq 'y') {
            AddParameter 'HideDupliDrive' '从文件资源管理器侧面板中隐藏重复的可移动驱动器条目'
        }

        # Only show option for disabling these specific folders for Windows 10 users
        if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") {
            Write-Output ""

            if ($( Read-Host -Prompt "您是否要从文件资源管理器侧面板中隐藏任何文件夹？(y/n)" ) -eq 'y') {
                Write-Output ""

                if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏 OneDrive 文件夹？(y/n)" ) -eq 'y') {
                    AddParameter 'HideOnedrive' '在文件资源管理器侧面板中隐藏 OneDrive 文件夹'
                }

                Write-Output ""
                
                if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏 3D 对象文件夹？(y/n)" ) -eq 'y') {
                    AddParameter 'Hide3dObjects' "在文件资源管理器中的'此电脑'下隐藏 3D 对象文件夹" 
                }
                
                Write-Output ""

                if ($( Read-Host -Prompt "   从文件资源管理器侧面板中隐藏音乐文件夹？(y/n)" ) -eq 'y') {
                    AddParameter 'HideMusic' "在文件资源管理器中的'此电脑'下隐藏音乐文件夹"
                }
            }
        }
    }

    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output "按回车键确认您的选择并执行脚本，或按 CTRL+C 退出..."
        Read-Host | Out-Null
    }

    PrintHeader '自定义模式'
}



##################################################################################################################
#                                                                                                                #
#                                                  脚本开始                                                      #
#                                                                                                                #
##################################################################################################################



# Check if winget is installed & if it is, check if the version is at least v1.4
if ((Get-AppxPackage -Name "*Microsoft.DesktopAppInstaller*") -and ([int](((winget -v) -replace 'v','').split('.')[0..1] -join '') -gt 14)) {
    $script:wingetInstalled = $true
}
else {
    $script:wingetInstalled = $false

    # Show warning that requires user confirmation, Suppress confirmation if Silent parameter was passed
    if (-not $Silent) {
        Write-Warning "Winget 未安装或版本过旧。这可能会阻止 Win11Debloat 移除某些应用。"
        Write-Output ""
        Write-Output "按任意键继续..."
        $null = [System.Console]::ReadKey()
    }
}

# Get current Windows build version to compare against features
$WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild

$script:Params = $PSBoundParameters
$script:FirstSelection = $true
$SPParams = 'WhatIf', 'Confirm', 'Verbose', 'Silent', 'Sysprep', 'Debug', 'User', 'CreateRestorePoint', 'LogPath'
$SPParamCount = 0

# Count how many SPParams exist within Params
# This is later used to check if any options were selected
foreach ($Param in $SPParams) {
    if ($script:Params.ContainsKey($Param)) {
        $SPParamCount++
    }
}

# Hide progress bars for app removal, as they block Win11Debloat's output
if (-not ($script:Params.ContainsKey("Verbose"))) {
    $ProgressPreference = 'SilentlyContinue'
}
else {
    Write-Host "详细模式已启用"
    Write-Output ""
    Write-Output "按任意键继续..."
    $null = [System.Console]::ReadKey()

    $ProgressPreference = 'Continue'
}

# Make sure all requirements for Sysprep are met, if Sysprep is enabled
if ($script:Params.ContainsKey("Sysprep")) {
    $defaultUserPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\NTUSER.DAT'

    # Exit script if default user directory or NTUSER.DAT file cannot be found
    if (-not (Test-Path "$defaultUserPath")) {
        Write-Host "错误：无法在 Sysprep 模式下启动 Win11Debloat，在 '$defaultUserPath' 找不到默认用户文件夹" -ForegroundColor Red
        AwaitKeyToExit
    }
    # Exit script if run in Sysprep mode on Windows 10
    if ($WinVersion -lt 22000) {
        Write-Host "错误：Win11Debloat Sysprep 模式不支持 Windows 10" -ForegroundColor Red
        AwaitKeyToExit
    }
}

# Make sure all requirements for User mode are met, if User is specified
if ($script:Params.ContainsKey("User")) {
    $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$($script:Params.Item("User"))\NTUSER.DAT"

    # Exit script if user directory or NTUSER.DAT file cannot be found
    if (-not (Test-Path "$userPath")) {
        Write-Host "错误：无法为用户 $($script:Params.Item("User")) 运行 Win11Debloat，在 '$userPath' 找不到用户数据" -ForegroundColor Red
        AwaitKeyToExit
    }
}

# Remove SavedSettings file if it exists and is empty
if ((Test-Path "$PSScriptRoot/SavedSettings") -and ([String]::IsNullOrWhiteSpace((Get-content "$PSScriptRoot/SavedSettings")))) {
    Remove-Item -Path "$PSScriptRoot/SavedSettings" -recurse
}

# Only run the app selection form if the 'RunAppsListGenerator' parameter was passed to the script
if ($RunAppConfigurator -or $RunAppsListGenerator) {
    PrintHeader "自定义应用列表生成器"

    $result = ShowAppSelectionForm

    # Show different message based on whether the app selection was saved or cancelled
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "应用程序选择窗口已关闭而未保存。" -ForegroundColor Red
    }
    else {
        Write-Output "您的应用选择已保存到 'CustomAppsList' 文件，位于："
        Write-Host "$PSScriptRoot" -ForegroundColor Yellow
    }

    AwaitKeyToExit
}

# Change script execution based on provided parameters or user input
if ((-not $script:Params.Count) -or $RunDefaults -or $RunWin11Defaults -or $RunSavedSettings -or ($SPParamCount -eq $script:Params.Count)) {
    if ($RunDefaults -or $RunWin11Defaults) {
        $Mode = '1'
    }
    elseif ($RunSavedSettings) {
        if (-not (Test-Path "$PSScriptRoot/SavedSettings")) {
            PrintHeader '自定义模式'
            Write-Host "错误：未找到保存的设置，未进行任何更改" -ForegroundColor Red
            AwaitKeyToExit
        }

        $Mode = '4'
    }
    else {
        # Show menu and wait for user input, loops until valid input is provided
        Do { 
            $ModeSelectionMessage = "请选择一个选项 (1/2/3/0)" 

            PrintHeader '菜单'

            Write-Output "(1) 默认模式：快速应用推荐的更改"
            Write-Output "(2) 自定义模式：手动选择要进行的更改"
            Write-Output "(3) 应用移除模式：选择并移除应用，不进行其他更改"

            # Only show this option if SavedSettings file exists
            if (Test-Path "$PSScriptRoot/SavedSettings") {
                Write-Output "(4) 应用上次保存的自定义设置"
                
                $ModeSelectionMessage = "请选择一个选项 (1/2/3/4/0)" 
            }

            Write-Output ""
            Write-Output "(0) 显示更多信息"
            Write-Output ""
            Write-Output ""

            $Mode = Read-Host $ModeSelectionMessage

            if ($Mode -eq '0') {
                # Print information screen from file
                PrintFromFile "$PSScriptRoot/Assets/Menus/Info_CN" "信息"

                Write-Output "按任意键返回..."
                $null = [System.Console]::ReadKey()
            }
            elseif (($Mode -eq '4') -and -not (Test-Path "$PSScriptRoot/SavedSettings")) {
                $Mode = $null
            }
        }
        while ($Mode -ne '1' -and $Mode -ne '2' -and $Mode -ne '3' -and $Mode -ne '4') 
    }

    # Add execution parameters based on the mode
    switch ($Mode) {
        # Default mode, loads defaults after confirmation
        '1' { 
            # Show the default settings with confirmation, unless Silent parameter was passed
            if (-not $Silent) {
                PrintFromFile "$PSScriptRoot/Assets/Menus/DefaultSettings_CN" "默认模式"

                Write-Output "按回车键执行脚本或按 CTRL+C 退出..."
                Read-Host | Out-Null
            }

            $DefaultParameterNames = 'CreateRestorePoint','RemoveApps','DisableTelemetry','DisableBing','DisableLockscreenTips','DisableSuggestions','ShowKnownFileExt','DisableWidgets','HideChat','DisableCopilot','DisableFastStartup'

            PrintHeader '默认模式'

            # Add default parameters, if they don't already exist
            foreach ($ParameterName in $DefaultParameterNames) {
                if (-not $script:Params.ContainsKey($ParameterName)) {
                    $script:Params.Add($ParameterName, $true)
                }
            }

            # Only add this option for Windows 10 users, if it doesn't already exist
            if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -and (-not $script:Params.ContainsKey('Hide3dObjects'))) {
                $script:Params.Add('Hide3dObjects', $Hide3dObjects)
            }
        }

        # Custom mode, show & add options based on user input
        '2' { 
            DisplayCustomModeOptions
        }

        # App removal, remove apps based on user selection
        '3' {
            PrintHeader "应用移除"

            $result = ShowAppSelectionForm

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Output "您已选择移除 $($script:SelectedApps.Count) 个应用"
                AddParameter 'RemoveAppsCustom' "移除 $($script:SelectedApps.Count) 个应用："

                # Suppress prompt if Silent parameter was passed
                if (-not $Silent) {
                    Write-Output ""
                    Write-Output ""
                    Write-Output "按回车键移除选定的应用或按 CTRL+C 退出..."
                    Read-Host | Out-Null
                    PrintHeader "应用移除"
                }
            }
            else {
                Write-Host "选择已取消，未移除任何应用" -ForegroundColor Red
                Write-Output ""
            }
        }

        # Load custom options from the "SavedSettings" file
        '4' {
            PrintHeader '自定义模式'
            Write-Output "Win11Debloat 将进行以下更改："

            # Print the saved settings info from file
            Foreach ($line in (Get-Content -Path "$PSScriptRoot/SavedSettings" )) { 
                # Remove any spaces before and after the line
                $line = $line.Trim()
            
                # Check if the line contains a comment
                if (-not ($line.IndexOf('#') -eq -1)) {
                    $parameterName = $line.Substring(0, $line.IndexOf('#'))

                    # Print parameter description and add parameter to Params list
                    if ($parameterName -eq "RemoveAppsCustom") {
                        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                            # Apps file does not exist, skip
                            continue
                        }
                        
                        $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
                        Write-Output "- 移除 $($appsList.Count) 个应用："
                        Write-Host $appsList -ForegroundColor DarkGray
                    }
                    else {
                        Write-Output $line.Substring(($line.IndexOf('#') + 1), ($line.Length - $line.IndexOf('#') - 1))
                    }

                    if (-not $script:Params.ContainsKey($parameterName)) {
                        $script:Params.Add($parameterName, $true)
                    }
                }
            }

            if (-not $Silent) {
                Write-Output ""
                Write-Output ""
                Write-Output "按回车键执行脚本或按 CTRL+C 退出..."
                Read-Host | Out-Null
            }

            PrintHeader '自定义模式'
        }
    }
}
else {
    PrintHeader '自定义模式'
}

# If the number of keys in SPParams equals the number of keys in Params then no modifications/changes were selected
#  or added by the user, and the script can exit without making any changes.
if ($SPParamCount -eq $script:Params.Keys.Count) {
    Write-Output "脚本已完成，未进行任何更改。"

    AwaitKeyToExit
}

# Execute all selected/provided parameters
switch ($script:Params.Keys) {
    'CreateRestorePoint' {
        CreateSystemRestorePoint
        continue
    }
    'RemoveApps' {
        $appsList = ReadAppslistFromFile "$PSScriptRoot/Appslist.txt" 
        Write-Output "> 正在移除默认选择的 $($appsList.Count) 个应用..."
        RemoveApps $appsList
        continue
    }
    'RemoveAppsCustom' {
        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
            Write-Host "> 错误：无法从文件加载自定义应用列表，未移除任何应用" -ForegroundColor Red
            Write-Output ""
            continue
        }
        
        $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
        Write-Output "> 正在移除 $($appsList.Count) 个应用..."
        RemoveApps $appsList
        continue
    }
    'RemoveCommApps' {
        $appsList = 'Microsoft.windowscommunicationsapps', 'Microsoft.People'
        Write-Output "> 正在移除邮件、日历和联系人应用..."
        RemoveApps $appsList
        continue
    }
    'RemoveW11Outlook' {
        $appsList = 'Microsoft.OutlookForWindows'
        Write-Output "> 正在移除新的 Outlook for Windows 应用..."
        RemoveApps $appsList
        continue
    }
    'RemoveDevApps' {
        $appsList = 'Microsoft.PowerAutomateDesktop', 'Microsoft.RemoteDesktop', 'Windows.DevHome'
        Write-Output "> 正在移除开发者相关应用..."
        RemoveApps $appsList
        continue
    }
    'RemoveGamingApps' {
        $appsList = 'Microsoft.GamingApp', 'Microsoft.XboxGameOverlay', 'Microsoft.XboxGamingOverlay'
        Write-Output "> 正在移除游戏相关应用..."
        RemoveApps $appsList
        continue
    }
    'RemoveHPApps' {
        $appsList = 'AD2F1837.HPAIExperienceCenter', 'AD2F1837.HPJumpStarts', 'AD2F1837.HPPCHardwareDiagnosticsWindows', 'AD2F1837.HPPowerManager', 'AD2F1837.HPPrivacySettings', 'AD2F1837.HPSupportAssistant', 'AD2F1837.HPSureShieldAI', 'AD2F1837.HPSystemInformation', 'AD2F1837.HPQuickDrop', 'AD2F1837.HPWorkWell', 'AD2F1837.myHP', 'AD2F1837.HPDesktopSupportUtilities', 'AD2F1837.HPQuickTouch', 'AD2F1837.HPEasyClean', 'AD2F1837.HPConnectedMusic', 'AD2F1837.HPFileViewer', 'AD2F1837.HPRegistration', 'AD2F1837.HPWelcome', 'AD2F1837.HPConnectedPhotopoweredbySnapfish', 'AD2F1837.HPPrinterControl'
        Write-Output "> 正在移除 HP 应用..."
        RemoveApps $appsList
        continue
    }
    "ForceRemoveEdge" {
        ForceRemoveEdge
        continue
    }
    'DisableDVR' {
        RegImport "> 正在禁用 Xbox 游戏/屏幕录制..." "Disable_DVR.reg"
        continue
    }
    'DisableTelemetry' {
        RegImport "> 正在禁用遥测、诊断数据、活动历史记录、应用启动跟踪和定向广告..." "Disable_Telemetry.reg"
        continue
    }
    {$_ -in "DisableSuggestions", "DisableWindowsSuggestions"} {
        RegImport "> 正在禁用整个 Windows 中的提示、技巧、建议和广告..." "Disable_Windows_Suggestions.reg"
        continue
    }
    {$_ -in "DisableLockscrTips", "DisableLockscreenTips"} {
        RegImport "> 正在禁用锁屏上的提示和技巧..." "Disable_Lockscreen_Tips.reg"
        continue
    }
    'DisableDesktopSpotlight' {
        RegImport "> 正在禁用 'Windows 聚焦' 桌面背景选项..." "Disable_Desktop_Spotlight.reg"
        continue
    }
    'DisableSettings365Ads' {
        RegImport "> 正在禁用设置主页中的 Microsoft 365 广告..." "Disable_Settings_365_Ads.reg"
        continue
    }
    'DisableSettingsHome' {
        RegImport "> 正在禁用设置主页..." "Disable_Settings_Home.reg"
        continue
    }
    {$_ -in "DisableBingSearches", "DisableBing"} {
        RegImport "> 正在从 Windows 搜索中禁用 Bing 网页搜索、Bing AI 和 Cortana..." "Disable_Bing_Cortana_In_Search.reg"
        
        # Also remove the app package for Bing search
        $appsList = 'Microsoft.BingSearch'
        RemoveApps $appsList
        continue
    }
    'DisableCopilot' {
        RegImport "> 正在禁用并移除 Microsoft Copilot..." "Disable_Copilot.reg"

        # Also remove the app package for Copilot
        $appsList = 'Microsoft.Copilot'
        RemoveApps $appsList
        continue
    }
    'DisableRecall' {
        RegImport "> 正在禁用 Windows Recall 快照..." "Disable_AI_Recall.reg"
        continue
    }
    'RevertContextMenu' {
        RegImport "> 正在恢复旧的 Windows 10 样式上下文菜单..." "Disable_Show_More_Options_Context_Menu.reg"
        continue
    }
    'DisableMouseAcceleration' {
        RegImport "> 正在关闭增强指针精度..." "Disable_Enhance_Pointer_Precision.reg"
        continue
    }
    'DisableStickyKeys' {
        RegImport "> 正在禁用粘滞键键盘快捷键..." "Disable_Sticky_Keys_Shortcut.reg"
        continue
    }
    'DisableFastStartup' {
        RegImport "> 正在禁用快速启动..." "Disable_Fast_Startup.reg"
        continue
    }
    'ClearStart' {
        Write-Output "> 正在为用户 $(GetUserName) 从开始菜单中移除所有固定应用..."
        ReplaceStartMenu
        Write-Output ""
        continue
    }
    'ReplaceStart' {
        Write-Output "> 正在为用户 $(GetUserName) 替换开始菜单..."
        ReplaceStartMenu $script:Params.Item("ReplaceStart")
        Write-Output ""
        continue
    }
    'ClearStartAllUsers' {
        ReplaceStartMenuForAllUsers
        continue
    }
    'ReplaceStartAllUsers' {
        ReplaceStartMenuForAllUsers $script:Params.Item("ReplaceStartAllUsers")
        continue
    }
    'DisableStartRecommended' {
        RegImport "> 正在禁用并隐藏开始菜单推荐部分..." "Disable_Start_Recommended.reg"
        continue
    }
    'DisableStartPhoneLink' {
        RegImport "> 正在禁用开始菜单中的 Phone Link 移动设备集成..." "Disable_Phone_Link_In_Start.reg"
        continue
    }
    'EnableDarkMode' {
        RegImport "> 正在为系统和应用启用深色模式..." "Enable_Dark_Mode.reg"
        continue
    }
    'DisableTransparency' {
        RegImport "> 正在禁用透明度效果..." "Disable_Transparency.reg"
        continue
    }
    'DisableAnimations' {
        RegImport "> 正在禁用动画和视觉效果..." "Disable_Animations.reg"
        continue
    }
    'TaskbarAlignLeft' {
        RegImport "> 正在将任务栏按钮对齐到左侧..." "Align_Taskbar_Left.reg"
        continue
    }
    'HideSearchTb' {
        RegImport "> 正在从任务栏隐藏搜索图标..." "Hide_Search_Taskbar.reg"
        continue
    }
    'ShowSearchIconTb' {
        RegImport "> 正在将任务栏搜索更改为仅图标..." "Show_Search_Icon.reg"
        continue
    }
    'ShowSearchLabelTb' {
        RegImport "> 正在将任务栏搜索更改为带标签的图标..." "Show_Search_Icon_And_Label.reg"
        continue
    }
    'ShowSearchBoxTb' {
        RegImport "> 正在将任务栏搜索更改为搜索框..." "Show_Search_Box.reg"
        continue
    }
    'HideTaskview' {
        RegImport "> 正在从任务栏隐藏任务视图按钮..." "Hide_Taskview_Taskbar.reg"
        continue
    }
    {$_ -in "HideWidgets", "DisableWidgets"} {
        RegImport "> 正在禁用小组件服务并从任务栏隐藏小组件图标..." "Disable_Widgets_Taskbar.reg"

        # Also remove the app package for Widgets
        $appsList = 'Microsoft.StartExperiencesApp'
        RemoveApps $appsList
        continue
    }
    {$_ -in "HideChat", "DisableChat"} {
        RegImport "> 正在从任务栏隐藏聊天图标..." "Disable_Chat_Taskbar.reg"
        continue
    }
    'EnableEndTask' {
        RegImport "> 正在启用任务栏右键菜单中的'结束任务'选项..." "Enable_End_Task.reg"
        continue
    }
    'ExplorerToHome' {
        RegImport "> 正在将文件资源管理器打开的默认位置更改为'主页'..." "Launch_File_Explorer_To_Home.reg"
        continue
    }
    'ExplorerToThisPC' {
        RegImport "> 正在将文件资源管理器打开的默认位置更改为'此电脑'..." "Launch_File_Explorer_To_This_PC.reg"
        continue
    }
    'ExplorerToDownloads' {
        RegImport "> 正在将文件资源管理器打开的默认位置更改为'下载'..." "Launch_File_Explorer_To_Downloads.reg"
        continue
    }
    'ExplorerToOneDrive' {
        RegImport "> 正在将文件资源管理器打开的默认位置更改为'OneDrive'..." "Launch_File_Explorer_To_OneDrive.reg"
        continue
    }
    'ShowHiddenFolders' {
        RegImport "> 正在显示隐藏的文件、文件夹和驱动器..." "Show_Hidden_Folders.reg"
        continue
    }
    'ShowKnownFileExt' {
        RegImport "> 正在为已知文件类型启用文件扩展名..." "Show_Extensions_For_Known_File_Types.reg"
        continue
    }
    'HideHome' {
        RegImport "> 正在从文件资源管理器导航面板中隐藏主页部分..." "Hide_Home_from_Explorer.reg"
        continue
    }
    'HideGallery' {
        RegImport "> 正在从文件资源管理器导航面板中隐藏库部分..." "Hide_Gallery_from_Explorer.reg"
        continue
    }
    'HideDupliDrive' {
        RegImport "> 正在从文件资源管理器导航面板中隐藏重复的可移动驱动器条目..." "Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer.reg"
        continue
    }
    {$_ -in "HideOnedrive", "DisableOnedrive"} {
        RegImport "> 正在从文件资源管理器导航面板中隐藏 OneDrive 文件夹..." "Hide_Onedrive_Folder.reg"
        continue
    }
    {$_ -in "Hide3dObjects", "Disable3dObjects"} {
        RegImport "> 正在从文件资源管理器导航面板中隐藏 3D 对象文件夹..." "Hide_3D_Objects_Folder.reg"
        continue
    }
    {$_ -in "HideMusic", "DisableMusic"} {
        RegImport "> 正在从文件资源管理器导航面板中隐藏音乐文件夹..." "Hide_Music_folder.reg"
        continue
    }
    {$_ -in "HideIncludeInLibrary", "DisableIncludeInLibrary"} {
        RegImport "> 正在在上下文菜单中隐藏'包含在库中'..." "Disable_Include_in_library_from_context_menu.reg"
        continue
    }
    {$_ -in "HideGiveAccessTo", "DisableGiveAccessTo"} {
        RegImport "> 正在在上下文菜单中隐藏'授予访问权限'..." "Disable_Give_access_to_context_menu.reg"
        continue
    }
    {$_ -in "HideShare", "DisableShare"} {
        RegImport "> 正在在上下文菜单中隐藏'共享'..." "Disable_Share_from_context_menu.reg"
        continue
    }
}

RestartExplorer

Write-Output ""
Write-Output ""
Write-Output ""
Write-Output "脚本已完成！请检查上方是否有任何错误。"

AwaitKeyToExit