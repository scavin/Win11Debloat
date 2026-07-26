<#
    .SYNOPSIS
        Shows a modal category-selection dialog for importing or exporting configuration.
#>
function Show-ImportExportConfigWindow {
    param (
        [System.Windows.Window]$Owner,
        [bool]$UsesDarkMode,
        [string]$Title,
        [string]$Prompt,
        [string[]]$Categories = @('应用程序', '系统调整', '部署设置'),
        [string[]]$DisabledCategories = @(),
        [hashtable]$CategoryDetails = @(),
        [string]$ActionLabel = 'OK'
    )

    # Show overlay on owner window
    $overlay = $null
    $overlayWasAlreadyVisible = $false
    try {
        $overlay = $Owner.FindName('ModalOverlay')
        if ($overlay) {
            $overlayWasAlreadyVisible = ($overlay.Visibility -eq 'Visible')
            if (-not $overlayWasAlreadyVisible) {
                $Owner.Dispatcher.Invoke([action]{ $overlay.Visibility = 'Visible' })
            }
        }
    } 
    catch { }

    # Load XAML from schema file
    $schemaPath = $script:ImportExportConfigSchema

    if (-not $schemaPath -or -not (Test-Path $schemaPath)) {
        Show-MessageBox -Message '找不到导入/导出窗口架构文件。' -Title '错误' -Button 'OK' -Icon 'Error' -Owner $Owner | Out-Null
        if ($overlay -and -not $overlayWasAlreadyVisible) {
            try { $Owner.Dispatcher.Invoke([action]{ $overlay.Visibility = 'Collapsed' }) } catch { }
        }
        return $null
    }

    $xaml = Get-Content -Path $schemaPath -Raw
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    try {
        $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    }
    finally {
        $reader.Close()
    }

    $dlg.Owner = $Owner
    Set-WindowThemeResources -window $dlg -usesDarkMode $UsesDarkMode

    # Copy the CheckBox default style from the main window so checkboxes get the themed template
    try {
        $mainCheckBoxStyle = $Owner.FindResource([type][System.Windows.Controls.CheckBox])
        if ($mainCheckBoxStyle) {
            $dlg.Resources.Add([type][System.Windows.Controls.CheckBox], $mainCheckBoxStyle)
        }
    }
    catch { }

    # Populate named elements
    $dlg.Title = $Title
    $dlg.FindName('TitleText').Text = $Title
    $dlg.FindName('PromptText').Text = $Prompt

    $titleBar = $dlg.FindName('TitleBar')
    $titleBar.Add_MouseLeftButtonDown({ $dlg.DragMove() })

    # Add a themed checkbox per category
    $checkboxPanel = $dlg.FindName('CheckboxPanel')
    $checkboxes = @{}
    foreach ($cat in $Categories) {
        # Create a container for the checkbox and details
        $container = New-Object System.Windows.Controls.StackPanel
        $container.Orientation = [System.Windows.Controls.Orientation]::Vertical
        $container.Margin = [System.Windows.Thickness]::new(0,0,0,12)

        # Create checkbox
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $cat
        $cb.IsChecked = $true
        $cb.Margin = [System.Windows.Thickness]::new(0,0,0,4)
        $cb.FontSize = 14
        $cb.FontWeight = [System.Windows.FontWeights]::Medium
        $cb.Foreground = $dlg.FindResource("AppFgColor")
        if ($DisabledCategories -contains $cat) {
            $cb.IsChecked = $false
            $cb.IsEnabled = $false
            $cb.Opacity = 0.65
            $cb.ToolTip = '此类别中没有可用的已选设置。'
        }
        
        $container.Children.Add($cb) | Out-Null
        
        # Add details if available
        if ($CategoryDetails -and $CategoryDetails[$cat]) {
            $detailsText = New-Object System.Windows.Controls.TextBlock
            $detailsText.Text = $CategoryDetails[$cat]
            $detailsText.FontSize = 12
            $detailsText.Foreground = $dlg.FindResource("AppFgColor")
            $detailsText.Margin = [System.Windows.Thickness]::new(32,0,0,0)
            $detailsText.Opacity = if ($DisabledCategories -contains $cat) { 0.45 } else { 0.75 }
            $detailsText.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $container.Children.Add($detailsText) | Out-Null
        }
        
        $checkboxPanel.Children.Add($container) | Out-Null
        $checkboxes[$cat] = $cb
    }

    $okBtn = $dlg.FindName('OkButton')
    $cancelBtn = $dlg.FindName('CancelButton')
    $okBtn.Content = $ActionLabel
    $okBtn.Add_Click({ $dlg.Tag = 'OK'; $dlg.Close() })
    $cancelBtn.Add_Click({ $dlg.Tag = 'Cancel'; $dlg.Close() })

    # Handle Escape key
    $dlg.Add_KeyDown({
        param($s, $e)
        if ($e.Key -eq 'Escape') { $dlg.Tag = 'Cancel'; $dlg.Close() }
    })

    try {
        $dlg.ShowDialog() | Out-Null
    }
    finally {
        # Hide overlay
        if ($overlay -and -not $overlayWasAlreadyVisible) {
            try { $Owner.Dispatcher.Invoke([action]{ $overlay.Visibility = 'Collapsed' }) } catch { }
        }
    }

    if ($dlg.Tag -ne 'OK') { return $null }

    $selected = @()
    foreach ($cat in $Categories) {
        if ($checkboxes[$cat].IsEnabled -and $checkboxes[$cat].IsChecked) { $selected += $cat }
    }
    if ($selected.Count -eq 0) { return $null }
    return $selected
}

function Get-SelectedApplications {
    param (
        [System.Windows.Controls.Panel]$AppsPanel
    )

    $selectedApps = @()
    foreach ($child in $AppsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked) {
            $selectedApps += $child.Tag
        }
    }

    return $selectedApps
}

function Get-SelectedTweakSettings {
    param (
        [System.Windows.Window]$Owner,
        [hashtable]$UiControlMappings
    )

    $tweakSettings = @()
    if (-not $UiControlMappings) {
        return $tweakSettings
    }

    foreach ($mappingKey in $UiControlMappings.Keys) {
        $control = $Owner.FindName($mappingKey)
        if (-not $control) { continue }

        $mapping = $UiControlMappings[$mappingKey]
        if ($control -is [System.Windows.Controls.CheckBox] -and $control.IsChecked) {
            if ($mapping.Type -eq 'feature') {
                $tweakSettings += @{ Name = $mapping.FeatureId; Value = $true }
            }
        }
        elseif ($control -is [System.Windows.Controls.ComboBox] -and $control.SelectedIndex -gt 0) {
            if ($mapping.Type -eq 'group') {
                $selectedValue = $mapping.Values[$control.SelectedIndex - 1]
                foreach ($fid in $selectedValue.FeatureIds) {
                    $tweakSettings += @{ Name = $fid; Value = $true }
                }
            }
            elseif ($mapping.Type -eq 'feature') {
                $tweakSettings += @{ Name = $mapping.FeatureId; Value = $true }
            }
        }
    }

    return $tweakSettings
}

function Get-DeploymentSettings {
    param (
        [System.Windows.Window]$Owner,
        [System.Windows.Controls.ComboBox]$UserSelectionCombo,
        [System.Windows.Controls.TextBox]$OtherUsernameTextBox
    )

    $deploySettings = @(
        @{ Name = 'UserSelectionIndex'; Value = $UserSelectionCombo.SelectedIndex }
    )

    if ($UserSelectionCombo.SelectedIndex -eq 1) {
        $deploySettings += @{ Name = 'OtherUsername'; Value = $OtherUsernameTextBox.Text.Trim() }
    }

    $appRemovalScopeCombo = $Owner.FindName('AppRemovalScopeCombo')
    if ($appRemovalScopeCombo) {
        $deploySettings += @{ Name = 'AppRemovalScopeIndex'; Value = $appRemovalScopeCombo.SelectedIndex }
    }

    $restorePointCheckBox = $Owner.FindName('RestorePointCheckBox')
    if ($restorePointCheckBox) {
        $deploySettings += @{ Name = 'CreateRestorePoint'; Value = [bool]$restorePointCheckBox.IsChecked }
    }

    $registryBackupCheckBox = $Owner.FindName('RegistryBackupCheckBox')
    if ($registryBackupCheckBox) {
        $deploySettings += @{ Name = 'SkipRegistryBackup'; Value = -not [bool]$registryBackupCheckBox.IsChecked }
    }

    $restartExplorerCheckBox = $Owner.FindName('RestartExplorerCheckBox')
    if ($restartExplorerCheckBox) {
        $deploySettings += @{ Name = 'RestartExplorer'; Value = [bool]$restartExplorerCheckBox.IsChecked }
    }

    return $deploySettings
}

function Get-AvailableImportExportCategories {
    param (
        $Config
    )

    $availableCategories = @()
    if ($Config.Apps) { $availableCategories += '应用程序' }
    if ($Config.Tweaks) { $availableCategories += '系统调整' }
    if ($Config.Deployment) { $availableCategories += '部署设置' }

    return $availableCategories
}

function Get-DeploymentCategoryDetailString {
    param (
        [array]$DeploymentSettings
    )

    $lookup = @{}
    foreach ($setting in @($DeploymentSettings)) {
        if ($setting -and $setting.Name) {
            $lookup[$setting.Name] = $setting.Value
        }
    }

    $line1 = @()

    if ($lookup.ContainsKey('UserSelectionIndex')) {
        switch ([int]$lookup['UserSelectionIndex']) {
            0 { $line1 += '用户：当前用户' }
            1 { $line1 += "用户：$(if ($lookup['OtherUsername']) { $lookup['OtherUsername'] } else { '其他用户' })" }
            2 { $line1 += '用户：Sysprep' }
        }
    }

    if ($lookup.ContainsKey('AppRemovalScopeIndex')) {
        switch ([int]$lookup['AppRemovalScopeIndex']) {
            0 { $line1 += '应用移除：所有用户' }
            1 { $line1 += '应用移除：当前用户' }
            2 { $line1 += "应用移除：$(if ($lookup['OtherUsername']) { $lookup['OtherUsername'] } else { '其他用户' })" }
        }
    }

    $options = @()
    if ($lookup.ContainsKey('CreateRestorePoint') -and [bool]$lookup['CreateRestorePoint']) { $options += '还原点' }
    if (-not ($lookup.ContainsKey('SkipRegistryBackup') -and [bool]$lookup['SkipRegistryBackup'])) { $options += '注册表备份' }
    if ($lookup.ContainsKey('RestartExplorer')    -and [bool]$lookup['RestartExplorer'])    { $options += '重启资源管理器' }

    $lines = @()
    if ($line1.Count -gt 0)   { $lines += $line1 -join ', ' }
    if ($options.Count -gt 0) { $lines += "选项：$($options -join '，')" }

    if ($lines.Count -gt 0) { return $lines -join "`n" }
    return '默认部署设置'
}

function Build-CategoryDetails {
    param (
        [int]$AppCount = 0,
        [int]$TweakCount = 0,
        [array]$DeploymentSettings
    )

    $details = @{}

    if ($AppCount -gt 0) {
        $details['应用程序'] = "已选择 $AppCount 个应用程序"
    }
    else {
        $details['应用程序'] = '未选择应用程序'
    }

    if ($TweakCount -gt 0) {
        $details['系统调整'] = "已选择 $TweakCount 项调整"
    }
    else {
        $details['系统调整'] = '未选择调整'
    }

    if ($DeploymentSettings) {
        $details['部署设置'] = Get-DeploymentCategoryDetailString -DeploymentSettings $DeploymentSettings
    }

    return $details
}

<#
    .SYNOPSIS
        Applies imported application selections to the application checkboxes.
#>
function Set-ImportedApplications {
    param (
        [System.Windows.Controls.Panel]$AppsPanel,
        [string[]]$AppIds
    )

    foreach ($child in $AppsPanel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox]) {
            $child.IsChecked = ($AppIds -contains $child.Tag)
        }
    }
}

<#
    .SYNOPSIS
        Applies imported tweak settings to their mapped UI controls.
#>
function Set-ImportedTweakSettings {
    param (
        [System.Windows.Window]$Owner,
        [hashtable]$UiControlMappings,
        [array]$TweakSettings
    )

    $settingsJson = [PSCustomObject]@{ Settings = @($TweakSettings) }
    Apply-SettingsToUiControls -window $Owner -settingsJson $settingsJson -uiControlMappings $UiControlMappings
}

<#
    .SYNOPSIS
        Applies imported deployment settings to the deployment controls.
#>
function Set-ImportedDeploymentSettings {
    param (
        [System.Windows.Window]$Owner,
        [System.Windows.Controls.ComboBox]$UserSelectionCombo,
        [System.Windows.Controls.TextBox]$OtherUsernameTextBox,
        [array]$DeploymentSettings
    )

    $lookup = @{}
    foreach ($setting in $DeploymentSettings) {
        $lookup[$setting.Name] = $setting.Value
    }

    if ($lookup.ContainsKey('UserSelectionIndex')) {
        $UserSelectionCombo.SelectedIndex = [int]$lookup['UserSelectionIndex']
    }
    if ($lookup.ContainsKey('OtherUsername') -and $UserSelectionCombo.SelectedIndex -eq 1) {
        $OtherUsernameTextBox.Text = $lookup['OtherUsername']
    }

    $appRemovalScopeCombo = $Owner.FindName('AppRemovalScopeCombo')
    if ($lookup.ContainsKey('AppRemovalScopeIndex') -and $appRemovalScopeCombo) {
        $appRemovalScopeCombo.SelectedIndex = [int]$lookup['AppRemovalScopeIndex']
    }

    $restorePointCheckBox = $Owner.FindName('RestorePointCheckBox')
    if ($lookup.ContainsKey('CreateRestorePoint') -and $restorePointCheckBox) {
        $restorePointCheckBox.IsChecked = [bool]$lookup['CreateRestorePoint']
    }

    $registryBackupCheckBox = $Owner.FindName('RegistryBackupCheckBox')
    if ($registryBackupCheckBox) {
        if ($lookup.ContainsKey('SkipRegistryBackup')) {
            $registryBackupCheckBox.IsChecked = -not [bool]$lookup['SkipRegistryBackup']
        }
    }

    $restartExplorerCheckBox = $Owner.FindName('RestartExplorerCheckBox')
    if ($lookup.ContainsKey('RestartExplorer') -and $restartExplorerCheckBox) {
        $restartExplorerCheckBox.IsChecked = [bool]$lookup['RestartExplorer']
    }
}

<#
    .SYNOPSIS
        Exports selected application, tweak, and deployment settings to a configuration file.
#>
function Export-Configuration {
    param (
        [System.Windows.Window]$Owner,
        [bool]$UsesDarkMode,
        [System.Windows.Controls.Panel]$AppsPanel,
        [hashtable]$UiControlMappings,
        [System.Windows.Controls.ComboBox]$UserSelectionCombo,
        [System.Windows.Controls.TextBox]$OtherUsernameTextBox
    )

    # Precompute exportable data so empty categories can be disabled in the picker.
    $selectedApps = Get-SelectedApplications -AppsPanel $AppsPanel
    $tweakSettings = Get-SelectedTweakSettings -Owner $Owner -UiControlMappings $UiControlMappings

    $disabledCategories = @()
    if ($selectedApps.Count -eq 0) { $disabledCategories += '应用程序' }
    if ($tweakSettings.Count -eq 0) { $disabledCategories += '系统调整' }

    $deploymentSettings = Get-DeploymentSettings -Owner $Owner -UserSelectionCombo $UserSelectionCombo -OtherUsernameTextBox $OtherUsernameTextBox
    $categoryDetails = Build-CategoryDetails -AppCount $selectedApps.Count -TweakCount $tweakSettings.Count -DeploymentSettings $deploymentSettings

    $categories = Show-ImportExportConfigWindow -Owner $Owner -UsesDarkMode $UsesDarkMode -Title '导出配置' -Prompt '基于当前已选设置创建配置文件。您可以选择要包含在导出中的设置类别。' -DisabledCategories $disabledCategories -CategoryDetails $categoryDetails -ActionLabel '导出设置'
    if (-not $categories) {
        Write-Host '导出已取消。'
        return
    }

    $config = @{ Version = '1.0' }

    if ($categories -contains '应用程序') {
        $config['Apps'] = @($selectedApps)
    }
    if ($categories -contains '系统调整') {
        $config['Tweaks'] = @($tweakSettings)
    }
    if ($categories -contains '部署设置') {
        $config['Deployment'] = @($deploymentSettings)
    }

    # Show native save-file dialog
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Title = '导出配置'
    $saveDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    $saveDialog.DefaultExt = '.json'
    $saveDialog.FileName = "Win11Debloat-Config-$(Get-Date -Format 'yyyyMMdd').json"

    if ($saveDialog.ShowDialog($Owner) -ne $true) {
        Write-Host '导出保存对话框已取消。'
        return
    }

    Write-Host "正在导出配置到 '$($saveDialog.FileName)'...（类别：$($categories -join '，')）"

    if ($script:Params.ContainsKey("WhatIf")) {
        Write-Host "[WhatIf] 导出配置到 '$($saveDialog.FileName)'" -ForegroundColor Cyan
        Show-MessageBox -Message "[WhatIf] 配置将被导出到此文件（未写入文件）。" -Title '导出配置' -Button 'OK' -Icon 'Information' | Out-Null
        return
    }

    if (Save-ToFile -Config $config -FilePath $saveDialog.FileName) {
        Write-Host "配置导出成功：$($saveDialog.FileName)"
        Show-MessageBox -Message "配置导出成功。" -Title '导出配置' -Button 'OK' -Icon 'Information' | Out-Null
    }
    else {
        Write-Error "导出配置到 '$($saveDialog.FileName)' 失败"
        Show-MessageBox -Message "导出配置失败" -Title '错误' -Button 'OK' -Icon 'Error' | Out-Null
    }
}

<#
    .SYNOPSIS
        Imports selected application, tweak, and deployment settings from a configuration file.
#>
function Import-Configuration {
    param (
        [System.Windows.Window]$Owner,
        [bool]$UsesDarkMode,
        [System.Windows.Controls.Panel]$AppsPanel,
        [hashtable]$UiControlMappings,
        [System.Windows.Controls.ComboBox]$UserSelectionCombo,
        [System.Windows.Controls.TextBox]$OtherUsernameTextBox,
        [scriptblock]$OnAppsImported,
        [scriptblock]$OnImportCompleted
    )

    # Show native open-file dialog
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Title = '选择配置文件'
    $openDialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    $openDialog.DefaultExt = '.json'

    if ($openDialog.ShowDialog($Owner) -ne $true) {
        Write-Host '导入文件对话框已取消。'
        return
    }

    Write-Host "正在从 '$($openDialog.FileName)' 导入配置..."

    $config = Import-JsonFile -filePath $openDialog.FileName -expectedVersion '1.0'
    if (-not $config) {
        Write-Error "读取配置文件 '$($openDialog.FileName)' 失败"
        Show-MessageBox -Message "读取配置文件失败" -Title '无效配置' -Button 'OK' -Icon 'Error' | Out-Null
        return
    }

    if (-not $config.Version) {
        Write-Error "配置文件格式无效：'$($openDialog.FileName)'"
        Show-MessageBox -Message "配置文件格式无效。" -Title '无效配置' -Button 'OK' -Icon 'Error' | Out-Null
        return
    }

    $availableCategories = Get-AvailableImportExportCategories -Config $config

    if ($availableCategories.Count -eq 0) {
        Write-Warning "配置文件 '$($openDialog.FileName)' 不包含可导入的数据。"
        Show-MessageBox -Message "所选文件不包含可导入的数据。" -Title '无效配置' -Button 'OK' -Icon 'Error' | Out-Null
        return
    }

    Write-Host "配置中的可用类别：$($availableCategories -join '，')"

    $appCount = @($config.Apps | Where-Object { $_ -is [string] -and -not [string]::IsNullOrWhiteSpace($_) }).Count
    $tweakCount = @($config.Tweaks | Where-Object { $_ -and $_.Name -and $_.Value -eq $true }).Count
    $categoryDetails = Build-CategoryDetails -AppCount $appCount -TweakCount $tweakCount -DeploymentSettings @($config.Deployment)

    $categories = Show-ImportExportConfigWindow -Owner $Owner -UsesDarkMode $UsesDarkMode -Title '导入配置' -Prompt '选择要导入的设置类别。您可以在应用前查看和修改导入的设置。' -Categories $availableCategories -CategoryDetails $categoryDetails -ActionLabel '导入设置'
    if (-not $categories) {
        Write-Host '导入已取消。'
        return
    }

    if ($categories -contains '应用程序' -and $config.Apps) {
        $appIds = @(
            $config.Apps | 
            Where-Object { $_ -is [string] } | 
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        Write-Host "正在导入 $($appIds.Count) 个应用选择。"
        Set-ImportedApplications -AppsPanel $AppsPanel -AppIds $appIds
        
        if ($OnAppsImported) { 
            & $OnAppsImported
        }
    }
    if ($categories -contains '系统调整' -and $config.Tweaks) {
        $tweakCount = @($config.Tweaks).Count
        Write-Host "正在导入 $tweakCount 项调整。"
        Set-ImportedTweakSettings -Owner $Owner -UiControlMappings $UiControlMappings -TweakSettings @($config.Tweaks)
    }
    if ($categories -contains '部署设置' -and $config.Deployment) {
        Write-Host '正在导入部署设置。'
        Set-ImportedDeploymentSettings -Owner $Owner -UserSelectionCombo $UserSelectionCombo -OtherUsernameTextBox $OtherUsernameTextBox -DeploymentSettings @($config.Deployment)
    }

    Write-Host '配置导入成功。'
    Show-MessageBox -Message "配置导入成功。" -Title '导入配置' -Button 'OK' -Icon 'Information' | Out-Null

    if ($OnImportCompleted) {
        & $OnImportCompleted $categories
    }
}
