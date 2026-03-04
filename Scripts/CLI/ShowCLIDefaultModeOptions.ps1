# Show CLI default mode options for removing apps, or set selection if RunDefaults or RunDefaultsLite parameter was passed
function ShowCLIDefaultModeOptions {
    if ($RunDefaults) {
        $RemoveAppsInput = '1'
    }
    elseif ($RunDefaultsLite) {
        $RemoveAppsInput = '0'                
    }
    else {
        $RemoveAppsInput = ShowCLIDefaultModeAppRemovalOptions

        if ($RemoveAppsInput -eq '2' -and ($script:SelectedApps.contains('Microsoft.XboxGameOverlay') -or $script:SelectedApps.contains('Microsoft.XboxGamingOverlay')) -and 
          $( Read-Host -Prompt "是否禁用 Game Bar 集成和游戏/屏幕录制？这也会阻止 ms-gamingoverlay 和 ms-gamebar 弹窗 (y/n)" ) -eq 'y') {
            $DisableGameBarIntegrationInput = $true;
        }
    }

    PrintHeader 'Default Mode'

    # Add default settings based on user input
    try {
        # Select app removal options based on user input
        switch ($RemoveAppsInput) {
            '1' {
                AddParameter 'RemoveApps'
                AddParameter 'Apps' 'Default'
            }
            '2' {
                AddParameter 'RemoveAppsCustom'

                if ($DisableGameBarIntegrationInput) {
                    AddParameter 'DisableDVR'
                    AddParameter 'DisableGameBarIntegration'
                }
            }
        }

        # Load settings from DefaultSettings.json and add to params
        LoadSettings -filePath $script:DefaultSettingsFilePath -expectedVersion "1.0"
    }
    catch {
        Write-Error "从 DefaultSettings.json 文件加载设置失败：$_"
        AwaitKeyToExit
    }

    SaveSettings

    # Skip change summary if Silent parameter was passed
    if ($Silent) {
        return
    }

    PrintPendingChanges
    PrintHeader 'Default Mode'
}