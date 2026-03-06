# Shows the CLI last used settings from LastUsedSettings.json file, displays pending changes and prompts the user to apply them.
function ShowCLILastUsedSettings {
    PrintHeader 'Custom Mode'

    try {
        # Load settings from LastUsedSettings.json and add to params
        LoadSettings -filePath $script:SavedSettingsFilePath -expectedVersion "1.0"
    }
    catch {
        Write-Error "从 LastUsedSettings.json 文件加载设置失败：$_"
        AwaitKeyToExit
    }

    PrintPendingChanges
    PrintHeader 'Custom Mode'
}