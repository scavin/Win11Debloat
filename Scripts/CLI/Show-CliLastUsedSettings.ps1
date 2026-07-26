# Shows the CLI last used settings from LastUsedSettings.json file, displays pending changes and prompts the user to apply them.
function Show-CliLastUsedSettings {
    Write-CliHeader '自定义模式'

    try {
        Import-Settings -filePath $script:SavedSettingsFilePath -expectedVersion "1.0"
    }
    catch {
        Write-Error "从 LastUsedSettings.json 文件加载设置失败：$_"
        Wait-ForKeyPress
    }

    if ($Silent) {
        # Skip change summary and confirmation prompt
        return
    }

    Write-PendingChanges
    Write-CliHeader '自定义模式'
}
