<#
    .SYNOPSIS
        Loads a registry backup from a JSON file and normalizes its contents.

    .DESCRIPTION
        Loads a registry backup from disk and returns a normalized representation
        of its contents suitable for use by the restore workflow. Throws if the
        file is missing, unreadable, or not valid JSON.

    .PARAMETER FilePath
        The absolute path to the registry backup JSON file to load.

    .OUTPUTS
        PSCustomObject
        A normalized registry backup object produced by ConvertTo-NormalizedRegistryBackup.
#>
function Import-RegistryBackup {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "未找到备份文件：$FilePath"
    }

    try {
        $rawBackup = Get-Content -LiteralPath $FilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "无法读取备份文件 '$FilePath'。该文件不是有效的 JSON。"
    }

    return ConvertTo-NormalizedRegistryBackup -Backup $rawBackup
}

<#
    .SYNOPSIS
        Validates and normalizes a raw registry backup object.

    .DESCRIPTION
        Validates the structure and content of the supplied backup and converts
        it into a normalized representation that can be safely consumed by the
        restore workflow. Throws if validation fails.

    .PARAMETER Backup
        The raw backup object (typically parsed from JSON) to normalize.

    .OUTPUTS
        PSCustomObject
        A normalized backup with Version, BackupType, CreatedAt, CreatedBy,
        ComputerName, Target, SelectedFeatures, SelectedUndoFeatures, and
        RegistryKeys properties.
#>
function ConvertTo-NormalizedRegistryBackup {
    param(
        [Parameter(Mandatory)]
        $Backup
    )

    $errors = New-Object System.Collections.Generic.List[string]

    if (-not $Backup.PSObject.Properties['Version']) {
        $errors.Add('缺少属性：Version')
    }
    elseif ([string]$Backup.Version -ne '1.0') {
        $errors.Add("不支持的备份版本 '$($Backup.Version)'。")
    }

    if (-not $Backup.PSObject.Properties['BackupType']) {
        $errors.Add('缺少属性：BackupType')
    }
    elseif ([string]$Backup.BackupType -ne 'RegistryState') {
        $errors.Add("不支持的备份类型 '$($Backup.BackupType)'。")
    }

    $normalizedTarget = ''
    if (-not $Backup.PSObject.Properties['Target'] -or [string]::IsNullOrWhiteSpace([string]$Backup.Target)) {
        $errors.Add('缺少属性：Target')
    }
    else {
        $normalizedTarget = [string]$Backup.Target

        if ($normalizedTarget -eq 'DefaultUserProfile') {
            # Valid target format.
        }
        elseif ($normalizedTarget -like 'User:*') {
            $targetUserName = $normalizedTarget.Substring(5)
            $targetValidation = Test-TargetUserName -UserName $targetUserName
            if (-not $targetValidation.IsValid) {
                $errors.Add("无效的用户 '$normalizedTarget'")
            }
        }
        elseif ($normalizedTarget -like 'CurrentUser:*') {
            $targetCurrentUserName = $normalizedTarget.Substring(12)
            if (Test-RunningAsSystem) {
                $errors.Add("此备份为用户 '$targetCurrentUserName' 创建且仅适用于该用户。请以该用户身份重新运行；SYSTEM 无法恢复 CurrentUser 备份。")
            }
            elseif ([string]::IsNullOrWhiteSpace($targetCurrentUserName) -or
                -not (Test-UserNameMatch -UserNameA $targetCurrentUserName -UserNameB $env:USERNAME)) {
                 $errors.Add("备份是为 '$targetCurrentUserName' 创建的，与当前用户 '$env:USERNAME' 不匹配。")
            }
        }
        else {
            $errors.Add("不支持的目标 '$normalizedTarget'。")
        }
    }

    $registryKeys = @()
    if (-not $Backup.PSObject.Properties['RegistryKeys']) {
        $errors.Add('缺少属性：RegistryKeys')
    }
    else {
        $registryKeys = @($Backup.RegistryKeys)
    }

    $normalizedKeys = @()
    foreach ($keySnapshot in $registryKeys) {
        $normalizedKeys += @(Normalize-RegistryKeySnapshot -Snapshot $keySnapshot)
    }

    $selectedFeatureParseResult = Get-NormalizedSelectedFeatureIdsFromBackup -Backup $Backup
    $selectedFeatures = @($selectedFeatureParseResult.SelectedFeatures)
    foreach ($selectedFeatureParseError in @($selectedFeatureParseResult.Errors)) {
        $errors.Add([string]$selectedFeatureParseError)
    }

    $selectedUndoFeatureParseResult = Get-NormalizedSelectedUndoFeatureIdsFromBackup -Backup $Backup
    $selectedUndoFeatures = @($selectedUndoFeatureParseResult.SelectedUndoFeatures)
    foreach ($selectedUndoFeatureParseError in @($selectedUndoFeatureParseResult.Errors)) {
        $errors.Add([string]$selectedUndoFeatureParseError)
    }

    $allSelectedFeatures = @($selectedFeatures) + @($selectedUndoFeatures)
    if ($allSelectedFeatures.Count -eq 0) {
        $errors.Add('备份必须在 SelectedFeatures 或 SelectedUndoFeatures 中包含至少一个功能 ID。')
    }
    else {
        try {
            $allowListValidationErrors = @(Test-RegistryBackupMatchesSelectedFeatures -SelectedFeatureIds @($selectedFeatures) -SelectedUndoFeatureIds @($selectedUndoFeatures) -Target $normalizedTarget -RegistryKeys @($normalizedKeys))
            foreach ($allowListValidationError in $allowListValidationErrors) {
                $errors.Add([string]$allowListValidationError)
            }
        }
        catch {
            $errors.Add("验证备份失败：$($_.Exception.Message)")
        }
    }

    if ($errors.Count -gt 0) {
        Write-Error "备份验证失败：$($errors -join ' ')"
        if ($errors.Count -eq 1) {
            throw ("验证失败：$($errors[0])")
        }
        else {
            throw ("验证失败，共 $($errors.Count) 个错误。详情请参见控制台输出。")
        }
    }

    return [PSCustomObject]@{
        Version = [string]$Backup.Version
        BackupType = [string]$Backup.BackupType
        CreatedAt = [string]$Backup.CreatedAt
        CreatedBy = [string]$Backup.CreatedBy
        ComputerName = [string]$Backup.ComputerName
        Target = $normalizedTarget
        SelectedFeatures = @($selectedFeatures)
        SelectedUndoFeatures = @($selectedUndoFeatures)
        RegistryKeys = @($normalizedKeys)
    }
}

<#
    .SYNOPSIS
        Restores registry state from a normalized backup object.

    .DESCRIPTION
        Applies the registry state described by the supplied backup back to the
        registry, loading the appropriate user hive when required.

    .PARAMETER Backup
        A normalized backup object (as produced by ConvertTo-NormalizedRegistryBackup) whose
        RegistryKeys snapshots should be restored.

    .OUTPUTS
        PSCustomObject
        Returns an object with a Result property set to $true when the restore
        completes successfully.
#>
function Restore-RegistryBackupState {
    param(
        [Parameter(Mandatory)]
        $Backup
    )

    $friendlyTarget = Get-FriendlyRegistryBackupTarget -Target ([string]$Backup.Target)

    if ($script:Params.ContainsKey("WhatIf")) {
        Write-Host "[WhatIf] 恢复 $friendlyTarget 的注册表备份" -ForegroundColor Cyan
        return [PSCustomObject]@{ Result = $true }
    }

    $restoreAction = {
        param($normalizedBackup)

        Write-Host "正在从 $(@($normalizedBackup.RegistryKeys).Count) 个根快照应用注册表恢复。"
        foreach ($rootSnapshot in @($normalizedBackup.RegistryKeys)) {
            Restore-RegistryKeySnapshot -Snapshot $rootSnapshot
        }
    }

    Write-Host "正在开始恢复 $friendlyTarget。"

    if ($Backup.Target -eq 'DefaultUserProfile' -or $Backup.Target -like 'User:*') {
        Write-Host "恢复需要加载目标用户配置单元。"
        Invoke-WithLoadedRestoreHive -Target $Backup.Target -ScriptBlock $restoreAction -ArgumentObject $Backup
        Write-Host "$friendlyTarget 的恢复已完成。"
        return [PSCustomObject]@{ Result = $true }
    }

    & $restoreAction $Backup
    Write-Host "$friendlyTarget 的恢复已完成。"
    return [PSCustomObject]@{ Result = $true }
}
