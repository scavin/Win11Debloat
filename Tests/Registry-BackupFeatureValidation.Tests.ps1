BeforeAll {
    function Get-RegistryBackupCapturePlans {}

    . (Join-Path $PSScriptRoot '..\Scripts\Helpers\Registry-PathHelpers.ps1')
    . (Join-Path $PSScriptRoot '..\Scripts\Features\Registry-BackupValidation.ps1')
}

Describe 'Test-RegistryBackupMatchesSelectedFeatures' {
    BeforeEach {
        $script:Features = @{
            ApplyFeature = [PSCustomObject]@{ RegistryKey = 'apply.reg'; RegistryUndoKey = 'undo.reg' }
        }
        $script:CapturePlanCalls = [System.Collections.Generic.List[bool]]::new()
        Mock Get-RegistryBackupCapturePlans {
            param($SelectedRegistryFeatures, $UndoRegistryFeatures, $UseSysprepRegFiles)
            $script:CapturePlanCalls.Add([bool]$UseSysprepRegFiles)
            @([PSCustomObject]@{
                Path = 'HKEY_CURRENT_USER\Software\Example'
                IncludeSubKeys = $false
                CaptureAllValues = $false
                ValueNames = @('Enabled')
            })
        }
    }

    It 'derives an allow list from selected features and accepts a matching snapshot' {
        $snapshot = [PSCustomObject]@{
            Path = 'HKEY_CURRENT_USER\Software\Example'
            Values = @([PSCustomObject]@{ Name = 'Enabled'; Exists = $true; Kind = 'DWord'; Data = 1 })
            SubKeys = @()
        }

        $errors = @(Test-RegistryBackupMatchesSelectedFeatures -SelectedFeatureIds @('ApplyFeature') -SelectedUndoFeatureIds @() -Target 'CurrentUser' -RegistryKeys @($snapshot))

        $errors | Should -BeNullOrEmpty
        $script:CapturePlanCalls | Should -Be @($false)
    }

    It 'uses Sysprep registry files when restoring a user-profile backup' {
        $errors = @(Test-RegistryBackupMatchesSelectedFeatures -SelectedFeatureIds @('ApplyFeature') -SelectedUndoFeatureIds @() -Target 'User:Alice' -RegistryKeys @())

        $errors | Should -BeNullOrEmpty
        $script:CapturePlanCalls | Should -Be @($true)
    }

    It 'reports unknown selected features without deriving capture plans' {
        $errors = @(Test-RegistryBackupMatchesSelectedFeatures -SelectedFeatureIds @('UnknownFeature') -SelectedUndoFeatureIds @() -Target 'CurrentUser' -RegistryKeys @())

        $errors | Should -Contain "在当前功能目录中未找到所选功能 'UnknownFeature'。"
        $script:CapturePlanCalls | Should -BeNullOrEmpty
    }

    It 'rejects registry snapshots when no feature definitions are loaded' {
        $script:Features = @{}
        $snapshot = [PSCustomObject]@{ Path = 'HKEY_CURRENT_USER\Software\Example'; Values = @(); SubKeys = @() }

        $errors = @(Test-RegistryBackupMatchesSelectedFeatures -SelectedFeatureIds @('ApplyFeature') -SelectedUndoFeatureIds @() -Target 'CurrentUser' -RegistryKeys @($snapshot))

        $errors | Should -Contain '无法验证注册表备份允许列表，因为功能定义未加载。'
    }
}
