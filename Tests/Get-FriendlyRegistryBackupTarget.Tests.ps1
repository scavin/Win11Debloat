BeforeAll {
    $friendlyTargetScriptPath = Join-Path $PSScriptRoot '..\Scripts\Helpers\Get-FriendlyRegistryBackupTarget.ps1'
    . $friendlyTargetScriptPath
}

Describe 'Get-FriendlyRegistryBackupTarget' {
    It 'formats <Case> as <Expected>' -ForEach @(
        @{ Case = 'a null target'; Target = $null; Expected = '未知' }
        @{ Case = 'the default profile'; Target = 'DefaultUserProfile'; Expected = '默认用户配置文件' }
        @{ Case = 'the current-user marker'; Target = 'CurrentUser'; Expected = '当前用户' }
        @{ Case = 'the all-users marker'; Target = 'AllUsers'; Expected = '所有用户' }
        @{ Case = 'a named current user'; Target = 'CurrentUser:Alice'; Expected = '当前用户 (Alice)' }
        @{ Case = 'a named target user'; Target = 'User:Bob'; Expected = '用户 (Bob)' }
    ) {
        Get-FriendlyRegistryBackupTarget -Target $Target | Should -Be $Expected
    }

    It 'keeps unrecognized target text visible to the user' {
        Get-FriendlyRegistryBackupTarget -Target 'Custom:Value' | Should -Be 'Custom:Value'
    }
}
