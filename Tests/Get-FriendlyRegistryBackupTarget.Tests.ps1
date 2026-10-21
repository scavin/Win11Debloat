BeforeAll {
    $friendlyTargetScriptPath = Join-Path $PSScriptRoot '..\Scripts\Helpers\Get-FriendlyRegistryBackupTarget.ps1'
    . $friendlyTargetScriptPath
    . (Join-Path $PSScriptRoot '..\Scripts\FileIO\Import-JsonFile.ps1')
    . (Join-Path $PSScriptRoot '..\Scripts\FileIO\Import-LanguageFile.ps1')
    $script:LanguagesPath = Join-Path $PSScriptRoot '..\Config\Languages'
    $script:Lang = Import-LanguageFile -LanguageCode 'en-US'
}

Describe 'Get-FriendlyRegistryBackupTarget' {

    It 'keeps unrecognized target text visible to the user' {
        Get-FriendlyRegistryBackupTarget -Target 'Custom:Value' | Should -Be 'Custom:Value'
    }
}
