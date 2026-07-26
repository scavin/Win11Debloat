<#
    .SYNOPSIS
        Confirms removal of applications that require an extra safety warning.
#>
function Confirm-UnsafeAppRemoval {
    param (
        [string[]]$SelectedApps,
        $Owner = $null
    )

    # Skip all warnings in Silent mode
    if ($Silent) {
        return $true
    }

    # Microsoft Store warning
    if ($SelectedApps -contains "Microsoft.WindowsStore") {
        $result = Show-MessageBox -Message '确定要卸载 Microsoft Store 吗？此应用卸载后不易重新安装。' -Title '确定吗？' -Button 'YesNo' -Icon 'Warning' -Owner $Owner

        if ($result -ne 'Yes') {
            return $false
        }
    }

    # Windows Terminal warning
    if ($SelectedApps -contains "Microsoft.WindowsTerminal") {
        $result = Show-MessageBox -Message '确定要移除 Windows Terminal 吗？Windows Terminal 是 Windows 的默认命令行工具。请确保当前不是通过 Windows Terminal 运行 Win11Debloat，以免在操作过程中出现故障。' -Title '确定吗？' -Button 'YesNo' -Icon 'Warning' -Owner $Owner

        if ($result -ne 'Yes') {
            return $false
        }
    }

    return $true
}
