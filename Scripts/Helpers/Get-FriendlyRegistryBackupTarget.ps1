<#
    .SYNOPSIS
        Converts a registry-backup target identifier into a user-friendly label.
#>
function Get-FriendlyRegistryBackupTarget {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Target
    )

    if ([string]::IsNullOrWhiteSpace($Target)) {
        return '未知'
    }

    if ($Target -eq 'DefaultUserProfile') {
        return '默认用户配置文件'
    }

    if ($Target -eq 'CurrentUser') {
        return '当前用户'
    }

    if ($Target -eq 'AllUsers') {
        return '所有用户'
    }

    if ($Target -like 'CurrentUser:*') {
        $userName = $Target.Substring(12)
        if ([string]::IsNullOrWhiteSpace($userName)) {
            return '当前用户'
        }

        return "当前用户 ($userName)"
    }

    if ($Target -like 'User:*') {
        $userName = $Target.Substring(5)
        if ([string]::IsNullOrWhiteSpace($userName)) {
            return '用户'
        }

        return "用户 ($userName)"
    }

    return $Target
}
