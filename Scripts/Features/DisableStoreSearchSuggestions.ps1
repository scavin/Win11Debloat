# Disables Microsoft Store search suggestions in the start menu for all users by denying access to the Store app database file for each user
function DisableStoreSearchSuggestionsForAllUsers {
    # Get path to Store app database for all users
    $userPathString = GetUserDirectory -userName "*" -fileName "AppData\Local\Packages"
    $usersStoreDbPaths = get-childitem -path $userPathString

    # Go through all users and disable start search suggestions
    ForEach ($storeDbPath in $usersStoreDbPaths) {
        DisableStoreSearchSuggestions ($storeDbPath.FullName + "\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\store.db")
    }

    # Also disable start search suggestions for the default user profile
    $defaultStoreDbPath = GetUserDirectory -userName "Default" -fileName "AppData\Local\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\store.db" -exitIfPathNotFound $false
    DisableStoreSearchSuggestions $defaultStoreDbPath
}


# Disables Microsoft Store search suggestions in the start menu by denying access to the Store app database file
function DisableStoreSearchSuggestions {
    param (
        $StoreAppsDatabase = "$env:LocalAppData\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\store.db"
    )

    # Change path to correct user if a user was specified
    if ($script:Params.ContainsKey("User")) {
        $StoreAppsDatabase = GetUserDirectory -userName "$(GetUserName)" -fileName "AppData\Local\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe\LocalState\store.db" -exitIfPathNotFound $false
    }

    $userName = [regex]::Match($StoreAppsDatabase, '(?:Users\\)([^\\]+)(?:\\AppData)').Groups[1].Value

    # This file doesn't exist in EEA (No Store app suggestions).
    if (-not (Test-Path -Path $StoreAppsDatabase))
    {
        Write-Host "找不到用户 $userName 的 Store 应用数据库，正在创建以防止 Windows 后续自动创建..." -ForegroundColor Yellow

        $storeDbDir = Split-Path -Path $StoreAppsDatabase -Parent

        if (-not (Test-Path -Path $storeDbDir)) {
            New-Item -Path $storeDbDir -ItemType Directory -Force | Out-Null
        }

        New-Item -Path $StoreAppsDatabase -ItemType File -Force | Out-Null
    }
    
    $AccountSid = [System.Security.Principal.SecurityIdentifier]::new('S-1-1-0') # 'EVERYONE' group
    $Acl = Get-Acl -Path $StoreAppsDatabase
    $Ace = [System.Security.AccessControl.FileSystemAccessRule]::new($AccountSid, 'FullControl', 'Deny')
    $Acl.SetAccessRule($Ace) | Out-Null
    Set-Acl -Path $StoreAppsDatabase -AclObject $Acl | Out-Null

    Write-Host "已禁用用户 $userName 的 Microsoft Store 搜索建议"
}