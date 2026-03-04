# Prints the header for the script
function PrintHeader {
    param (
        $title
    )

    $fullTitle = " Win11Debloat Script - $title"

    if ($script:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle （Sysprep 模式）"
    }
    else {
        $fullTitle = "$fullTitle （用户：$(GetUserName)）"
    }

    Clear-Host
    Write-Host "-------------------------------------------------------------------------------------------"
    Write-Host $fullTitle
    Write-Host "-------------------------------------------------------------------------------------------"
}