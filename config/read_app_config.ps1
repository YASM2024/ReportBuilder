param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $true)]
    [string]$Key
)

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    return
}

$content = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
$pattern = "(?m)^\s*$([regex]::Escape($Key))\s*:\s*(.+)$"
if ($content -match $pattern) {
    $value = $Matches[1].Trim()
    if ($value.Length -ge 2) {
        if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
            $value = $value.Substring(1, $value.Length - 2)
        }
    }
    Write-Output $value
}
