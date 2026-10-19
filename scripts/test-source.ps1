$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$version = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim()
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Invalid VERSION: $version" }
$versionPattern = [regex]::Escape($version)
$checks = @(
    @{ File = 'src/Setup.cs'; Pattern = 'AssemblyVersion\("' + $versionPattern + '\.0"\)' },
    @{ File = 'src/Setup.cs'; Pattern = 'RibbonAssembly = "ChatGPTPasteRibbon-' + $versionPattern + '\.dll"' },
    @{ File = 'src/Setup.cs'; Pattern = 'RibbonVersion = "' + $versionPattern + '\.0"' },
    @{ File = 'src/Setup.cs'; Pattern = '"DisplayVersion", "' + $versionPattern + '"' },
    @{ File = 'src/ChatGPTPasteRibbon.cs'; Pattern = 'AssemblyVersion\("' + $versionPattern + '\.0"\)' },
    @{ File = 'src/ChatGPTPasteRibbon.cs'; Pattern = 'return "' + $versionPattern + '"' },
    @{ File = 'src/ChatGPTPasteHotkey.cs'; Pattern = 'AssemblyVersion\("' + $versionPattern + '\.0"\)' }
)
foreach ($check in $checks) {
    $path = Join-Path $root $check.File
    $content = Get-Content -LiteralPath $path -Raw
    if ($content -notmatch $check.Pattern) { throw "Version mismatch: $($check.File) ($($check.Pattern))" }
}
foreach ($script in Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.ps1') {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count) { throw "PowerShell syntax error in $($script.Name): $($errors[0])" }
}
$setup = Get-Content -LiteralPath (Join-Path $root 'src/Setup.cs') -Raw
foreach ($resource in @('ChatGPTPasteHotkey.exe', "ChatGPTPasteRibbon-$version.dll", 'Prepare-Word-Paste.ps1', 'Uninstall.ps1', 'Cleanup-Word-Paste.ps1')) {
    if (-not $setup.Contains("CopyResource(`"$resource`"")) {
        if ($resource -notlike 'ChatGPTPasteRibbon-*' -or -not $setup.Contains('CopyResource(RibbonAssembly,')) {
            throw "Installer does not embed/copy $resource"
        }
    }
}
Write-Host "Source checks passed for v$version."
