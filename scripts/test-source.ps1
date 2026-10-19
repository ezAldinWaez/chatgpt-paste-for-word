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
foreach ($script in Get-ChildItem -LiteralPath @((Join-Path $root 'src'), (Join-Path $root 'scripts')) -Filter '*.ps1') {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count) { throw "PowerShell syntax error in $($script.Name): $($errors[0])" }
}
# Validate the payload bit depth as well as the ICO directory. A directory
# advertising 32 bits with a 4-bit DIB payload produces a nearly empty icon.
$iconBytes = [IO.File]::ReadAllBytes((Join-Path $root 'assets/chatgpt-paste.ico'))
if ([BitConverter]::ToUInt16($iconBytes, 0) -ne 0 -or [BitConverter]::ToUInt16($iconBytes, 2) -ne 1) { throw 'Invalid Windows icon header' }
$iconSizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
if ([BitConverter]::ToUInt16($iconBytes, 4) -ne $iconSizes.Count) { throw 'Missing Windows icon sizes' }
for ($index = 0; $index -lt $iconSizes.Count; $index++) {
    $entry = 6 + 16 * $index
    $size = $iconSizes[$index]
    $dimension = if ($size -eq 256) { 0 } else { $size }
    if ($iconBytes[$entry] -ne $dimension -or $iconBytes[$entry + 1] -ne $dimension) { throw "Incorrect icon frame size: $size" }
    $offset = [BitConverter]::ToUInt32($iconBytes, $entry + 12)
    $length = [BitConverter]::ToUInt32($iconBytes, $entry + 8)
    if ($offset + $length -gt $iconBytes.Length) { throw "Truncated icon frame: $size" }
    if ($size -eq 256) {
        $signature = [BitConverter]::ToString($iconBytes, $offset, 8)
        if ($signature -ne '89-50-4E-47-0D-0A-1A-0A') { throw 'The 256 pixel icon must be a PNG' }
        if ($iconBytes[$offset + 24] -ne 8 -or $iconBytes[$offset + 25] -ne 6) { throw 'The large icon must use 8-bit RGBA channels' }
    }
    else {
        if ([BitConverter]::ToInt32($iconBytes, $offset + 4) -ne $size -or
            [BitConverter]::ToInt32($iconBytes, $offset + 8) -ne $size * 2 -or
            [BitConverter]::ToUInt16($iconBytes, $offset + 14) -ne 32) { throw "Icon payload is not a full-color 32-bit frame: $size" }
    }
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
