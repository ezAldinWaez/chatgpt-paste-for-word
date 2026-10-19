param(
    [string]$OfficeAssembly,
    [string]$ExtensibilityAssembly,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'test-source.ps1')
$version = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim()
$source = Join-Path $root 'src'
$assets = Join-Path $root 'assets'
$icon = Join-Path $assets 'chatgpt-paste.ico'
$ribbonImage = Join-Path $assets 'chatgpt-paste-32.png'
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $root 'dist' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$csc = Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) { throw 'The 64-bit .NET Framework C# compiler is required.' }
function Find-Assembly([string]$explicit, [string[]]$candidates, [string]$name) {
    if ($explicit) {
        if (-not (Test-Path -LiteralPath $explicit)) { throw "$name not found: $explicit" }
        return [IO.Path]::GetFullPath($explicit)
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            $match = Get-ChildItem -LiteralPath $candidate -Recurse -File -Filter "$name.dll" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($match) { return $match.FullName }
        }
    }
    throw "$name.dll was not found. Install the Office/Visual Studio PIAs, or pass its path explicitly."
}
$office = Find-Assembly $OfficeAssembly @(
    (Join-Path ${env:ProgramFiles} 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA'),
    (Join-Path $env:windir 'assembly\GAC_MSIL\office'),
    (Join-Path $env:windir 'Microsoft.NET\assembly\GAC_MSIL\office')
) 'office'
$ext = Find-Assembly $ExtensibilityAssembly @(
    (Join-Path $env:windir 'assembly\GAC\Extensibility'),
    (Join-Path $env:windir 'assembly\GAC_MSIL\Extensibility'),
    (Join-Path $env:windir 'Microsoft.NET\assembly\GAC_MSIL\Extensibility')
) 'Extensibility'
function Invoke-Compiler([string[]]$arguments) {
    & $csc @arguments
    if ($LASTEXITCODE -ne 0) { throw "C# compilation failed (exit code $LASTEXITCODE)." }
}
$hotkey = Join-Path $OutputDirectory 'ChatGPTPasteHotkey.exe'
$ribbon = Join-Path $OutputDirectory "ChatGPTPasteRibbon-$version.dll"
$installer = Join-Path $OutputDirectory "ChatGPT-Paste-for-Word-Setup-$version.exe"
Invoke-Compiler @('/nologo', '/target:winexe', '/platform:x64', "/out:$hotkey", "/win32icon:$icon", '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll', '/reference:Microsoft.CSharp.dll', (Join-Path $source 'ChatGPTPasteHotkey.cs'), (Join-Path $source 'WordPasteFormatting.cs'))
Invoke-Compiler @('/nologo', '/target:library', '/platform:anycpu', "/out:$ribbon", "/resource:$ribbonImage,ChatGPTPasteIcon.png", '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll', '/reference:Microsoft.CSharp.dll', "/link:$office", "/link:$ext", (Join-Path $source 'ChatGPTPasteRibbon.cs'), (Join-Path $source 'WordPasteFormatting.cs'))
$resources = @(
    "$hotkey,ChatGPTPasteHotkey.exe",
    "$ribbon,ChatGPTPasteRibbon-$version.dll"
)
foreach ($name in @('Prepare-Word-Paste.ps1', 'Uninstall.ps1', 'Cleanup-Word-Paste.ps1')) {
    $resources += "$(Join-Path $source $name),$name"
}
$args = @('/nologo', '/target:winexe', '/platform:x64', "/win32manifest:$(Join-Path $source 'app.manifest')", "/win32icon:$icon", "/out:$installer", '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll', '/reference:System.Core.dll', '/reference:System.IO.Compression.dll', '/reference:System.IO.Compression.FileSystem.dll')
$args += $resources | ForEach-Object { "/resource:$_" }
$args += Join-Path $source 'Setup.cs'
Invoke-Compiler $args
Write-Host "Built $installer"
