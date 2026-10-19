param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$version = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim()
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $root 'dist' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
& (Join-Path $PSScriptRoot 'build.ps1') -OutputDirectory $OutputDirectory
$installer = Join-Path $OutputDirectory "ChatGPT-Paste-for-Word-Setup-$version.exe"
$digest = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
$name = Split-Path -Leaf $installer
$hashFile = Join-Path $OutputDirectory 'SHA256SUMS.txt'
[IO.File]::WriteAllText($hashFile, "$digest  $name`n", [Text.UTF8Encoding]::new($false))
Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination (Join-Path $OutputDirectory 'README.md') -Force
Copy-Item -LiteralPath (Join-Path $root 'LICENSE') -Destination (Join-Path $OutputDirectory 'LICENSE') -Force
Write-Host "Release files are ready in $OutputDirectory"
