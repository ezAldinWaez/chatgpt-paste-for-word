param([Parameter(Mandatory=$true)][string]$InstallDir)

$expected = [IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ChatGPTPasteForWord'))
$target = [IO.Path]::GetFullPath($InstallDir)
if (-not [string]::Equals($target, $expected, [StringComparison]::OrdinalIgnoreCase)) { exit 2 }

$uninstallKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord'
for ($attempt = 0; $attempt -lt 259200; $attempt++) {
    # If the user has reinstalled, the new installation must not be removed.
    $installed = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($uninstallKey)
    if ($installed) { $installed.Close(); break }
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10; continue }
    try {
        if (Test-Path -LiteralPath $target) {
            foreach ($item in Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue) {
                $installed = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($uninstallKey)
                if ($installed) { $installed.Close(); break }
                if ($item.PSIsContainer) { continue }
                try { Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop } catch { }
            }
            $installed = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($uninstallKey)
            if ($installed) { $installed.Close(); break }
            if (-not (Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue | Select-Object -First 1)) {
                try { Remove-Item -LiteralPath $target -Force -ErrorAction Stop } catch { }
            }
        }
        if (-not (Test-Path -LiteralPath $target)) { break }
    } catch { }
    Start-Sleep -Seconds 10
}
try { Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue } catch { }
