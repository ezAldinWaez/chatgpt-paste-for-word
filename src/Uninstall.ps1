param([string]$CleanupSource,[switch]$SkipCleanupLaunch)
$ErrorActionPreference = 'Stop'
$installDir = [IO.Path]::GetFullPath((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ChatGPTPasteForWord'))
$expected = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'ChatGPTPasteForWord'))
if (-not [string]::Equals($installDir, $expected, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected install path; no files were removed.' }

$appExe = Join-Path $installDir 'ChatGPTPasteHotkey.exe'
Get-Process -Name ChatGPTPasteHotkey -ErrorAction SilentlyContinue | ForEach-Object {
    try { if ([string]::Equals($_.Path, $appExe, [StringComparison]::OrdinalIgnoreCase)) { Stop-Process -Id $_.Id -Force -ErrorAction Stop } } catch { }
}

# Disable Word loading and startup before touching files. A locked ribbon DLL
# must not leave an active registration pointed at a half-removed directory.
$progId = 'ChatGPTPasteForWord.Ribbon'
$clsid = '{E1379E36-6276-4EAB-8A0E-5A30FD4C91A7}'
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $root.DeleteSubKeyTree("Software\Microsoft\Office\Word\Addins\$progId", $false)
        $root.DeleteSubKeyTree("Software\Classes\$progId", $false)
        $root.DeleteSubKeyTree("Software\Classes\CLSID\$clsid", $false)
        foreach ($recordId in @('{289E9AF1-4973-11D1-AE81-00A0C90F26F4}', '{289E9AF2-4973-11D1-AE81-00A0C90F26F4}')) {
            $recordRoot = "Software\Classes\Record\$recordId"
            $versions = $root.OpenSubKey($recordRoot)
            if ($versions) {
                $names = @($versions.GetSubKeyNames())
                $versions.Close()
                foreach ($version in $names) {
                    $path = "$recordRoot\$version"
                    $key = $root.OpenSubKey($path)
                    if ($key) {
                        $assembly = [string]$key.GetValue('Assembly', '')
                        $key.Close()
                        if ($assembly.StartsWith('ChatGPTPasteRibbon-', [StringComparison]::Ordinal)) { $root.DeleteSubKeyTree($path, $false) }
                    }
                }
            }
        }
    } finally { $root.Close() }
}
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
Remove-ItemProperty -Path $runKey -Name ChatGPTPasteForWord -ErrorAction SilentlyContinue
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ChatGPTPasteForWord'
Remove-Item -LiteralPath $uninstallKey -Recurse -Force -ErrorAction SilentlyContinue

# Copy the cleanup helper outside the install directory before removing it.
if (-not $CleanupSource) { $CleanupSource = Join-Path $installDir 'Cleanup-Word-Paste.ps1' }
$cleanupTemp = Join-Path $env:TEMP ('ChatGPTPasteForWord-cleanup-' + [Guid]::NewGuid().ToString('N') + '.ps1')
if (-not $SkipCleanupLaunch -and (Test-Path -LiteralPath $CleanupSource)) { Copy-Item -LiteralPath $CleanupSource -Destination $cleanupTemp -Force }
if (Test-Path -LiteralPath $installDir) {
    foreach ($item in Get-ChildItem -LiteralPath $installDir -Force -ErrorAction SilentlyContinue) {
        if ($item.PSIsContainer) { continue }
        try { Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop } catch { }
    }
    if (-not (Get-ChildItem -LiteralPath $installDir -Force -ErrorAction SilentlyContinue | Select-Object -First 1)) {
        try { Remove-Item -LiteralPath $installDir -Force -ErrorAction Stop } catch { }
    }
}
if (Test-Path -LiteralPath $installDir) {
    if ($SkipCleanupLaunch) {
        Write-Output 'ChatGPT Paste was removed from Word and startup. Close every Word window to finish deleting files currently in use.'
        exit 0
    }
    if (-not (Test-Path -LiteralPath $cleanupTemp)) { throw 'The add-in was disabled, but files still in use could not be scheduled for cleanup. Close Word and run setup again.' }
    $powershell = Join-Path $PSHOME 'powershell.exe'
    $arguments = @('-NoProfile','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',('"' + $cleanupTemp + '"'),'-InstallDir',('"' + $installDir + '"'))
    $cleanup = Start-Process -FilePath $powershell -ArgumentList $arguments -WindowStyle Hidden -PassThru
    if (-not $cleanup) { throw 'The add-in was disabled, but cleanup could not start. Close Word and run setup again.' }
    Write-Output 'ChatGPT Paste was removed from Word and startup. Close every Word window to finish deleting files currently in use.'
    exit 0
}
Remove-Item -LiteralPath $cleanupTemp -Force -ErrorAction SilentlyContinue
Write-Output 'ChatGPT Paste for Word was removed completely.'
