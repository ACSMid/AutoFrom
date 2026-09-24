[CmdletBinding(SupportsShouldProcess=$true)]
param([switch]$RemoveSettings, [switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$views = @([Microsoft.Win32.RegistryView]::Registry32)
if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
$keys = @(
    'Software\Microsoft\Office\Outlook\Addins\AutoFrom.Connect',
    'Software\Classes\AutoFrom.Connect',
    'Software\Classes\CLSID\{948FBFCA-513F-4BC4-8DA1-B83085E45B38}',
    'Software\Microsoft\Windows\CurrentVersion\Uninstall\AutoFrom'
)
$running = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)
Write-Host "Checking AutoFrom for Windows user $env:USERNAME."
$found = 0
$machineWide = $false
foreach ($view in $views) {
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
        try {
            foreach ($path in $keys) {
                $key = $base.OpenSubKey($path)
                if ($null -ne $key) {
                    $key.Dispose()
                    Write-Host "Found: $hive / $view / $path"
                    if ($hive -eq [Microsoft.Win32.RegistryHive]::CurrentUser) { $found++ } else { $machineWide = $true }
                }
            }
        } finally { $base.Dispose() }
    }
}
if ($CheckOnly) {
    Write-Host "Current-user registrations found: $found. Outlook processes: $($running.Count). No changes made."
    return
}
if ($running.Count -gt 0) {
    throw "Outlook is still running (process IDs: $($running.Id -join ', ')). Save your drafts and use Outlook File > Exit. If it remains in the background, restart Windows, then run the uninstaller before opening Outlook. Nothing was removed."
}
$failures = New-Object 'System.Collections.Generic.List[string]'
foreach ($view in $views) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        foreach ($path in $keys) {
            try {
                if ($PSCmdlet.ShouldProcess("HKCU / $view / $path", 'Remove AutoFrom registration')) {
                    $base.DeleteSubKeyTree($path, $false)
                    $remaining = $base.OpenSubKey($path)
                    if ($null -ne $remaining) { $remaining.Dispose(); throw 'Registration still exists after removal.' }
                }
            } catch { $failures.Add("$view / $path : $($_.Exception.Message)") }
        }
    } finally { $base.Dispose() }
}

# Only remove named files under the verified installation folder; never delete recursively.
$localRoot = [IO.Path]::GetFullPath($env:LOCALAPPDATA).TrimEnd('\')
$installRoot = [IO.Path]::GetFullPath((Join-Path $localRoot 'AutoFrom'))
if ([IO.Path]::GetDirectoryName($installRoot) -ne $localRoot -or [IO.Path]::GetFileName($installRoot) -ne 'AutoFrom') {
    throw 'Unexpected installation path; file removal stopped.'
}
if (Test-Path -LiteralPath $installRoot) {
    $folder = Get-Item -LiteralPath $installRoot -Force
    if (($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        $failures.Add('The installation folder is a link/junction. File removal was skipped; verify its destination before removing files manually.')
    } else {
        $files = @('AutoFrom.dll', 'uninstall.ps1')
        if ($RemoveSettings) { $files += @('rules.json', 'rules.json.bak') }
        foreach ($name in $files) {
            $target = [IO.Path]::GetFullPath((Join-Path $installRoot $name))
            if ([IO.Path]::GetDirectoryName($target) -ne $installRoot) { throw 'File target escaped the installation folder.' }
            try {
                if ((Test-Path -LiteralPath $target) -and $PSCmdlet.ShouldProcess($target, 'Remove installed AutoFrom file')) {
                    Remove-Item -LiteralPath $target -Force
                    if (Test-Path -LiteralPath $target) { throw 'File still exists after removal.' }
                }
            } catch { $failures.Add("$name : $($_.Exception.Message)") }
        }
        if (@(Get-ChildItem -LiteralPath $installRoot -Force).Count -eq 0 -and $PSCmdlet.ShouldProcess($installRoot, 'Remove empty AutoFrom folder')) {
            Remove-Item -LiteralPath $installRoot
        }
    }
}
if ($machineWide) { $failures.Add('An AutoFrom registration also exists under HKLM. This per-user uninstaller did not modify it; its administrator/installer must remove it.') }
if ($failures.Count -gt 0) { throw ("Uninstall was incomplete:`r`n" + ($failures -join "`r`n")) }
if ($WhatIfPreference) { Write-Host 'Preview only. No uninstall changes were made.'; return }
Write-Host 'AutoFrom current-user registration and installed DLL were removed and verified.'
if (-not $RemoveSettings) { Write-Host "Settings were kept in $installRoot. Use -RemoveSettings to also remove rules.json and its backup." }
Write-Host 'The downloaded project folder was not removed. Restart Outlook to clear its cached ribbon tab.'
