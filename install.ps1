[CmdletBinding()]
param([ValidateSet('Auto','x86','x64')][string]$OutlookBitness = 'Auto')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OutlookDetection.ps1')
if ($OutlookBitness -eq 'Auto') { $OutlookBitness = Find-OutlookBitness }
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close classic Outlook before installing.' }
$dll = Join-Path $PSScriptRoot 'bin\AutoFrom.dll'
if (-not (Test-Path -LiteralPath $dll)) { throw 'Run build.ps1 first.' }
if ($OutlookBitness -eq 'x64' -and -not [Environment]::Is64BitOperatingSystem) { throw '64-bit Outlook requires 64-bit Windows.' }
$dest = Join-Path $env:LOCALAPPDATA 'AutoFrom'
New-Item -ItemType Directory -Path $dest -Force | Out-Null
Copy-Item -LiteralPath $dll -Destination (Join-Path $dest 'AutoFrom.dll') -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'uninstall.ps1') -Destination (Join-Path $dest 'uninstall.ps1') -Force
$configPath = Join-Path $dest 'rules.json'
if (-not (Test-Path -LiteralPath $configPath)) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'rules.json') -Destination $configPath }
$view = if ($OutlookBitness -eq 'x64') { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$clsid = '{948FBFCA-513F-4BC4-8DA1-B83085E45B38}'
try {
    $key = $base.CreateSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
    try {
        $key.SetValue('', 'mscoree.dll')
        $key.SetValue('ThreadingModel', 'Both')
        $key.SetValue('Class', 'AutoFrom.Connect')
        $key.SetValue('Assembly', [Reflection.AssemblyName]::GetAssemblyName($dll).FullName)
        $key.SetValue('RuntimeVersion', 'v4.0.30319')
        $key.SetValue('CodeBase', ([Uri](Join-Path $dest 'AutoFrom.dll')).AbsoluteUri)
    } finally { $key.Dispose() }
    $key = $base.CreateSubKey("Software\Classes\CLSID\$clsid\ProgId")
    try { $key.SetValue('', 'AutoFrom.Connect') } finally { $key.Dispose() }
    $key = $base.CreateSubKey('Software\Classes\AutoFrom.Connect\CLSID')
    try { $key.SetValue('', $clsid) } finally { $key.Dispose() }
    $key = $base.CreateSubKey('Software\Microsoft\Office\Outlook\Addins\AutoFrom.Connect')
    try {
        $key.SetValue('FriendlyName', 'AutoFrom - automatic sender rules')
        $key.SetValue('Description', 'Changes the From sender in classic Outlook using local rules.')
        $key.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
    } finally { $key.Dispose() }
    $key = $base.CreateSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall\AutoFrom')
    try {
        $shellPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $uninstallPath = Join-Path $dest 'uninstall.ps1'
        $key.SetValue('DisplayName', 'AutoFrom for classic Outlook (pilot)')
        $key.SetValue('DisplayVersion', [Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString())
        $key.SetValue('InstallLocation', $dest)
        $key.SetValue('UninstallString', ('"{0}" -NoExit -NoProfile -File "{1}"' -f $shellPath, $uninstallPath))
        $key.SetValue('QuietUninstallString', ('"{0}" -NoProfile -File "{1}"' -f $shellPath, $uninstallPath))
        $key.SetValue('NoModify', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        $key.SetValue('NoRepair', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
    } finally { $key.Dispose() }
} finally { $base.Dispose() }
Write-Host "Installed for the current Windows user ($OutlookBitness). Start classic Outlook and open AutoFrom > Settings."
Write-Host 'Configure a default sender or recipient rules in AutoFrom > Settings. Existing settings were preserved. No mailbox permissions or tenant settings were changed.'
