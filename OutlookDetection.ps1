function Get-OutlookExecutableBitness {
    param([Parameter(Mandatory=$true)][string]$Path)
    $stream = [IO.File]::OpenRead($Path)
    $reader = New-Object IO.BinaryReader($stream)
    try {
        if ($stream.Length -lt 64 -or $reader.ReadUInt16() -ne 0x5A4D) { throw 'Invalid executable header.' }
        $stream.Position = 0x3C
        $offset = $reader.ReadInt32()
        if ($offset -lt 64 -or $offset -gt ($stream.Length - 6)) { throw 'Invalid PE header offset.' }
        $stream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x4550) { throw 'Invalid PE signature.' }
        switch ($reader.ReadUInt16()) {
            0x014C { return 'x86' }
            0x8664 { return 'x64' }
            default { throw 'Unsupported Outlook executable architecture.' }
        }
    } finally { $reader.Dispose(); $stream.Dispose() }
}

function Find-OutlookBitness {
    $candidates = New-Object 'System.Collections.Generic.List[string]'
    $views = @([Microsoft.Win32.RegistryView]::Registry32)
    if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
    foreach ($view in $views) {
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            try {
                $key = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE')
                if ($null -ne $key) {
                    try {
                        $value = [string]$key.GetValue('')
                        if ($value) { $candidates.Add([Environment]::ExpandEnvironmentVariables($value).Trim('"')) }
                    } finally { $key.Dispose() }
                }
                foreach ($version in @('16.0','15.0','14.0')) {
                    $key = $base.OpenSubKey("Software\Microsoft\Office\$version\Outlook\InstallRoot")
                    if ($null -ne $key) {
                        try {
                            $value = [string]$key.GetValue('Path')
                            if ($value) { $candidates.Add((Join-Path $value 'OUTLOOK.EXE')) }
                        } finally { $key.Dispose() }
                    }
                }
            } finally { $base.Dispose() }
        }
    }
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432) | Where-Object { $_ } | Select-Object -Unique) {
        foreach ($relative in @('Microsoft Office\root\Office16\OUTLOOK.EXE','Microsoft Office\Office16\OUTLOOK.EXE','Microsoft Office\Office15\OUTLOOK.EXE','Microsoft Office\Office14\OUTLOOK.EXE')) {
            $candidates.Add((Join-Path $root $relative))
        }
    }
    $results = @()
    foreach ($path in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $results += [pscustomobject]@{ Path=$path; Bitness=(Get-OutlookExecutableBitness -Path $path) }
        }
    }
    $architectures = @($results | Select-Object -ExpandProperty Bitness -Unique)
    if ($architectures.Count -eq 0) { throw 'Classic Outlook was not found. If installed in a custom location, verify File > Office Account > About Outlook and rerun with -OutlookBitness x86 or x64.' }
    if ($architectures.Count -ne 1) { throw 'Both 32-bit and 64-bit Outlook executables were found. Verify the Outlook version you use and rerun with -OutlookBitness x86 or x64.' }
    Write-Host "Detected classic Outlook $($architectures[0]): $($results[0].Path)"
    return $architectures[0]
}
