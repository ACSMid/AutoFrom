$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\OutlookDetection.ps1')
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('AutoFrom-PE-' + [Guid]::NewGuid().ToString('N') + '.exe')
try {
    foreach ($case in @(@(0x014C,'x86'), @(0x8664,'x64'), @(0xAA64,'reject'))) {
        $data = New-Object byte[] 134
        $data[0]=0x4D; $data[1]=0x5A; $data[60]=128
        $data[128]=0x50; $data[129]=0x45
        [BitConverter]::GetBytes([uint16]$case[0]).CopyTo($data,132)
        [IO.File]::WriteAllBytes($fixture,$data)
        $actual = $null
        try { $actual = Get-OutlookExecutableBitness $fixture } catch { if ($case[1] -ne 'reject') { throw }; $actual='reject' }
        if ($actual -ne $case[1]) { throw "Architecture test failed: $($case[1])" }
    }
    [IO.File]::WriteAllText($fixture,'invalid executable')
    $rejected=$false
    try { Get-OutlookExecutableBitness $fixture | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'Invalid executable was accepted.' }
    Write-Host '4 executable detection tests passed.'
} finally { if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture } }
