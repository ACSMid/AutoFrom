[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path -LiteralPath $compiler)) { throw '.NET Framework 4.8 is required.' }
$bin = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Path $bin -Force | Out-Null
$sources = @('Rules.cs','Controller.cs','OutlookPort.cs','Connect.cs','Ribbon.cs','SettingsStore.cs','SettingsForm.cs','AliasIdentity.cs') | ForEach-Object { Join-Path $PSScriptRoot "src\$_" }
& $compiler /nologo /target:library /platform:anycpu /optimize+ /warnaserror+ "/out:$bin\AutoFrom.dll" /reference:System.dll /reference:System.Core.dll /reference:System.Web.Extensions.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:Microsoft.CSharp.dll @sources
if ($LASTEXITCODE -ne 0) { throw 'Add-in compilation failed.' }
& $compiler /nologo /target:exe /platform:anycpu /optimize+ /warnaserror+ "/out:$bin\AutoFrom.Tests.exe" /reference:System.dll /reference:System.Core.dll /reference:System.Web.Extensions.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.dll "/reference:$bin\AutoFrom.dll" (Join-Path $PSScriptRoot 'tests\Tests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& "$bin\AutoFrom.Tests.exe" (Join-Path $PSScriptRoot 'rules.json') (Join-Path $PSScriptRoot 'rules.example.json')
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
Write-Host 'Build and tests passed. No Outlook or registry changes were made.'
