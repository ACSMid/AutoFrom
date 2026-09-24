[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RulesPath)
$ErrorActionPreference = 'Stop'
Add-Type -Path (Join-Path $PSScriptRoot 'bin\AutoFrom.dll')
$config = [AutoFrom.Config]::Load((Resolve-Path -LiteralPath $RulesPath).Path)
Write-Host "Valid: enabled=$($config.Enabled), $($config.Senders.Length) senders, $($config.Rules.Length) rules."
Write-Host 'This validates syntax and references only. Account availability and Exchange permissions require Outlook testing.'
