<#!
.SYNOPSIS
Install, verify, update or uninstall one shared skill directory for Codex, Pi and Claude Code.
#>
[CmdletBinding()]
param(
    [ValidateSet('Codex','Pi','Claude','All')][string]$Agent = 'All',
    [ValidateSet('User','Project')][string]$Scope = 'User',
    [ValidateSet('Install','Verify','Uninstall')][string]$Action = 'Install',
    [string]$ProjectRoot = (Get-Location).Path,
    [string]$Source = (Split-Path -Parent $PSScriptRoot),
    [switch]$UpdateLink,
    [switch]$CodexLegacy
)

$ErrorActionPreference = 'Stop'
$sourcePath = [System.IO.Path]::GetFullPath($Source)
if (-not (Test-Path -LiteralPath (Join-Path $sourcePath 'SKILL.md') -PathType Leaf)) {
    throw "Source is not an Agent Skill: $sourcePath"
}
$basePath = if ($Scope -eq 'User') { [Environment]::GetFolderPath('UserProfile') } else { [System.IO.Path]::GetFullPath($ProjectRoot) }
$targets = [System.Collections.Generic.List[string]]::new()
if ($Agent -in @('Codex','Pi','All')) {
    $targets.Add((Join-Path $basePath '.agents\skills\academic-docx-editor'))
}
if ($Agent -in @('Claude','All')) {
    $targets.Add((Join-Path $basePath '.claude\skills\academic-docx-editor'))
}
if ($CodexLegacy -and $Scope -eq 'User' -and $Agent -in @('Codex','All')) {
    $targets.Add((Join-Path $basePath '.codex\skills\academic-docx-editor'))
}

$failed = $false
foreach ($target in ($targets | Select-Object -Unique)) {
    $item = Get-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
    $isLink = $null -ne $item -and $item.LinkType -in @('Junction','SymbolicLink')
    $same = $false
    if ($isLink) {
        $linkTarget = [string](@($item.Target)[0])
        if ($linkTarget) {
            $same = [System.IO.Path]::GetFullPath($linkTarget).TrimEnd('\') -ieq $sourcePath.TrimEnd('\')
        }
    }
    $status = ''
    switch ($Action) {
        'Verify' {
            $status = if ($same) { 'verified' } elseif ($null -eq $item) { 'missing' } else { 'conflict' }
            if (-not $same) { $failed = $true }
        }
        'Install' {
            if ($same) { $status = 'already-installed' }
            elseif ($null -ne $item -and (-not $isLink -or -not $UpdateLink)) {
                $status = 'conflict'
                $failed = $true
            } else {
                if ($isLink) { Remove-Item -LiteralPath $target -Force }
                $parent = Split-Path -Parent $target
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
                New-Item -ItemType Junction -Path $target -Target $sourcePath | Out-Null
                $status = if ($isLink) { 'updated' } else { 'installed' }
            }
        }
        'Uninstall' {
            if ($same) {
                Remove-Item -LiteralPath $target -Force
                $status = 'uninstalled'
            } elseif ($null -eq $item) {
                $status = 'already-absent'
            } else {
                $status = 'conflict-preserved'
                $failed = $true
            }
        }
    }
    [pscustomobject]@{Agent=$Agent; Scope=$Scope; Action=$Action; Link=$target; Source=$sourcePath; Status=$status}
}
if ($failed) { exit 2 }
