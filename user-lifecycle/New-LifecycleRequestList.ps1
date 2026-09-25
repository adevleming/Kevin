#Requires -Version 5.1
<#
.SYNOPSIS
    One-time setup: creates the "Employee Lifecycle Requests" SharePoint list with every
    column the form, the approval flow and the IT automation use.

.DESCRIPTION
    Creating a list needs more than the automation's day-to-day access. With Sites.Selected,
    raise the app's grant on the site to 'manage' for this one run, then set it back to
    'write' (reading and updating items only needs 'write'). Or run it with a separate
    admin app that has Sites.Manage.All.
    Prints the new list ID to paste into config.psd1 (Requests.ListId).

.EXAMPLE
    ./New-LifecycleRequestList.ps1 -ConfigPath ./config.psd1 -WhatIf   # show the columns only
    ./New-LifecycleRequestList.ps1 -ConfigPath ./config.psd1
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [string]$DisplayName = 'Employee Lifecycle Requests'
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'PaycomLifecycle.psm1')
Import-Module (Join-Path $PSScriptRoot 'LifecycleRequests.psm1') -Force

$config = Import-PowerShellDataFile -Path $ConfigPath
$columns = Get-LifecycleRequestListColumns -Config $config

if (-not $PSCmdlet.ShouldProcess("site $($config.Requests.SiteId)", "Create list '$DisplayName' with $($columns.Count) columns")) {
    $columns | ForEach-Object { '{0,-20} {1}' -f $_.name, $_.displayName }
    return
}

Connect-LifecycleGraph -Graph $config.Graph
$list = New-LifecycleRequestList -Config $config -SiteId $config.Requests.SiteId -DisplayName $DisplayName
Write-Host "Created list '$DisplayName' ($($list.Id))"
foreach ($w in $list.Warnings) { Write-Warning $w }
Write-Host ''
Write-Host "Next: set Requests.ListId = '$($list.Id)' in config.psd1, then follow docs/FORM-AND-FLOW.md to set up the form, permissions, Teams tab and flow."
