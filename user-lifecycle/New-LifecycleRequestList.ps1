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
$body = @{ displayName = $DisplayName; list = @{ template = 'genericList' }; columns = $columns }

if (-not $PSCmdlet.ShouldProcess("site $($config.Requests.SiteId)", "Create list '$DisplayName' with $($columns.Count) columns")) {
    $columns | ForEach-Object { '{0,-20} {1}' -f $_.name, $_.displayName }
    return
}

Connect-LifecycleGraph -Graph $config.Graph
$list = Invoke-LifecycleGraph -Method POST -Uri "/v1.0/sites/$($config.Requests.SiteId)/lists" -Body $body
$listId = Get-FieldValue $list 'id'
Write-Host "Created list '$DisplayName' ($listId)"

$colUri = "/v1.0/sites/$($config.Requests.SiteId)/lists/$listId/columns"
$cols = @(Get-FieldValue (Invoke-LifecycleGraph -Method GET -Uri "$colUri`?`$select=id,name,indexed,required") 'value')
$byName = @{}
foreach ($c in $cols) { $byName[(Get-FieldValue $c 'name')] = $c }

# The automation filters on Status; make sure the index actually took.
if ($byName.ContainsKey('Status') -and -not (Get-FieldValue $byName['Status'] 'indexed')) {
    try { Invoke-LifecycleGraph -Method PATCH -Uri "$colUri/$(Get-FieldValue $byName['Status'] 'id')" -Body @{ indexed = $true } | Out-Null; Write-Host 'Status column indexed.' }
    catch { Write-Warning "Couldn't index Status ($($_.Exception.Message)). Do it in List settings > Indexed columns." }
}

# The built-in Title column is required by default; the form doesn't use it (the flow fills it in).
if ($byName.ContainsKey('Title')) {
    try { Invoke-LifecycleGraph -Method PATCH -Uri "$colUri/$(Get-FieldValue $byName['Title'] 'id')" -Body @{ required = $false } | Out-Null; Write-Host 'Title column set to optional.' }
    catch { Write-Warning "Couldn't make Title optional ($($_.Exception.Message)). Do it in List settings > Title > Require = No." }
}

Write-Host ''
Write-Host "Next: set Requests.ListId = '$listId' in config.psd1, then follow docs/FORM-AND-FLOW.md to set up the form, permissions, Teams tab and flow."
