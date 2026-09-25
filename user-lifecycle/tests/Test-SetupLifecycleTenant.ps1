#Requires -Version 5.1
<#
    Tests Setup-LifecycleTenant.ps1 against a fake Microsoft Graph (no tenant needed):
      - -WhatIf makes read-only calls only
      - a real run creates only new objects, and every write targets something it created
      - an existing group / app with the same name stops the run before any write
      - a re-run with the manifest doesn't create anything twice
    Run:  pwsh ./tests/Test-SetupLifecycleTenant.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'TestHarness.ps1')

# A stand-in for Microsoft.Graph.Authentication that routes every call to $global:FakeGraph.
$fakeRoot = Join-Path $PSScriptRoot '.fake-modules'
$fakeMod = Join-Path $fakeRoot 'Microsoft.Graph.Authentication'
New-Item -ItemType Directory -Path $fakeMod -Force | Out-Null
@'
function Connect-MgGraph { param([string[]]$Scopes, [switch]$NoWelcome, $TenantId) $global:FakeGraph.Scopes = $Scopes }
function Get-MgContext { [pscustomobject]@{ Account = 'admin@iac.aero'; TenantId = 'tenant-1' } }
function Invoke-MgGraphRequest {
    param([string]$Method, [string]$Uri, $Body, [string]$ContentType, $Headers, [string]$OutputFilePath)
    $parsed = if ($Body -is [string]) { $Body | ConvertFrom-Json } else { $Body }
    $global:FakeGraph.Calls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $parsed })
    & $global:FakeGraph.Handler $Method $Uri $parsed
}
function New-SelfSignedCertificate {
    param($Subject, $CertStoreLocation, $KeyExportPolicy, $KeySpec, $KeyLength, $HashAlgorithm, $NotAfter, $Provider)
    $global:FakeGraph.CertKeySpec = $KeySpec
    [pscustomobject]@{ Thumbprint = 'THUMB123'; RawData = [byte[]](1, 2, 3); NotAfter = $NotAfter }
}
Export-ModuleMember -Function *
'@ | Set-Content (Join-Path $fakeMod 'Microsoft.Graph.Authentication.psm1')
$env:PSModulePath = $fakeRoot + [IO.Path]::PathSeparator + $env:PSModulePath

$state = Join-Path $PSScriptRoot '.test-state-setup'
$cfgPath = Join-Path $state 'config.psd1'
$setupScript = Join-Path $root 'Setup-LifecycleTenant.ps1'

function Reset-Fake {
    param([switch]$ExistingGroup, [switch]$ExistingApp)
    $global:FakeGraph = @{ Calls = New-Object Collections.Generic.List[object]; Scopes = $null; CertKeySpec = $null }
    $flags = @{ ExistingGroup = [bool]$ExistingGroup; ExistingApp = [bool]$ExistingApp }
    $global:FakeGraph.Handler = {
        param($Method, $Uri, $Body)
        switch -Regex ($Uri) {
            '^/v1\.0/users/(.+)\?' { $u = [uri]::UnescapeDataString($Matches[1]); return @{ id = "id-$u"; displayName = $u; userPrincipalName = $u; accountEnabled = $true } }
            '^/v1\.0/groups\?\$filter' { if ($flags.ExistingGroup) { return @{ value = @(@{ id = 'someone-elses-group'; displayName = 'Existing' }) } } return @{ value = @() } }
            '^/v1\.0/applications\?\$filter' { if ($flags.ExistingApp) { return @{ value = @(@{ id = 'someone-elses-app'; appId = 'x' }) } } return @{ value = @() } }
            "^/v1\.0/servicePrincipals\(appId='00000003" { return @{ appRoles = @('User.ReadWrite.All', 'User.RevokeSessions.All', 'GroupMember.ReadWrite.All', 'Sites.Selected', 'AuditLog.Read.All', 'Mail.Send' | ForEach-Object { @{ id = "role-$_"; value = $_ } }) } }
            "^/v1\.0/servicePrincipals\(appId='00000002" { return @{ appRoles = @(@{ id = 'role-exo'; value = 'Exchange.ManageAsApp' }) } }
            '^/v1\.0/groups$' { return @{ id = 'new-group' } }
            '^/v1\.0/groups/new-group/team$' { return $null }
            '^/v1\.0/groups/new-group/members/\$ref$' { return $null }
            '^/v1\.0/groups/new-group/sites/root' { return @{ id = 'new-site'; webUrl = 'https://leascorp.sharepoint.com/sites/HiringStaffingRequests' } }
            '^/v1\.0/sites/new-site/lists$' { return @{ id = 'new-list'; webUrl = 'https://leascorp.sharepoint.com/sites/HiringStaffingRequests/Lists/Employee Lifecycle Requests' } }
            '^/v1\.0/sites/new-site/lists/new-list/columns\?' { return @{ value = @(@{ id = 'col-title'; name = 'Title'; indexed = $false }, @{ id = 'col-status'; name = 'Status'; indexed = $false }) } }
            '^/v1\.0/sites/new-site/lists/new-list/columns/col-' { return $null }
            '^/v1\.0/sites/new-site/drive\?' { return @{ id = 'new-drive' } }
            '^/v1\.0/drives/new-drive/root/children$' { return @{ id = 'new-folder' } }
            '^/v1\.0/applications$' { return @{ id = 'new-app-object'; appId = 'new-app-id' } }
            '^/v1\.0/servicePrincipals$' { return @{ id = 'new-sp' } }
            '^/v1\.0/sites/new-site/permissions$' { return @{ id = 'new-perm' } }
            default { throw "Fake Graph: unexpected $Method $Uri" }
        }
    }.GetNewClosure()
}

function New-TestConfig {
    Remove-Item $state -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $state | Out-Null
    $cfg = Get-Content (Join-Path $root 'config.example.psd1') -Raw
    $cfg = $cfg -replace "StatePath     = './state'", "StatePath     = './'"
    # 25 people, to exercise the 20-per-request limit on group creation.
    $extra = (1..21 | ForEach-Object { "            'member$_@iac.aero'" }) -join "`n"
    $cfg = $cfg -replace "            # HR and management: add their addresses here, one per line", $extra
    Set-Content -Path $cfgPath -Value $cfg
}
$writes = { @($global:FakeGraph.Calls | Where-Object { $_.Method -ne 'GET' }) }

try {
    Write-Host 'Dry run'
    New-TestConfig; Reset-Fake
    It '-WhatIf makes read-only calls only and writes no manifest' {
        & $setupScript -ConfigPath $cfgPath -WhatIf 6>$null | Out-Null
        Assert-Equal 0 @(& $writes).Count ((& $writes | ForEach-Object { "$($_.Method) $($_.Uri)" }) -join '; ')
        Assert-True (-not (Test-Path (Join-Path $state 'tenant-setup.json')))
    }

    Write-Host 'Name collisions stop before any write'
    It 'an existing group with the same nickname stops the run' {
        New-TestConfig; Reset-Fake -ExistingGroup
        $threw = $false
        try { & $setupScript -ConfigPath $cfgPath 6>$null | Out-Null } catch { $threw = $_.Exception.Message -match 'already exists' }
        Assert-True $threw; Assert-Equal 0 @(& $writes).Count
    }
    It 'an existing app registration with the same name stops the run' {
        New-TestConfig; Reset-Fake -ExistingApp
        $threw = $false
        try { & $setupScript -ConfigPath $cfgPath 6>$null | Out-Null } catch { $threw = $_.Exception.Message -match 'already exists' }
        Assert-True $threw; Assert-Equal 0 @(& $writes).Count
    }

    Write-Host 'Real run'
    New-TestConfig; Reset-Fake
    & $setupScript -ConfigPath $cfgPath 6>$null | Out-Null
    $manifest = Get-Content (Join-Path $state 'tenant-setup.json') -Raw | ConvertFrom-Json
    It 'creates the group, team, list, app, service principal and site grant (and no roster folder on the team site)' {
        foreach ($k in 'groupId', 'siteId', 'listId', 'appId', 'spObjectId', 'sitePermissionId', 'certThumbprint') {
            Assert-True $manifest.$k "manifest missing $k"
        }
        Assert-True $manifest.teamCreated
        Assert-Equal 0 @($global:FakeGraph.Calls | Where-Object { $_.Uri -match '/drives/' }).Count 'the Paycom roster must not live where hiring managers can read it'
    }
    It 'every write targets an object this run created' {
        $allowed = '^/v1\.0/(groups|groups/new-group/(team|members/\$ref)|sites/new-site/(lists|lists/new-list/columns/col-(title|status)|permissions)|drives/new-drive/root/children|applications|servicePrincipals)$'
        $bad = @(& $writes | Where-Object { $_.Uri -notmatch $allowed })
        Assert-Equal 0 $bad.Count (($bad | ForEach-Object { "$($_.Method) $($_.Uri)" }) -join '; ')
        Assert-Equal 0 @(& $writes | Where-Object { $_.Method -eq 'DELETE' }).Count 'never deletes'
        foreach ($p in @(& $writes | Where-Object { $_.Method -eq 'PATCH' })) { Assert-True ($p.Uri -match '/lists/new-list/columns/') "PATCH outside the new list: $($p.Uri)" }
    }
    It 'creates the team as private with members unable to add channels, tabs or apps' {
        $g = ($global:FakeGraph.Calls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq '/v1.0/groups' }).Body
        Assert-Equal 'Private' $g.visibility; Assert-Equal 'HiringStaffingRequests' $g.mailNickname
        Assert-True (@($g.'owners@odata.bind').Count + @($g.'members@odata.bind').Count -le 20) 'at most 20 directory objects in the create call'
        $team = ($global:FakeGraph.Calls | Where-Object { $_.Method -eq 'PUT' }).Body
        Assert-Equal $false $team.memberSettings.allowCreateUpdateRemoveTabs
    }
    It 'adds everyone, including those beyond the first 20' {
        $g = ($global:FakeGraph.Calls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq '/v1.0/groups' }).Body
        $later = @($global:FakeGraph.Calls | Where-Object { $_.Uri -eq '/v1.0/groups/new-group/members/$ref' }).Count
        Assert-Equal 26 (@($g.'members@odata.bind').Count + $later) '4 owners + Brian + 21 more, all as members'
    }
    It 'requests the right permissions, without Mail.Send, and grants only write on the new site' {
        $app = ($global:FakeGraph.Calls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq '/v1.0/applications' }).Body
        $graphIds = @($app.requiredResourceAccess | Where-Object resourceAppId -eq '00000003-0000-0000-c000-000000000000').resourceAccess.id
        foreach ($r in 'User.ReadWrite.All', 'User.RevokeSessions.All', 'GroupMember.ReadWrite.All', 'Sites.Selected') { Assert-True ($graphIds -contains "role-$r") "missing $r" }
        Assert-True ($graphIds -notcontains 'role-Mail.Send') 'Mail.Send must not be requested'
        Assert-Equal 'role-exo' (@($app.requiredResourceAccess | Where-Object resourceAppId -eq '00000002-0000-0ff1-ce00-000000000000').resourceAccess.id)
        $perm = ($global:FakeGraph.Calls | Where-Object { $_.Uri -eq '/v1.0/sites/new-site/permissions' }).Body
        Assert-Equal 'write' (@($perm.roles) -join ',')
        Assert-Equal 'KeyExchange' $global:FakeGraph.CertKeySpec 'Exchange app-only needs a KeyExchange certificate'
    }
    It 'signs in with delegated scopes only' {
        Assert-True ($global:FakeGraph.Scopes -contains 'Sites.FullControl.All')
        Assert-True (-not ($global:FakeGraph.Scopes -match 'Directory\.ReadWrite|RoleManagement'))
    }
    It 're-running with the manifest creates nothing new' {
        Reset-Fake
        & $setupScript -ConfigPath $cfgPath 6>$null | Out-Null
        Assert-Equal 0 @(& $writes).Count ((& $writes | ForEach-Object { "$($_.Method) $($_.Uri)" }) -join '; ')
    }
}
finally {
    Remove-Item $state, $fakeRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Complete-Tests
