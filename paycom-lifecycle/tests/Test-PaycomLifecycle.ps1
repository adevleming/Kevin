#Requires -Version 5.1
<#
    Self-contained tests (no Pester needed). Run:  pwsh ./tests/Test-PaycomLifecycle.ps1
    Exits non-zero on failure. Nothing here talks to Microsoft Graph.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$fx = Join-Path $PSScriptRoot 'fixtures'
Import-Module (Join-Path $root 'PaycomLifecycle.psm1') -Force

$script:passed = 0; $script:failed = 0
function It {
    param([string]$Name, [scriptblock]$Test)
    try { & $Test; $script:passed++; Write-Host "  [pass] $Name" -ForegroundColor Green }
    catch { $script:failed++; Write-Host "  [FAIL] $Name`n         $($_.Exception.Message)" -ForegroundColor Red }
}
function Assert-Equal {
    param($Expected, $Actual, [string]$Because = '')
    $e = ($Expected | ForEach-Object { "$_" }) -join ','; $a = ($Actual | ForEach-Object { "$_" }) -join ','
    if ($e -ne $a) { throw "Expected [$e] but got [$a]. $Because" }
}
function Assert-True { param($Condition, [string]$Because = '') if (-not $Condition) { throw "Expected true. $Because" } }

$config = Import-PowerShellDataFile (Join-Path $PSScriptRoot 'test-config.psd1')
$asOf = [datetime]'2026-09-24'
$prev = @(Import-PaycomRoster -Path (Join-Path $fx 'roster-previous.csv') -Columns $config.Roster.Columns -ActiveStatusPattern $config.Roster.ActiveStatusPattern -AsOf $asOf)
$cur = @(Import-PaycomRoster -Path (Join-Path $fx 'roster-current.csv') -Columns $config.Roster.Columns -ActiveStatusPattern $config.Roster.ActiveStatusPattern -AsOf $asOf)
$dir = @(Get-Content (Join-Path $fx 'directory-users.json') -Raw | ConvertFrom-Json)

Write-Host 'Helpers'
It 'strips accents and punctuation from names' { Assert-Equal 'JoseOBrien-Nunez' (Get-AsciiName "José O'Brien-Núñez") }
It 'parses Paycom date formats' {
    Assert-Equal ([datetime]'2026-09-28') (ConvertTo-RosterDate '09/28/2026')
    Assert-Equal ([datetime]'2026-09-28') (ConvertTo-RosterDate '9/28/2026')
    Assert-Equal ([datetime]'2026-09-28') (ConvertTo-RosterDate '2026-09-28')
    Assert-True ($null -eq (ConvertTo-RosterDate ''))
}
It 'builds First.Last UPNs and numbers collisions' {
    Assert-Equal 'Jose.Nunez@iac.aero' (New-UpnCandidate -FirstName 'José' -LastName 'Núñez' -Domain 'iac.aero')
    Assert-Equal 'Alice.Anders2@iac.aero' (New-UpnCandidate -FirstName 'Alice' -LastName 'Anders' -Domain 'iac.aero' -ExistingUpns @('alice.anders@iac.aero'))
    Assert-Equal 'Alice.Anders3@iac.aero' (New-UpnCandidate -FirstName 'Alice' -LastName 'Anders' -Domain 'iac.aero' -ExistingUpns @('Alice.Anders@iac.aero', 'Alice.Anders2@iac.aero'))
}
It 'generates strong random passwords' {
    $a = New-RandomPassword; $b = New-RandomPassword
    Assert-True ($a.Length -ge 32 -and $a -ne $b)
}

Write-Host 'Roster import'
It 'reads every row with a nickname as preferred name' {
    Assert-Equal 10 $cur.Count
    Assert-Equal 'Bob Baker' ($cur | Where-Object EmployeeId -eq 'E1002').DisplayName
}
It 'treats status T as inactive' { Assert-True (-not ($cur | Where-Object EmployeeId -eq 'E1004').IsActive) }
It 'ignores an old term date on a rehire' { Assert-True ($cur | Where-Object EmployeeId -eq 'E1008').IsActive }
It 'fails clearly when a required column is missing' {
    $bad = Join-Path ([IO.Path]::GetTempPath()) "bad-$([guid]::NewGuid()).csv"
    "Id,First,Last`n1,A,B" | Set-Content $bad
    try { Import-PaycomRoster -Path $bad -Columns $config.Roster.Columns | Out-Null; throw 'no error' }
    catch { Assert-True ($_.Exception.Message -match "missing required column 'Employee_Code'") $_.Exception.Message }
    finally { Remove-Item $bad -Force }
}

Write-Host 'Roster diff'
$diff = Compare-PaycomRoster -Previous $prev -Current $cur
It 'finds new hires, including the rehire' { Assert-Equal @('E1011', 'E1008', 'E1010', 'E1012') ($diff.NewHires | ForEach-Object EmployeeId) }
It 'finds terminations: termed in Paycom and dropped off the report' {
    Assert-Equal @('E1004', 'E1005') ($diff.Terminations | ForEach-Object { $_.Employee.EmployeeId })
    Assert-Equal ([datetime]'2026-09-21') ($diff.Terminations | Where-Object { $_.Employee.EmployeeId -eq 'E1004' }).TermDate
    Assert-True (($diff.Terminations | Where-Object { $_.Employee.EmployeeId -eq 'E1005' }).Reason -match 'No longer on')
}
It 'finds job title changes' {
    Assert-Equal 'E1003' ($diff.Changes | ForEach-Object { $_.Employee.EmployeeId })
    Assert-Equal 'JobTitle' $diff.Changes[0].Changes.Field
}
It 'treats a run with no previous snapshot as the baseline' {
    $first = Compare-PaycomRoster -Previous $null -Current $cur
    Assert-True $first.FirstRun; Assert-Equal 0 @($first.NewHires).Count
}

Write-Host 'Directory reconciliation'
$recon = Compare-RosterToDirectory -Roster $cur -DirectoryUsers $dir -Scope $config.Scope
It 'flags only real orphans (not service, room, guest, Eirtech or disabled accounts)' {
    # Erin dropped off the roster this week, so her account is orphaned too.
    Assert-Equal @('Erin Estes', 'Mike Scanlon', 'Olga Oldham') ($recon.Orphaned | ForEach-Object displayName)
}
It 'flags termed-but-enabled accounts' { Assert-Equal 'Dan.Dorsey@iac.aero' ($recon.TermedButEnabled | ForEach-Object { $_.User.userPrincipalName }) }
It 'lists active employees without an account' { Assert-Equal @('E1011', 'E1002', 'E1010', 'E1012') ($recon.NoAccount | ForEach-Object EmployeeId) }
It 'does not match a new hire to a same-name account owned by another employee code' {
    Assert-True (@($recon.Matches | Where-Object { $_.Employee.EmployeeId -eq 'E1011' }).Count -eq 0)
}
It 'matches the rehire to their old disabled account by employee ID' {
    $m = $recon.Matches | Where-Object { $_.Employee.EmployeeId -eq 'E1008' }
    Assert-Equal 'EmployeeId' $m.MatchedBy; Assert-Equal 'u-hank' $m.User.id
}
It 'queues employee ID backfill for email and name matches' {
    Assert-Equal @('E1003:Email', 'E1006:Name') ($recon.IdBackfill | ForEach-Object { "$($_.Employee.EmployeeId):$($_.MatchedBy)" } | Sort-Object)
}
It 'matches a dropped-off leaver from the previous roster by name' {
    $erin = $prev | Where-Object EmployeeId -eq 'E1005'
    Assert-Equal 'u-erin' (Find-DirectoryMatch $erin $recon.Index).User.id
}
It 'applies onboarding eligibility rules' {
    Assert-True (Test-OnboardingEligible ($cur | Where-Object EmployeeId -eq 'E1010') $config.Onboarding.Eligibility)
    Assert-True (-not (Test-OnboardingEligible ($cur | Where-Object EmployeeId -eq 'E1012') $config.Onboarding.Eligibility))
}

Write-Host 'Safety checks'
It 'passes a normal week' { Assert-True (Test-RosterSafety -Previous $prev -Current $cur -Diff $diff -Safety $config.Safety).IsSafe }
It 'blocks a truncated export' {
    $trunc = @(Import-PaycomRoster -Path (Join-Path $fx 'roster-truncated.csv') -Columns $config.Roster.Columns -AsOf $asOf)
    $s = Test-RosterSafety -Previous $prev -Current $trunc -Diff (Compare-PaycomRoster -Previous $prev -Current $trunc) -Safety $config.Safety
    Assert-True (-not $s.IsSafe); Assert-Equal 3 $s.Problems.Count ($s.Problems -join ' | ')
}

Write-Host 'Graph actions (mocked)'
$calls = New-Object Collections.Generic.List[object]
Set-LifecycleGraphInvoker {
    param($Method, $Uri, $Body, $Out)
    $calls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
    if ($Uri -match '/memberOf/') {
        return @{ value = @(
                @{ id = 'g-lic'; displayName = 'License - Business Premium'; groupTypes = @(); onPremisesSyncEnabled = $null; assignedLicenses = @(@{ skuId = 'x' }) },
                @{ id = 'g-dyn'; displayName = 'All Staff (dynamic)'; groupTypes = @('DynamicMembership'); onPremisesSyncEnabled = $null; assignedLicenses = @() },
                @{ id = 'g-eng'; displayName = 'Engineering'; groupTypes = @('Unified'); onPremisesSyncEnabled = $null; assignedLicenses = @() }
            )
        }
    }
    if ($Method -eq 'POST' -and $Uri -eq '/v1.0/users') { return @{ id = 'new-user-id' } }
    return $null
}
It 'offboarding blocks sign-in, revokes sessions and keeps licence/dynamic groups' {
    $calls.Clear()
    $cfg = @{ DirectoryMode = 'Cloud'; Offboarding = @{ RemoveGroupMemberships = $true; KeepGroupIds = @(); AddToGroupId = 'g-off' } }
    $log = Invoke-LifecycleOffboarding -User ($dir | Where-Object id -eq 'u-dan') -Config $cfg
    Assert-True (-not ($log -match '^FAILED')) ($log -join ' | ')
    $patch = $calls | Where-Object Method -eq 'PATCH'
    Assert-Equal '/v1.0/users/u-dan' $patch.Uri; Assert-Equal $false $patch.Body.accountEnabled
    Assert-True ($calls | Where-Object Uri -eq '/v1.0/users/u-dan/revokeSignInSessions')
    Assert-Equal '/v1.0/groups/g-eng/members/u-dan/$ref' (($calls | Where-Object Method -eq 'DELETE').Uri)
    Assert-True ($calls | Where-Object { $_.Uri -eq '/v1.0/groups/g-off/members/$ref' })
}
It 'onboarding creates the account with employee ID, manager and department groups' {
    $calls.Clear()
    $cfg = @{ Onboarding = @{ UsageLocation = 'US'; CompanyName = 'IAC'; DefaultGroupIds = @('g-all'); DepartmentGroups = @{ '^Sales' = @('g-sales') } } }
    $e = $cur | Where-Object EmployeeId -eq 'E1010'
    $log = Invoke-LifecycleOnboarding -Employee $e -Upn 'Jose.Nunez@iac.aero' -Config $cfg -ManagerUser ($dir | Where-Object id -eq 'u-alice')
    $create = $calls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq '/v1.0/users' }
    Assert-Equal 'E1010' $create.Body.employeeId
    Assert-Equal 'José Núñez' $create.Body.displayName
    Assert-Equal '2026-09-28T00:00:00Z' $create.Body.employeeHireDate
    Assert-True $create.Body.passwordProfile.forceChangePasswordNextSignIn
    Assert-True ($calls | Where-Object { $_.Method -eq 'PUT' -and $_.Uri -eq '/v1.0/users/new-user-id/manager/$ref' })
    Assert-Equal @('/v1.0/groups/g-all/members/$ref', '/v1.0/groups/g-sales/members/$ref') ($calls | Where-Object { $_.Uri -match '^/v1.0/groups/' } | ForEach-Object Uri)
    Assert-True (-not (($log -join ' ') -match [regex]::Escape($create.Body.passwordProfile.password))) 'password must never be logged'
}
Set-LifecycleGraphInvoker $null

Write-Host 'End-to-end (offline, no email)'
$state = Join-Path $PSScriptRoot '.test-state'; $drop = Join-Path $PSScriptRoot '.test-drop'
Remove-Item $state, $drop -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $drop | Out-Null
$runner = Join-Path $root 'Invoke-PaycomLifecycle.ps1'
$cfgPath = Join-Path $PSScriptRoot 'test-config.psd1'
$dirJson = Join-Path $fx 'directory-users.json'
$run = { & $runner -ConfigPath $cfgPath -DirectoryJsonPath $dirJson -NoEmail -AsOf $asOf -WarningAction SilentlyContinue }
try {
    It 'first run sets the baseline' {
        Copy-Item (Join-Path $fx 'roster-previous.csv') (Join-Path $drop 'IT Current Employees 2026-09-17.csv')
        $r = & $run
        Assert-True $r.Diff.FirstRun
        Assert-Equal 1 @(Get-ChildItem (Join-Path $state 'snapshots')).Count
    }
    It 'skips a file that was already processed' {
        $r = & $run
        Assert-True ($null -eq $r)
        Assert-True (Get-ChildItem (Join-Path $state 'reports') -Filter 'stale-*.html')
    }
    It 'second run reports joiners/leavers and writes tickets' {
        Start-Sleep -Milliseconds 1100   # snapshot names are second-resolution
        Copy-Item (Join-Path $fx 'roster-current.csv') (Join-Path $drop 'IT Current Employees 2026-09-24.csv')
        $r = & $run
        Assert-Equal 'Report only' $r.Mode
        Assert-Equal 4 @($r.Diff.NewHires).Count
        Assert-Equal @('Dan.Dorsey@iac.aero', 'Erin.Estes@iac.aero') ($r.Plan.Offboard | ForEach-Object { $_.User.userPrincipalName })
        Assert-Equal @('Alice.Anders2@iac.aero', 'Jose.Nunez@iac.aero') ($r.Plan.Onboard | ForEach-Object Upn)
        Assert-Equal 'Onboarding,Onboarding,Onboarding,Onboarding,Offboarding,Offboarding,Change' (($r.Tickets | ForEach-Object Type) -join ',')
        $report = Get-ChildItem (Join-Path $state 'reports') -Filter 'lifecycle-*.html' | Sort-Object Name | Select-Object -Last 1 | Get-Content -Raw
        Assert-True ($report -match 'Olga\.Oldham@iac\.aero') 'orphan listed'
        Assert-Equal 1 ([regex]::Matches($report, 'Erin\.Estes@iac\.aero').Count) 'leaver listed once, not also as orphan'
        Assert-True ($report -notmatch 'etas\.ie') 'Eirtech out of scope'
        Assert-True ($report -match 'Jos&#233; N&#250;&#241;ez|José Núñez') 'accented names rendered'
        Assert-Equal 2 @(Get-ChildItem (Join-Path $state 'snapshots')).Count
    }
    It 'a truncated export is held: no tickets, baseline unchanged' {
        Start-Sleep -Milliseconds 1100
        Copy-Item (Join-Path $fx 'roster-truncated.csv') (Join-Path $drop 'IT Current Employees 2026-10-01.csv')
        $r = & $run
        Assert-True (-not $r.Safety.IsSafe)
        Assert-Equal 0 @($r.Tickets).Count
        Assert-Equal 2 @(Get-ChildItem (Join-Path $state 'snapshots')).Count
        $report = Get-ChildItem (Join-Path $state 'reports') -Filter 'lifecycle-*.html' | Sort-Object Name | Select-Object -Last 1 | Get-Content -Raw
        Assert-True ($report -match 'Safety check failed')
    }
}
finally {
    Remove-Item $state, $drop -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host "$script:passed passed, $script:failed failed" -ForegroundColor $(if ($script:failed) { 'Red' } else { 'Green' })
if ($script:failed) { exit 1 }
