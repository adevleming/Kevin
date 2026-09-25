#Requires -Version 5.1
<#
    Tests for the form-driven request processing. Run:  pwsh ./tests/Test-LifecycleRequests.ps1
    Offline: Graph and Exchange Online calls go to in-memory mocks.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$fx = Join-Path $PSScriptRoot 'fixtures'
Import-Module (Join-Path $root 'PaycomLifecycle.psm1') -Force
Import-Module (Join-Path $root 'LifecycleRequests.psm1') -Force
. (Join-Path $PSScriptRoot 'TestHarness.ps1')

function Utc([string]$Value) { [DateTimeOffset]::Parse($Value, [Globalization.CultureInfo]::InvariantCulture).UtcDateTime }

$cfgPath = Join-Path $PSScriptRoot 'test-config.psd1'
$config = Import-PowerShellDataFile $cfgPath
$items = @(Get-Content (Join-Path $fx 'requests.json') -Raw | ConvertFrom-Json)
$requests = @($items | ForEach-Object { ConvertFrom-LifecycleListItem -Item $_ -Config $config })
$req = { param($id) $requests | Where-Object Id -eq $id }
$dir = @(Get-Content (Join-Path $fx 'directory-users.json') -Raw | ConvertFrom-Json)
Add-DirectoryUserDefaults $dir
$index = New-DirectoryIndex $dir
$now = Utc '2026-10-02T20:00:00Z'   # Friday 1pm in Spokane, before Dan's 6pm cutoff

Write-Host 'Dates and time zones'
It 'recovers the calendar date from a SharePoint date-only value' {
    Assert-Equal ([datetime]'2026-10-05') (ConvertFrom-ListDate '2026-10-05T07:00:00Z' 'Pacific Standard Time')
    Assert-Equal ([datetime]'2026-12-01') (ConvertFrom-ListDate '2026-12-01T08:00:00Z' 'Pacific Standard Time') 'standard time (UTC-8)'
    Assert-Equal ([datetime]'2026-10-05') (ConvertFrom-ListDate '2026-10-05T05:00:00Z' 'Central Standard Time')
    Assert-Equal ([datetime]'2026-10-05') (ConvertFrom-ListDate '2026-10-05' 'Pacific Standard Time')
    Assert-Equal ([datetime]'2026-10-05') (ConvertFrom-ListDate ([datetime]::SpecifyKind([datetime]'2026-10-05 07:00', 'Utc')) 'Pacific Standard Time')
    Assert-True ($null -eq (ConvertFrom-ListDate '' 'Pacific Standard Time'))
}
It 'computes the end-of-last-day cutoff in the site time zone' {
    Assert-Equal (Utc '2026-10-03T01:00:00Z') (Get-SiteLocalTimeUtc -Date ([datetime]'2026-10-02') -Time '18:00' -Site 'Spokane' -Config $config)
    Assert-Equal (Utc '2026-10-02T23:00:00Z') (Get-SiteLocalTimeUtc -Date ([datetime]'2026-10-02') -Time '18:00' -Site 'Amarillo (AMA)' -Config $config)
    Assert-Equal (Utc '2026-10-03T01:00:00Z') (Get-SiteLocalTimeUtc -Date ([datetime]'2026-10-02') -Time '18:00' -Site 'Unlisted site' -Config $config) 'falls back to the default time zone'
}

Write-Host 'Reading list items'
It 'normalises a new-hire item' {
    $r = & $req '1'
    Assert-Equal 'Maria Lopez' $r.DisplayName
    Assert-Equal ([datetime]'2026-10-05') $r.StartDate
    Assert-Equal 'Gina.Gomez@iac.aero' $r.RequesterEmail
    Assert-Equal 'Gina.Gomez@iac.aero' $r.ManagerEmail
    Assert-True (-not $r.DisableImmediately)
}
It 'uses the preferred name when given' { Assert-Equal 'Hank Hill' (& $req '9').DisplayName }
It 'resolves person columns through the lookup resolver first' {
    $item = [pscustomobject]@{ id = '50'; fields = [pscustomobject]@{ RequestType = 'Termination'; Status = 'Ready for IT'; EmployeeLookupId = '17'; EmployeeEmail = 'stale@iac.aero' } }
    $r = ConvertFrom-LifecycleListItem -Item $item -Config $config -ResolvePerson { param($id) if ($id -eq '17') { 'Dan.Dorsey@iac.aero' } }
    Assert-Equal 'Dan.Dorsey@iac.aero' $r.EmployeeEmail
}
It 'reads Graph hashtable items the same as JSON objects' {
    $item = @{ id = '51'; createdBy = @{ user = @{ email = 'x@iac.aero' } }; fields = @{ RequestType = 'New hire'; FirstName = ' Ann '; LastName = 'Lee'; DisableImmediately = 'Yes' } }
    $r = ConvertFrom-LifecycleListItem -Item $item -Config $config
    Assert-Equal 'Ann Lee' $r.DisplayName; Assert-Equal 'x@iac.aero' $r.RequesterEmail; Assert-True $r.DisableImmediately
}
It 'builds list columns from config' {
    $cols = Get-LifecycleRequestListColumns -Config $config
    $names = $cols | ForEach-Object { $_.name }
    foreach ($n in 'RequestType', 'Status', 'AccessType', 'Employee', 'Manager', 'LastDay', 'DisableImmediately', 'PaycomEmployeeId', 'ApprovedBy', 'ITLog') {
        Assert-True ($names -contains $n) "missing $n"
    }
    $site = $cols | Where-Object name -eq 'Site'
    Assert-Equal @('Spokane', 'Amarillo (AMA)') $site.choice.choices
    Assert-True ($cols | Where-Object name -eq 'Status').indexed
    Assert-Equal 'checkBoxes' ($cols | Where-Object name -eq 'Equipment').choice.displayAs
}

Write-Host 'Authorisation'
It 'holds a request marked Ready for IT without an approval' {
    $a = Test-LifecycleRequestAllowed -Request (& $req '6') -Config $config
    Assert-True (-not $a.Allowed); Assert-True ($a.Reason -match 'approval')
}
It 'lets an immediate termination through without waiting for approval' { Assert-True (Test-LifecycleRequestAllowed -Request (& $req '4') -Config $config).Allowed }
It 'holds a request last edited by someone other than HR or the flow' {
    $r = (& $req '1').PSObject.Copy(); $r.LastModifiedByEmail = 'Alice.Anders@iac.aero'
    $a = Test-LifecycleRequestAllowed -Request $r -Config $config
    Assert-True (-not $a.Allowed); Assert-True ($a.Reason -match 'trusted editor')
    Assert-True (Test-LifecycleRequestAllowed -Request (& $req '1') -Config $config).Allowed 'edited last by the flow account'
}
It 'rejects requesters outside the hiring groups' {
    $cfg = $config.Clone(); $cfg.Requests = $config.Requests.Clone(); $cfg.Requests.AuthorizedGroupIds = @('g-hiring')
    $a = Test-LifecycleRequestAllowed -Request (& $req '1') -Config $cfg -IsMemberOf { param($email, $groups) $false }
    Assert-True (-not $a.Allowed)
    Assert-True (Test-LifecycleRequestAllowed -Request (& $req '1') -Config $cfg -IsMemberOf { param($email, $groups) $email -eq 'Gina.Gomez@iac.aero' -and $groups -contains 'g-hiring' }).Allowed
}

Write-Host 'Decisions'
$expect = @{ '1' = 'CreateUser'; '2' = 'CreateContact'; '3' = 'Wait'; '4' = 'Offboard'; '5' = 'UpdateProfile'; '7' = 'RemoveContact'; '9' = 'CreateUser' }
foreach ($id in ($expect.Keys | Sort-Object { [int]$_ })) {
    It "request $id -> $($expect[$id])" { Assert-Equal $expect[$id] (Get-LifecycleRequestAction -Request (& $req $id) -NowUtc $now -Config $config).Action }
}
It 'schedules a voluntary termination for 6pm site time on the last day' {
    $d = Get-LifecycleRequestAction -Request (& $req '3') -NowUtc $now -Config $config
    Assert-Equal (Utc '2026-10-03T01:00:00Z') $d.RunAtUtc
    Assert-Equal 'Offboard' (Get-LifecycleRequestAction -Request (& $req '3') -NowUtc (Utc '2026-10-03T01:05:00Z') -Config $config).Action
}
It 'waits for a future change effective date' {
    Assert-Equal 'Wait' (Get-LifecycleRequestAction -Request (& $req '5') -NowUtc (Utc '2026-09-30T12:00:00Z') -Config $config).Action
}
It 'sends hybrid-directory hires to IT review' {
    $cfg = $config.Clone(); $cfg.DirectoryMode = 'Hybrid'
    Assert-Equal 'Review' (Get-LifecycleRequestAction -Request (& $req '1') -NowUtc $now -Config $cfg).Action
}
It 'treats a local-kind clock value correctly' {
    $local = (Utc '2026-10-03T01:05:00Z').ToLocalTime()
    Assert-Equal 'Offboard' (Get-LifecycleRequestAction -Request (& $req '3') -NowUtc $local -Config $config).Action
}

Write-Host 'Actions (mocked Graph and Exchange Online)'
$graph = New-Object Collections.Generic.List[object]
$exo = New-Object Collections.Generic.List[object]
$script:existingContacts = @('ada.painter@example.com')
Set-LifecycleGraphInvoker {
    param($Method, $Uri, $Body, $Out)
    $graph.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
    if ($Uri -match '/memberOf/') { return @{ value = @() } }
    if ($Method -eq 'POST' -and $Uri -eq '/v1.0/users') { return @{ id = 'new-id' } }
    return $null
}
Set-LifecycleExchangeInvoker {
    param($Command, $Parameters)
    $exo.Add([pscustomobject]@{ Command = $Command; Parameters = $Parameters })
    if ($Command -eq 'Get-MailContact') {
        if ($script:existingContacts -contains $Parameters.Identity) { return [pscustomobject]@{ Identity = $Parameters.Identity } }
        throw "Couldn't find object '$($Parameters.Identity)'."
    }
    return $null
}
$upns = New-Object Collections.Generic.List[string]
foreach ($u in $dir) { $upns.Add($u.userPrincipalName) }

It 'creates a full user with site and access groups, manager and no employee ID yet' {
    $graph.Clear()
    $o = Invoke-LifecycleRequestAction -Request (& $req '1') -Action CreateUser -Index $index -Config $config -ExistingUpns $upns
    Assert-Equal 'Completed' $o.Status ($o.Log -join ' | ')
    Assert-Equal 'Maria.Lopez@iac.aero' $o.Upn
    $create = $graph | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq '/v1.0/users' }
    Assert-Equal 'Amarillo (AMA)' $create.Body.officeLocation
    Assert-Equal 'Production Planner' $create.Body.jobTitle
    Assert-True (-not $create.Body.ContainsKey('employeeId')) 'no Paycom code on the request yet'
    Assert-Equal @('/v1.0/groups/g-full/members/$ref', '/v1.0/groups/g-ama/members/$ref') ($graph | Where-Object { $_.Uri -match '^/v1.0/groups/' } | ForEach-Object Uri)
    Assert-True ($graph | Where-Object { $_.Method -eq 'PUT' -and $_.Uri -eq '/v1.0/users/new-id/manager/$ref' })
    Assert-True ($upns -contains 'Maria.Lopez@iac.aero') 'UPN reserved for the rest of the run'
}
It 'holds a new hire whose Paycom code already has an account (rehire)' {
    $graph.Clear()
    $o = Invoke-LifecycleRequestAction -Request (& $req '9') -Action CreateUser -Index $index -Config $config -ExistingUpns $upns
    Assert-Equal 'Needs IT review' $o.Status
    Assert-True ($o.Log -match 'Hank.Hill@iac.aero')
    Assert-Equal 0 $graph.Count 'nothing created'
}
It 'does not create a second account when re-running a request that already made one' {
    $graph.Clear()
    $r = (& $req '1').PSObject.Copy(); $r.ITUpn = 'Alice.Anders@iac.aero'
    $o = Invoke-LifecycleRequestAction -Request $r -Action CreateUser -Index $index -Config $config -ExistingUpns $upns
    Assert-Equal 'Needs IT review' $o.Status; Assert-Equal 0 $graph.Count
}
It 'numbers the UPN and flags a same-name account' {
    $graph.Clear()
    $r = (& $req '1').PSObject.Copy(); $r.PreferredName = 'Alice'; $r.FirstName = 'Alice'; $r.LastName = 'Anders'; $r.DisplayName = 'Alice Anders'
    $o = Invoke-LifecycleRequestAction -Request $r -Action CreateUser -Index $index -Config $config -ExistingUpns $upns
    Assert-Equal 'Alice.Anders2@iac.aero' $o.Upn
    Assert-True ($o.Log -match '^NOTE: Existing account')
}
It 'creates a contact for a painter and adds it to the site list' {
    $exo.Clear()
    $o = Invoke-LifecycleRequestAction -Request (& $req '2') -Action CreateContact -Index $index -Config $config
    Assert-Equal 'Completed' $o.Status ($o.Log -join ' | ')
    $new = $exo | Where-Object Command -eq 'New-MailContact'
    Assert-Equal 'luis.garcia@example.com' $new.Parameters.ExternalEmailAddress
    Assert-Equal 'Luis Garcia' $new.Parameters.Name
    $set = $exo | Where-Object Command -eq 'Set-Contact'
    Assert-Equal 'Painter' $set.Parameters.Title; Assert-Equal 'Amarillo (AMA)' $set.Parameters.Office; Assert-Equal '806-555-0142' $set.Parameters.MobilePhone
    Assert-Equal 'E2002' ($exo | Where-Object Command -eq 'Set-MailContact').Parameters.CustomAttribute1
    Assert-Equal 'ama-floor@iac.aero' ($exo | Where-Object Command -eq 'Add-DistributionGroupMember').Parameters.Identity
}
It 'does not duplicate an existing contact' {
    $exo.Clear()
    $r = (& $req '2').PSObject.Copy(); $r.PersonalEmail = 'ada.painter@example.com'
    Invoke-LifecycleRequestAction -Request $r -Action CreateContact -Index $index -Config $config | Out-Null
    Assert-Equal 0 @($exo | Where-Object Command -eq 'New-MailContact').Count
}
It 'records a contact-only hire with no email without calling Exchange' {
    $exo.Clear()
    $r = (& $req '2').PSObject.Copy(); $r.PersonalEmail = $null
    $o = Invoke-LifecycleRequestAction -Request $r -Action CreateContact -Index $index -Config $config
    Assert-Equal 'Completed' $o.Status; Assert-Equal 0 $exo.Count
}
It 'offboards: blocks sign-in, revokes sessions, converts the mailbox and gives the delegate access' {
    $graph.Clear(); $exo.Clear()
    $o = Invoke-LifecycleRequestAction -Request (& $req '3') -Action Offboard -Index $index -Config $config
    Assert-Equal 'Completed' $o.Status ($o.Log -join ' | ')
    Assert-Equal $false ($graph | Where-Object Method -eq 'PATCH').Body.accountEnabled
    Assert-True ($graph | Where-Object Uri -eq '/v1.0/users/u-dan/revokeSignInSessions')
    Assert-Equal 'Shared' ($exo | Where-Object Command -eq 'Set-Mailbox').Parameters.Type
    $perm = $exo | Where-Object Command -eq 'Add-MailboxPermission'
    Assert-Equal 'Frank.Fisher@iac.aero' $perm.Parameters.User; Assert-Equal 'FullAccess' $perm.Parameters.AccessRights
}
It 'never offboards a protected account' {
    $graph.Clear()
    $ceo = [pscustomobject]@{ id = 'u-ceo'; displayName = 'Chief Exec'; givenName = 'Chief'; surname = 'Exec'; userPrincipalName = 'CEO@iac.aero'; mail = 'CEO@iac.aero'; employeeId = $null; accountEnabled = $true; userType = 'Member'; assignedLicenses = @() }
    Add-DirectoryUserDefaults @($ceo)
    $idx = New-DirectoryIndex (@($dir) + $ceo)
    $r = (& $req '3').PSObject.Copy(); $r.EmployeeEmail = 'ceo@iac.aero'
    $o = Invoke-LifecycleRequestAction -Request $r -Action Offboard -Index $idx -Config $config
    Assert-Equal 'Needs IT review' $o.Status; Assert-Equal 0 $graph.Count
}
It 'sends an offboard for an unknown account to review' {
    $r = (& $req '3').PSObject.Copy(); $r.EmployeeEmail = 'nobody@iac.aero'
    Assert-Equal 'Needs IT review' (Invoke-LifecycleRequestAction -Request $r -Action Offboard -Index $index -Config $config).Status
}
It 'removes a painter contact at the end of their last day' {
    $exo.Clear(); $script:existingContacts += 'old.painter@example.com'
    $o = Invoke-LifecycleRequestAction -Request (& $req '7') -Action RemoveContact -Index $index -Config $config
    Assert-Equal 'Completed' $o.Status
    Assert-Equal 'old.painter@example.com' ($exo | Where-Object Command -eq 'Remove-MailContact').Parameters.Identity
}
It 'updates title and manager for a change, and fills a missing employee ID' {
    $graph.Clear()
    $o = Invoke-LifecycleRequestAction -Request (& $req '5') -Action UpdateProfile -Index $index -Config $config
    Assert-Equal 'Completed' $o.Status ($o.Log -join ' | ')
    $patch = $graph | Where-Object Method -eq 'PATCH'
    Assert-Equal '/v1.0/users/u-carol' $patch.Uri
    Assert-Equal 'Controller' $patch.Body.jobTitle; Assert-Equal 'E1003' $patch.Body.employeeId
    Assert-True ($graph | Where-Object { $_.Method -eq 'PUT' -and $_.Uri -eq '/v1.0/users/u-carol/manager/$ref' })
}
It 'writes a completion notice without leaking markup from names' {
    $r = (& $req '1').PSObject.Copy(); $r.DisplayName = 'Maria <b>Lopez</b>'
    $n = New-LifecycleRequestNotice -Request $r -Outcome ([pscustomobject]@{ Status = 'Completed'; Upn = 'Maria.Lopez@iac.aero'; Log = @('OK: Created') }) -Action CreateUser
    Assert-True ($n.Body -match 'Maria &lt;b&gt;Lopez&lt;/b&gt;'); Assert-True ($n.Body -match 'No password is sent by email')
}

Write-Host 'Runner (offline)'
$state = Join-Path $PSScriptRoot '.test-state'
Remove-Item $state -Recurse -Force -ErrorAction SilentlyContinue
$runner = Join-Path $root 'Invoke-LifecycleRequests.ps1'
$reqJson = Join-Path $fx 'requests.json'; $dirJson = Join-Path $fx 'directory-users.json'
try {
    It 'dry run plans every open request and changes nothing' {
        $graph.Clear(); $exo.Clear()
        $rows = @(& $runner -ConfigPath $cfgPath -RequestsJsonPath $reqJson -DirectoryJsonPath $dirJson -NoEmail -NowUtc $now)
        Assert-Equal '1:CreateUser,2:CreateContact,3:Wait,4:Offboard,5:UpdateProfile,6:Review,7:RemoveContact,9:CreateUser' (($rows | ForEach-Object { "$($_.Id):$($_.Action)" }) -join ',')
        Assert-Equal 0 $graph.Count; Assert-Equal 0 $exo.Count
    }
    It 'apply run writes each outcome back to its list item' {
        $graph.Clear(); $exo.Clear(); $script:existingContacts = @('old.painter@example.com')
        $rows = @(& $runner -ConfigPath $cfgPath -RequestsJsonPath $reqJson -DirectoryJsonPath $dirJson -Apply -NoEmail -NowUtc $now)
        $writes = @{}
        foreach ($c in ($graph | Where-Object { $_.Method -eq 'PATCH' -and $_.Uri -match '/lists/list-1/items/(\d+)/fields$' })) {
            $null = $c.Uri -match '/items/(\d+)/fields$'; $writes[$Matches[1]] = $c.Body
        }
        Assert-Equal 'Completed' $writes['1'].Status; Assert-Equal 'Maria.Lopez@iac.aero' $writes['1'].ITUpn
        Assert-Equal 'Completed' $writes['2'].Status
        Assert-Equal 'Scheduled' $writes['3'].Status; Assert-True ($writes['3'].ITLog -match 'Runs at 2026-10-02 18:00 site time')
        Assert-Equal 'Completed' $writes['4'].Status
        Assert-Equal 'Needs IT review' $writes['6'].Status
        Assert-Equal 'Needs IT review' $writes['9'].Status
        Assert-True (-not $writes.ContainsKey('8')) 'completed requests are left alone'
        Assert-True ($writes['1'].ProcessedAt -eq '2026-10-02T20:00:00Z')
        Assert-True (Get-ChildItem (Join-Path $state 'request-logs') -Filter 'run-*.json')
    }
    It 'claims each request before acting so an interrupted run is not repeated' {
        $claims = @($graph | Where-Object { $_.Method -eq 'PATCH' -and $_.Uri -match '/items/1/fields$' })
        Assert-Equal 'In progress,Completed' (($claims | ForEach-Object { $_.Body.Status }) -join ',')
        Assert-Equal 'Maria.Lopez@iac.aero' $claims[0].Body.ITUpn 'planned account name recorded before creating'
        $stuck = @([pscustomobject]@{ id = '60'; createdBy = @{ user = @{ email = 'Gina.Gomez@iac.aero' } }
                fields = [pscustomobject]@{ RequestType = 'New hire'; Status = 'In progress'; AccessType = 'Full user'; FirstName = 'Half'; LastName = 'Done'; ApprovedBy = 'Shelbae (HR)' } })
        $tmp = Join-Path $state 'stuck.json'; $stuck | ConvertTo-Json -Depth 6 | Set-Content $tmp
        $graph.Clear()
        $rows = @(& $runner -ConfigPath $cfgPath -RequestsJsonPath $tmp -DirectoryJsonPath $dirJson -Apply -NoEmail -NowUtc $now)
        Assert-Equal 'Review' $rows[0].Action
        Assert-Equal 0 @($graph | Where-Object { $_.Uri -eq '/v1.0/users' }).Count 'no second account'
    }
    It 'keeps going when one request throws' {
        $bad = @(
            [pscustomobject]@{ id = '70'; createdBy = @{ user = @{ email = 'Gina.Gomez@iac.aero' } }
                fields = [pscustomobject]@{ RequestType = 'Termination'; Status = 'Ready for IT'; AccessType = 'Full user'; EmployeeEmail = 'Olga.Oldham@iac.aero'; Site = 'Mars Base'; LastDay = '2026-10-01T07:00:00Z'; ApprovedBy = 'Shelbae (HR)' } }
            [pscustomobject]@{ id = '71'; createdBy = @{ user = @{ email = 'Gina.Gomez@iac.aero' } }
                fields = [pscustomobject]@{ RequestType = 'Change'; Status = 'Ready for IT'; AccessType = 'Full user'; EmployeeEmail = 'Carol.Chen@iac.aero'; JobTitle = 'CFO'; ApprovedBy = 'Shelbae (HR)' } }
        )
        $cfgBad = Join-Path $state 'bad-tz.psd1'
        (Get-Content $cfgPath -Raw) -replace "DefaultTimeZone   = 'Pacific Standard Time'", "DefaultTimeZone   = 'Not A Real Zone'" -replace "StatePath     = './.test-state'", "StatePath     = '$($state -replace "'", "''")'" | Set-Content $cfgBad
        $tmp = Join-Path $state 'bad.json'; $bad | ConvertTo-Json -Depth 6 | Set-Content $tmp
        $graph.Clear()
        $rows = @(& $runner -ConfigPath $cfgBad -RequestsJsonPath $tmp -DirectoryJsonPath $dirJson -Apply -NoEmail -NowUtc $now -WarningAction SilentlyContinue)
        Assert-Equal 'Needs IT review' ($rows | Where-Object Id -eq '70').Status
        Assert-Equal 'Completed' ($rows | Where-Object Id -eq '71').Status
    }
    It 'caps terminations per run' {
        $graph.Clear()
        $many = @(foreach ($n in 1..7) {
                [pscustomobject]@{ id = "$(100 + $n)"; createdBy = @{ user = @{ email = 'Gina.Gomez@iac.aero' } }
                    fields = [pscustomobject]@{ RequestType = 'Termination'; Status = 'Ready for IT'; AccessType = 'Full user'; EmployeeEmail = 'Olga.Oldham@iac.aero'; DisableImmediately = $true } }
            })
        $tmp = Join-Path $state 'many.json'; $many | ConvertTo-Json -Depth 6 | Set-Content $tmp
        $rows = @(& $runner -ConfigPath $cfgPath -RequestsJsonPath $tmp -DirectoryJsonPath $dirJson -NoEmail -NowUtc $now)
        Assert-Equal 5 @($rows | Where-Object Action -eq 'Offboard').Count
        Assert-Equal 2 @($rows | Where-Object Action -eq 'Review').Count
    }
}
finally {
    Remove-Item $state -Recurse -Force -ErrorAction SilentlyContinue
    Set-LifecycleGraphInvoker $null; Set-LifecycleExchangeInvoker $null
}

Write-Host 'Weekly Paycom audit cross-check'
$pcfg = $config
$prev = @(Import-PaycomRoster -Path (Join-Path $fx 'roster-previous.csv') -Columns $pcfg.Roster.Columns -ActiveStatusPattern $pcfg.Roster.ActiveStatusPattern -AsOf ([datetime]'2026-09-24'))
$cur = @(Import-PaycomRoster -Path (Join-Path $fx 'roster-current.csv') -Columns $pcfg.Roster.Columns -ActiveStatusPattern $pcfg.Roster.ActiveStatusPattern -AsOf ([datetime]'2026-09-24'))
$diff = Compare-PaycomRoster -Previous $prev -Current $cur
$paycomIndex = (Compare-RosterToDirectory -Roster $cur -DirectoryUsers $dir -Scope $pcfg.Scope).Index
$check = Compare-RosterToRequests -Diff $diff -Requests $requests -Index $paycomIndex
It 'flags Paycom hires nobody filed a request for (accent-insensitive name match)' {
    Assert-Equal @('E1011', 'E1012') ($check.HiresWithoutRequest | ForEach-Object EmployeeId)
    Assert-True ($check.CoveredHireIds -contains 'E1010') 'José Núñez has a pending request'
    Assert-True ($check.CoveredHireIds -contains 'E1008') 'Hank matched by Paycom code'
}
It 'flags Paycom terminations nobody filed a request for, ignoring rejected ones' {
    # Dan and Erin's requests name their accounts, not their Paycom rows: matched through Entra.
    Assert-Equal 0 @($check.TerminationsWithoutRequest).Count
    $onlyRejected = @($requests | Where-Object { $_.Id -ne '4' })
    $c2 = Compare-RosterToRequests -Diff $diff -Requests $onlyRejected -Index $paycomIndex
    Assert-Equal 'E1005' ($c2.TerminationsWithoutRequest | ForEach-Object { $_.Employee.EmployeeId })
}
It 'the Paycom audit only raises tickets for the gaps' {
    $recon = Compare-RosterToDirectory -Roster $cur -DirectoryUsers $dir -Scope $pcfg.Scope
    $result = [pscustomobject]@{ Diff = $diff; Reconciliation = $recon; Plan = [pscustomobject]@{ Offboard = @(); Onboard = @() }; RequestCheck = $check }
    $tickets = @(New-LifecycleTickets -Result $result -TicketConfig $pcfg.Tickets)
    Assert-Equal @('[Onboarding] Alice Anders - starts 2026-09-28', '[Onboarding] Pat Price - starts 2026-09-28') ($tickets | Where-Object Type -eq 'Onboarding' | ForEach-Object Subject)
    Assert-Equal 0 @($tickets | Where-Object Type -eq 'Offboarding').Count
}

Complete-Tests
