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
$dir = @(Get-Content (Join-Path $fx 'directory-requests.json') -Raw | ConvertFrom-Json)
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
    Assert-Equal (Utc '2026-10-03T01:00:00Z') (Get-SiteLocalTimeUtc -Date ([datetime]'2026-10-02') -Time '18:00' -Site 'Spokane (GEG)' -Config $config)
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
    Assert-Equal 8 @($site.choice.choices).Count; Assert-Equal 'Amarillo (AMA)' $site.choice.choices[3]
    Assert-True ($cols | Where-Object name -eq 'HireOutcome').indexed
    foreach ($n in 'TerminationDate', 'RehireEligible', 'BadgeCollected', 'HrExitChecklist', 'HrNewHireChecklist', 'BenefitsStatus') { Assert-True ($names -contains $n) "missing $n" }
    Assert-True ($cols | Where-Object name -eq 'Status').indexed
    Assert-Equal 'checkBoxes' ($cols | Where-Object name -eq 'Equipment').choice.displayAs
}

Write-Host 'Authorisation'
It 'holds a request marked Ready for IT without an approval' {
    $a = Test-LifecycleRequestAllowed -Request (& $req '6') -Config $config -IsMemberOf { param($w, $g) $true }
    Assert-True (-not $a.Allowed); Assert-True ($a.Reason -match 'approval')
}
It 'lets an immediate termination through without waiting for approval' { Assert-True (Test-LifecycleRequestAllowed -Request (& $req '4') -Config $config -IsMemberOf { param($w, $g) $true }).Allowed }
It 'fails closed when no authorised group is configured' {
    $cfg = $config.Clone(); $cfg.Requests = $config.Requests.Clone(); $cfg.Requests.AuthorizedGroupIds = @()
    Assert-True (-not (Test-LifecycleRequestAllowed -Request (& $req '1') -Config $cfg -IsMemberOf { param($w, $g) $true }).Allowed)
}
It 'checks the requester by object ID when known, and only when first accepted' {
    $r = (& $req '1').PSObject.Copy(); $r.RequesterId = 'obj-gina'
    $seen = New-Object Collections.Generic.List[string]
    Test-LifecycleRequestAllowed -Request $r -Config $config -IsMemberOf { param($w, $g) $seen.Add($w); $true } | Out-Null
    Assert-Equal 'obj-gina' $seen[0]
    $sched = (& $req '3').PSObject.Copy(); $sched.Status = 'Scheduled'
    Assert-True (Test-LifecycleRequestAllowed -Request $sched -Config $config -IsMemberOf { param($w, $g) $false }).Allowed 'requester left the team after acceptance'
}
Write-Host 'Approval history (version-history provenance)'
$v = {
    param([string]$id, [string]$when, [string]$by, [hashtable]$fields)
    $base = @{ RequestType = 'Termination'; Status = 'Submitted'; EmployeeEmail = 'Dan.Dorsey@iac.aero'; LastDay = '2026-10-02T07:00:00Z'; BadgeCollected = $null }
    foreach ($k in $fields.Keys) { $base[$k] = $fields[$k] }
    [pscustomobject]@{ id = $id; lastModifiedDateTime = $when; lastModifiedBy = @{ user = @{ email = $by } }; fields = [pscustomobject]$base }
}
$created = & $v '1.0' '2026-09-28T15:00:00Z' 'frank.fisher@iac.aero' @{}
$pending = & $v '2.0' '2026-09-28T15:00:05Z' 'flows@iac.aero' @{ Status = 'Pending approval' }
$approved = & $v '3.0' '2026-09-28T16:00:00Z' 'flows@iac.aero' @{ Status = 'Ready for IT'; ApprovedBy = 'Shelbea Bean' }
It 'accepts an approval made by the flow' { Assert-True (Test-LifecycleRequestProvenance -Versions @($created, $pending, $approved) -Config $config).Allowed }
It 'rejects Ready for IT set by the requester' {
    $forged = & $v '3.0' '2026-09-28T16:00:00Z' 'frank.fisher@iac.aero' @{ Status = 'Ready for IT'; ApprovedBy = 'Shelbea Bean' }
    $r = Test-LifecycleRequestProvenance -Versions @($created, $pending, $forged) -Config $config
    Assert-True (-not $r.Allowed); Assert-True ($r.Reason -match 'frank.fisher')
}
It 'rejects an item created already marked Ready for IT' {
    $only = & $v '1.0' '2026-09-28T15:00:00Z' 'frank.fisher@iac.aero' @{ Status = 'Ready for IT' }
    Assert-True (-not (Test-LifecycleRequestProvenance -Versions @($only) -Config $config).Allowed)
}
It 'rejects a change to a critical field after approval by someone untrusted' {
    $moved = & $v '4.0' '2026-09-29T09:00:00Z' 'frank.fisher@iac.aero' @{ Status = 'Ready for IT'; LastDay = '2026-09-29T07:00:00Z' }
    $r = Test-LifecycleRequestProvenance -Versions @($created, $pending, $approved, $moved) -Config $config
    Assert-True (-not $r.Allowed); Assert-True ($r.Reason -match 'LastDay')
}
It 'allows the requester to update checklist fields after approval' {
    $badge = & $v '4.0' '2026-10-02T23:30:00Z' 'frank.fisher@iac.aero' @{ Status = 'Ready for IT'; BadgeCollected = 'Yes' }
    Assert-True (Test-LifecycleRequestProvenance -Versions @($created, $pending, $approved, $badge) -Config $config).Allowed
}
It 'allows the automation''s own status changes and HR edits' {
    $sched = [pscustomobject]@{ id = '4.0'; lastModifiedDateTime = '2026-09-28T16:10:00Z'; lastModifiedBy = @{ application = @{ displayName = 'IT Lifecycle Automation' } }; fields = [pscustomobject]@{ RequestType = 'Termination'; Status = 'Scheduled'; EmployeeEmail = 'Dan.Dorsey@iac.aero'; LastDay = '2026-10-02T07:00:00Z' } }
    $hr = & $v '5.0' '2026-09-29T10:00:00Z' 'shelbea.bean@iac.aero' @{ Status = 'Scheduled'; LastDay = '2026-10-03T07:00:00Z' }
    Assert-True (Test-LifecycleRequestProvenance -Versions @($created, $pending, $approved, $sched, $hr) -Config $config).Allowed
}
It 'uses the latest approval after a request is sent back and re-approved' {
    $moved = & $v '4.0' '2026-09-29T09:00:00Z' 'frank.fisher@iac.aero' @{ Status = 'Needs IT review'; LastDay = '2026-09-29T07:00:00Z' }
    $again = & $v '5.0' '2026-09-29T11:00:00Z' 'shelbea.bean@iac.aero' @{ Status = 'Ready for IT'; LastDay = '2026-09-29T07:00:00Z' }
    Assert-True (Test-LifecycleRequestProvenance -Versions @($approved, $moved, $created, $again, $pending) -Config $config).Allowed 'order-independent'
}
It 'fails closed without readable history' {
    Assert-True (-not (Test-LifecycleRequestProvenance -Versions @() -Config $config).Allowed)
    Assert-True (-not (Test-LifecycleRequestProvenance -Versions @([pscustomobject]@{ id = '1.0'; lastModifiedBy = @{} }) -Config $config).Allowed)
}

It 'rejects requesters outside the hiring groups' {
    $cfg = $config.Clone(); $cfg.Requests = $config.Requests.Clone(); $cfg.Requests.AuthorizedGroupIds = @('g-hiring')
    $a = Test-LifecycleRequestAllowed -Request (& $req '1') -Config $cfg -IsMemberOf { param($email, $groups) $false }
    Assert-True (-not $a.Allowed)
    Assert-True (Test-LifecycleRequestAllowed -Request (& $req '1') -Config $cfg -IsMemberOf { param($email, $groups) $email -eq 'Gina.Gomez@iac.aero' -and $groups -contains 'g-hiring' }).Allowed
}

Write-Host 'Decisions'
$expect = @{ '1' = 'CreateUser'; '2' = 'CreateContact'; '3' = 'Wait'; '4' = 'Offboard'; '5' = 'UpdateProfile'; '7' = 'RemoveContact'; '9' = 'CreateUser'
    '12' = 'ReverseHire'; '16' = 'CancelHire'; '17' = 'ReverseHire' }
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
    if ($Uri -match '/memberOf/') {
        if ($Uri -match 'u-nora') { return @{ value = @(@{ id = 'g-lic-basic'; displayName = 'License - Business Basic'; groupTypes = @(); onPremisesSyncEnabled = $null; assignedLicenses = @(@{ skuId = 'x' }) }) } }
        return @{ value = @() }
    }
    if ($Uri -match '^/v1\.0/users/(u-[a-z]+)\?\$select=signInActivity$') {
        if ($Matches[1] -eq 'u-sam') { return @{ signInActivity = @{ lastSuccessfulSignInDateTime = '2026-10-02T14:05:00Z' } } }
        return @{ signInActivity = @{ lastSuccessfulSignInDateTime = $null } }
    }
    if ($Method -eq 'POST' -and $Uri -eq '/v1.0/users') { return @{ id = 'new-id' } }
    if ($Method -eq 'GET' -and $Uri -match '^/v1\.0/users/(u-[a-z]+)/manager') {
        switch ($Matches[1]) { 'u-erin' { return @{ id = 'u-gina'; mail = 'Gina.Gomez@iac.aero'; userPrincipalName = 'Gina.Gomez@iac.aero' } } 'u-dan' { return @{ id = 'u-frank'; mail = 'Frank.Fisher@iac.aero'; userPrincipalName = 'Frank.Fisher@iac.aero' } } default { throw 'Request_ResourceNotFound' } }
    }
    if ($Method -eq 'PATCH' -and $Uri -eq '/v1.0/users/u-olga' -and $script:failBlock) { throw 'Insufficient privileges' }
    return $null
}
Set-LifecycleExchangeInvoker {
    param($Command, $Parameters)
    $exo.Add([pscustomobject]@{ Command = $Command; Parameters = $Parameters })
    if ($Command -eq 'Get-MailContact') {
        if (-not $Parameters.ContainsKey('Filter')) { throw 'Get-MailContact must be called with -Filter (an empty -Identity returns every contact).' }
        $null = $Parameters.Filter -match "ExternalEmailAddress -eq '(.+)'$"; $addr = $Matches[1]
        if ($addr -eq 'dup@example.com') { return @([pscustomobject]@{ Guid = 'g1' }, [pscustomobject]@{ Guid = 'g2' }) }
        if ($addr -eq 'handmade@example.com') { return [pscustomobject]@{ Guid = 'guid-handmade'; CustomAttribute2 = '' } }
        if ($script:existingContacts -contains $addr) { return [pscustomobject]@{ Guid = "guid-$addr"; ExternalEmailAddress = "SMTP:$addr"; CustomAttribute2 = 'Lifecycle request 2' } }
        return $null
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
It 'holds an immediate termination filed by someone who isn''t the manager or HR' {
    $graph.Clear()
    $r = (& $req '4').PSObject.Copy(); $r.RequesterEmail = 'Alice.Anders@iac.aero'
    $o = Invoke-LifecycleRequestAction -Request $r -Action Offboard -Index $index -Config $config
    Assert-Equal 'Needs IT review' $o.Status; Assert-True ($o.Log -match "isn't .*manager or HR")
    Assert-Equal 0 @($graph | Where-Object Method -ne 'GET').Count 'nothing changed'
}
It 'runs an immediate termination from the employee''s manager, but only hands the mailbox to the manager' {
    $graph.Clear(); $exo.Clear()
    $r = (& $req '4').PSObject.Copy(); $r.MailboxDelegateEmail = 'Alice.Anders@iac.aero'
    $o = Invoke-LifecycleRequestAction -Request $r -Action Offboard -Index $index -Config $config
    Assert-Equal 'Completed' $o.Status ($o.Log -join ' | ')
    Assert-True ($graph | Where-Object { $_.Method -eq 'PATCH' -and $_.Uri -eq '/v1.0/users/u-erin' })
    Assert-Equal 0 @($exo | Where-Object Command -eq 'Add-MailboxPermission').Count 'no mailbox for a non-manager on an unapproved immediate termination'
    Assert-True ($o.Log -match "wasn't granted automatically")
}
It 'leaves the mailbox alone when sign-in couldn''t be blocked' {
    $graph.Clear(); $exo.Clear(); $script:failBlock = $true
    try {
        $r = (& $req '3').PSObject.Copy(); $r.EmployeeEmail = 'Olga.Oldham@iac.aero'
        $o = Invoke-LifecycleRequestAction -Request $r -Action Offboard -Index $index -Config $config
        Assert-Equal 'Needs IT review' $o.Status
        Assert-Equal 0 @($exo | Where-Object { @('Set-Mailbox', 'Add-MailboxPermission') -contains $_.Command }).Count
    }
    finally { $script:failBlock = $false }
}
It 'refuses accounts outside the managed domains' {
    $graph.Clear()
    $r = (& $req '3').PSObject.Copy(); $r.EmployeeEmail = 'niamh.byrne@etas.ie'
    $o = Invoke-LifecycleRequestAction -Request $r -Action Offboard -Index $index -Config $config
    Assert-Equal 'Needs IT review' $o.Status; Assert-True ($o.Log -match "isn't in a domain")
    Assert-Equal 0 @($graph | Where-Object Method -ne 'GET').Count
}
It 'never removes a contact this automation didn''t create' {
    $exo.Clear()
    $r = (& $req '7').PSObject.Copy(); $r.PersonalEmail = 'handmade@example.com'
    Assert-Equal 'Needs IT review' (Invoke-LifecycleRequestAction -Request $r -Action RemoveContact -Index $index -Config $config).Status
    Assert-Equal 0 @($exo | Where-Object Command -eq 'Remove-MailContact').Count
}
It 'sees an account created earlier in the same run (no duplicate for a double submission)' {
    $graph.Clear()
    $idx = New-DirectoryIndex $dir
    $u2 = New-Object Collections.Generic.List[string]; foreach ($x in $dir) { $u2.Add($x.userPrincipalName) }
    $a = (& $req '1').PSObject.Copy(); $a.PaycomEmployeeId = 'E3001'
    Invoke-LifecycleRequestAction -Request $a -Action CreateUser -Index $idx -Config $config -ExistingUpns $u2 | Out-Null
    $o = Invoke-LifecycleRequestAction -Request $a -Action CreateUser -Index $idx -Config $config -ExistingUpns $u2
    Assert-Equal 'Needs IT review' $o.Status
    Assert-Equal 1 @($graph | Where-Object { $_.Method -eq 'POST' -and $_.Uri -eq '/v1.0/users' }).Count
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
    Assert-Equal 'guid-old.painter@example.com' ($exo | Where-Object Command -eq 'Remove-MailContact').Parameters.Identity 'removed by GUID after an exact match'
}
It 'never removes a contact identified only by name, or an ambiguous one' {
    $exo.Clear()
    $r = (& $req '7').PSObject.Copy(); $r.PersonalEmail = $null
    Assert-Equal 'Needs IT review' (Invoke-LifecycleRequestAction -Request $r -Action RemoveContact -Index $index -Config $config).Status
    $r.PersonalEmail = 'dup@example.com'
    Assert-Equal 'Needs IT review' (Invoke-LifecycleRequestAction -Request $r -Action RemoveContact -Index $index -Config $config).Status
    Assert-Equal 0 @($exo | Where-Object Command -eq 'Remove-MailContact').Count
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
It 'reverses a no-show: disables the account it created and frees the licence' {
    $graph.Clear()
    $o = Invoke-LifecycleRequestAction -Request (& $req '12') -Action ReverseHire -Index $index -Config $config
    Assert-Equal 'Reversed - did not start' $o.Status ($o.Log -join ' | ')
    Assert-Equal '/v1.0/users/u-nora' ($graph | Where-Object Method -eq 'PATCH').Uri
    Assert-True ($graph | Where-Object Uri -eq '/v1.0/users/u-nora/revokeSignInSessions')
    Assert-True ($graph | Where-Object { $_.Method -eq 'DELETE' -and $_.Uri -eq '/v1.0/groups/g-lic-basic/members/u-nora/$ref' }) 'licence group removed: no mailbox to keep'
    Assert-Equal 0 @($graph | Where-Object { $_.Method -eq 'DELETE' -and $_.Uri -eq '/v1.0/users/u-nora' }).Count 'not deleted by default'
}
It 'holds a no-show reversal if the account has actually been used' {
    $graph.Clear()
    $r = (& $req '12').PSObject.Copy(); $r.ITUpn = 'Sam.Starter@iac.aero'
    $o = Invoke-LifecycleRequestAction -Request $r -Action ReverseHire -Index $index -Config $config
    Assert-Equal 'Needs IT review' $o.Status; Assert-True ($o.Log -match 'signed in successfully')
    Assert-Equal 0 @($graph | Where-Object Method -ne 'GET').Count 'nothing changed'
}
It 'deletes the unused account only when configured' {
    $graph.Clear()
    $cfg = $config.Clone(); $cfg.Requests = $config.Requests.Clone(); $cfg.Requests.NoShow = @{ DeleteAccount = $true; CheckTime = '10:00' }
    Invoke-LifecycleRequestAction -Request (& $req '12') -Action ReverseHire -Index $index -Config $cfg | Out-Null
    Assert-True ($graph | Where-Object { $_.Method -eq 'DELETE' -and $_.Uri -eq '/v1.0/users/u-nora' })
}
It 'reverses a contact-only no-show by removing the contact' {
    $exo.Clear(); $script:existingContacts = @('gone.painter@example.com')
    $o = Invoke-LifecycleRequestAction -Request (& $req '17') -Action ReverseHire -Index $index -Config $config
    Assert-Equal 'Reversed - did not start' $o.Status
    Assert-True ($exo | Where-Object Command -eq 'Remove-MailContact')
}
It 'cancels a no-show before anything was set up' {
    Assert-Equal 'Cancelled' (Invoke-LifecycleRequestAction -Request (& $req '16') -Action CancelHire -Index $index -Config $config).Status
}
It 'writes the Employee Status email in HR''s format' {
    $n = New-LifecycleStatusNotice -Request (& $req '1') -Kind 'New Hire' -Config $config
    Assert-Equal 'AMA New Hire' $n.Subject
    foreach ($t in 'First Name', 'Last Name', 'Title', 'Effective Date', 'Location', 'Maria', 'Lopez', 'Production Planner', '10/05/2026', 'Amarillo (AMA)') { Assert-True ($n.Body -match [regex]::Escape($t)) "missing $t" }
    $dan = $dir | Where-Object id -eq 'u-dan'
    $t = New-LifecycleStatusNotice -Request (& $req '3') -Kind 'Term' -Config $config -User $dan
    Assert-Equal 'GEG Term' $t.Subject; Assert-True ($t.Body -match 'Dan') 'name from the directory for a picked employee'
}
It 'asks about start dates per site, in site time, after the check time' {
    $pending = @($requests | Where-Object { $_.RequestType -eq 'New hire' -and $_.Status -eq 'Completed' -and $_.HireOutcome -eq 'Pending start' })
    $b = @(Get-LifecycleStartCheckBatches -Requests $pending -NowUtc $now -Config $config)
    Assert-Equal 'Amarillo (AMA)' ($b | ForEach-Object Site) 'Spokane hire starts on the 5th'
    Assert-Equal @('13', '14') ($b[0].Requests | ForEach-Object Id)
    Assert-Equal 0 @(Get-LifecycleStartCheckBatches -Requests $pending -NowUtc (Utc '2026-10-02T14:00:00Z') -Config $config).Count 'not before 10:00 Central'
    $sent = @($pending | ForEach-Object { $x = $_.PSObject.Copy(); $x.StartCheckSentAt = '2026-10-02T15:00:00Z'; $x })
    Assert-Equal 0 @(Get-LifecycleStartCheckBatches -Requests $sent -NowUtc $now -Config $config).Count 'only asked once'
    $n = New-LifecycleStartCheckNotice -Site 'Amarillo (AMA)' -Requests $b[0].Requests -ListUrl 'https://x/Lists/R'
    Assert-True ($n.Subject -match '2 people'); Assert-True ($n.Body -match 'EditForm.aspx\?ID=14')
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
$reqJson = Join-Path $fx 'requests.json'; $dirJson = Join-Path $fx 'directory-requests.json'
try {
    It 'dry run plans every open request and changes nothing' {
        $graph.Clear(); $exo.Clear()
        $rows = @(& $runner -ConfigPath $cfgPath -RequestsJsonPath $reqJson -DirectoryJsonPath $dirJson -NoEmail -NowUtc $now)
        Assert-Equal '1:CreateUser,2:CreateContact,3:Wait,4:Offboard,5:UpdateProfile,6:Review,7:RemoveContact,9:CreateUser,12:ReverseHire,16:CancelHire,17:ReverseHire,13:AskIfStarted,14:AskIfStarted' (($rows | ForEach-Object { "$($_.Id):$($_.Action)" }) -join ',')
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
        Assert-Equal 'Reversed - did not start' $writes['12'].Status
        Assert-Equal 'Cancelled' $writes['16'].Status
        Assert-Equal 'Started' $writes['13'].HireOutcome 'signed in, so marked Started'
        Assert-True $writes['14'].StartCheckSentAt 'asked about the one who hasn''t signed in'
        Assert-True ($writes['1'].ProcessedAt -eq '2026-10-02T20:00:00Z')
        Assert-True (Get-ChildItem (Join-Path $state 'request-logs') -Filter 'run-*.json')
    }
    It 'sends the Employee Status email once, with the badge office for terminations' {
        $graph.Clear(); $exo.Clear(); $script:existingContacts = @('old.painter@example.com', 'gone.painter@example.com')
        $rows = @(& $runner -ConfigPath $cfgPath -RequestsJsonPath $reqJson -DirectoryJsonPath $dirJson -Apply -NowUtc $now -WarningAction SilentlyContinue)
        $mails = @($graph | Where-Object { $_.Uri -match '/sendMail$' })
        $subjects = @($mails | ForEach-Object { $_.Body.message.subject })
        Assert-True ($subjects -contains 'AMA New Hire') ($subjects -join ' | ')
        Assert-True ($subjects -contains 'GEG Term') 'scheduled termination still notifies payroll right away'
        $noShows = @($mails | Where-Object { $_.Body.message.subject -eq 'AMA No-show' })
        Assert-Equal 2 $noShows.Count 'Nora and Gus'
        Assert-True (@($noShows[0].Body.message.toRecipients | ForEach-Object { $_.emailAddress.address }) -contains 'shelbea.bean@iac.aero') 'HR told to reverse in Paycom'
        $check = @($mails | Where-Object { $_.Body.message.subject -match 'Did everyone start' })
        Assert-Equal 1 $check.Count
        Assert-True (@($check[0].Body.message.toRecipients | ForEach-Object { $_.emailAddress.address }) -contains 'diane.mendez@iac.aero') 'site orientation contact asked'
        $amaTerm = @($mails | Where-Object { $_.Body.message.subject -eq 'AMA Term' })
        Assert-Equal 1 $amaTerm.Count 'the Amarillo painter leaving (request 7)'
        Assert-True (@($amaTerm[0].Body.message.toRecipients | ForEach-Object { $_.emailAddress.address }) -contains 'badges@ama-airport.example') 'airport badge office told'
        $gegTerm = @($mails | Where-Object { $_.Body.message.subject -eq 'GEG Term' })
        Assert-True (@($gegTerm[0].Body.message.toRecipients | ForEach-Object { $_.emailAddress.address }) -notcontains 'badges@ama-airport.example') 'only the right site''s badge office'
        $item3 = @($graph | Where-Object { $_.Method -eq 'PATCH' -and $_.Uri -match '/items/3/fields$' })
        Assert-True ($item3 | Where-Object { $_.Body.NotifiedAt }) 'NotifiedAt recorded'
        $already = @($mails | Where-Object { $_.Body.message.subject -eq 'AMA New Hire' -and $_.Body.message.body.content -match 'Nora' })
        Assert-Equal 0 $already.Count 'already notified requests are not re-sent'
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
    It 'with approval-history checking on, runs the genuinely approved request and holds the forged one' {
        $ver = {
            param($id, $when, $by, $status)
            @{ id = $id; lastModifiedDateTime = $when; lastModifiedBy = @{ user = @{ email = $by } }
                fields = @{ RequestType = 'Change'; Status = $status; EmployeeEmail = 'Carol.Chen@iac.aero'; JobTitle = 'Controller' } }
        }
        $mk = {
            param($id, $approver)
            @{ id = $id; createdBy = @{ user = @{ email = 'Frank.Fisher@iac.aero' } }
                fields = @{ RequestType = 'Change'; Status = 'Ready for IT'; AccessType = 'Full user'; EmployeeEmail = 'Carol.Chen@iac.aero'; JobTitle = 'Controller'; ApprovedBy = 'Shelbea Bean' }
                versions = @((& $ver '1.0' '2026-09-30T10:00:00Z' 'frank.fisher@iac.aero' 'Submitted'), (& $ver '2.0' '2026-09-30T11:00:00Z' $approver 'Ready for IT')) }
        }
        $tmp = Join-Path $state 'history.json'
        @((& $mk '80' 'flows@iac.aero'), (& $mk '81' 'frank.fisher@iac.aero')) | ConvertTo-Json -Depth 8 | Set-Content $tmp
        $cfgHist = Join-Path $state 'history.psd1'
        (Get-Content $cfgPath -Raw) -replace 'VerifyApprovalHistory = \$false', 'VerifyApprovalHistory = $true' -replace "StatePath     = './.test-state'", "StatePath     = '$($state -replace "'", "''")'" | Set-Content $cfgHist
        $rows = @(& $runner -ConfigPath $cfgHist -RequestsJsonPath $tmp -DirectoryJsonPath $dirJson -NoEmail -NowUtc $now)
        Assert-Equal 'UpdateProfile' ($rows | Where-Object Id -eq '80').Action
        $forged = $rows | Where-Object Id -eq '81'
        Assert-Equal 'Review' $forged.Action; Assert-True ($forged.Reason -match 'trusted approver')
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
