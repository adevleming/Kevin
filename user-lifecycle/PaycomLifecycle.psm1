#Requires -Version 5.1
<#
    PaycomLifecycle - turns the weekly Paycom "IT Current Employees" push report into
    joiner / leaver / mover events, reconciles them against Entra ID, and (optionally)
    acts on them.

    Pure functions (roster import, diff, reconciliation, report building) have no
    external dependencies so they can be tested offline. Anything that talks to
    Microsoft Graph goes through Invoke-LifecycleGraph so it can be swapped out.
#>
Set-StrictMode -Version Latest

#region Helpers

function Get-ConfigValue {
    # Safe lookup of an optional setting, e.g. Get-ConfigValue $Config 'Offboarding.ProtectedUpns'.
    # Strict mode would otherwise throw when a key is left out of config.psd1.
    param($Object, [Parameter(Mandatory)][string]$Path)
    $v = $Object
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $v) { return $null }
        if ($v -is [System.Collections.IDictionary]) { $v = if ($v.Contains($part)) { $v[$part] } else { $null } }
        else { $prop = $v.PSObject.Properties[$part]; $v = if ($prop) { $prop.Value } else { $null } }
    }
    return $v
}

function Get-AsciiName {
    # "José O'Brien-Núñez" -> "JoseOBrien-Nunez". Keeps letters and hyphens only.
    param([string]$Value)
    if (-not $Value) { return '' }
    $decomposed = $Value.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $decomposed.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return ($sb.ToString() -replace '[^A-Za-z\-]', '')
}

function Get-NameKey {
    # Comparison key for fuzzy name matching: lowercase ASCII letters only.
    param([string]$Value)
    return ((Get-AsciiName $Value) -replace '-', '').ToLowerInvariant()
}

function ConvertTo-RosterDate {
    param([string]$Value)
    if (-not $Value -or -not $Value.Trim()) { return $null }
    $formats = @('M/d/yyyy', 'MM/dd/yyyy', 'M/d/yy', 'yyyy-MM-dd', 'yyyy-MM-ddTHH:mm:ss', 'M/d/yyyy h:mm:ss tt', 'M/d/yyyy H:mm')
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($Value.Trim(), [string[]]$formats, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.Date
    }
    return $null
}

function Get-EmailDomain {
    param([string]$Address)
    if ($Address -and $Address.Contains('@')) { return $Address.Split('@')[-1].ToLowerInvariant() }
    return ''
}

function ConvertTo-HtmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-RandomPassword {
    # Throwaway initial password. Never shown to anyone: techs issue a Temporary
    # Access Pass on day one instead of emailing passwords around.
    param([int]$Length = 32)
    $chars = 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*-_=+'.ToCharArray()
    $bytes = New-Object byte[] $Length
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    $pw = -join ($bytes | ForEach-Object { $chars[$_ % $chars.Length] })
    # Guarantee complexity classes regardless of the random draw.
    return 'Aa1!' + $pw
}

#endregion

#region Roster

function Import-PaycomRoster {
    <#
        Reads a Paycom report export (CSV) and normalises it using the column map in
        config. Only EmployeeId, FirstName and LastName are required; every other
        column is optional so the report can be extended over time.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Columns,
        [string]$ActiveStatusPattern = '^(a|active)',
        [datetime]$AsOf = (Get-Date).Date
    )

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { throw "Roster '$Path' contains no rows." }

    $headers = @($rows[0].PSObject.Properties.Name)
    foreach ($required in 'EmployeeId', 'FirstName', 'LastName') {
        $col = $Columns[$required]
        if (-not $col -or $headers -notcontains $col) {
            throw "Roster '$Path' is missing required column '$col' (mapped to $required). Columns found: $($headers -join ', ')"
        }
    }

    foreach ($row in $rows) {
        $value = @{}
        foreach ($field in $Columns.Keys) {
            $col = $Columns[$field]
            $value[$field] = if ($col -and $headers -contains $col) { ([string]$row.$col).Trim() } else { '' }
        }
        if (-not $value.EmployeeId) { continue }

        $hireDate = ConvertTo-RosterDate $value.HireDate
        $termDate = ConvertTo-RosterDate $value.TermDate
        # A rehire keeps their old termination date in Paycom; ignore it when it
        # predates the latest hire date.
        if ($termDate -and $hireDate -and $termDate -lt $hireDate) { $termDate = $null }

        $statusActive = (-not $value.Status) -or ($value.Status -match $ActiveStatusPattern)
        $termActive = (-not $termDate) -or ($termDate -gt $AsOf)

        $first = if ($value.ContainsKey('PreferredName') -and $value.PreferredName) { $value.PreferredName } else { $value.FirstName }

        [pscustomobject]@{
            EmployeeId   = $value.EmployeeId
            FirstName    = $value.FirstName
            PreferredName = $first
            LastName     = $value.LastName
            DisplayName  = "$first $($value.LastName)".Trim()
            Email        = $value.Email
            Department   = $value.Department
            JobTitle     = $value.JobTitle
            Manager      = $value.Manager
            ManagerEmail = $value.ManagerEmail
            Location     = $value.Location
            Status       = $value.Status
            HireDate     = $hireDate
            TermDate     = $termDate
            IsActive     = [bool]($statusActive -and $termActive)
        }
    }
}

function ConvertTo-RosterIndex {
    # EmployeeId -> row. On duplicate IDs (rehires can appear twice) the active row wins.
    param([object[]]$Roster)
    $index = @{}
    foreach ($e in @($Roster)) {
        if (-not $index.ContainsKey($e.EmployeeId) -or ($e.IsActive -and -not $index[$e.EmployeeId].IsActive)) {
            $index[$e.EmployeeId] = $e
        }
    }
    return $index
}

function Compare-PaycomRoster {
    <#
        Diffs two roster snapshots keyed on Paycom employee code.
          NewHires     - active now, not active last time (includes rehires)
          Terminations - active last time, not active now (termed in Paycom or dropped off the report)
          Changes      - active in both, with a tracked field changed
    #>
    [CmdletBinding()]
    param(
        [object[]]$Previous,
        [Parameter(Mandatory)][object[]]$Current,
        [string[]]$TrackedFields = @('LastName', 'PreferredName', 'Email', 'Department', 'JobTitle', 'Manager', 'Location')
    )

    $cur = ConvertTo-RosterIndex $Current
    $newHires = New-Object Collections.Generic.List[object]
    $terms = New-Object Collections.Generic.List[object]
    $changes = New-Object Collections.Generic.List[object]

    if (-not $Previous) {
        return [pscustomobject]@{ FirstRun = $true; NewHires = @(); Terminations = @(); Changes = @() }
    }
    $prev = ConvertTo-RosterIndex $Previous

    foreach ($id in $cur.Keys) {
        $c = $cur[$id]
        if (-not $c.IsActive) { continue }
        if (-not $prev.ContainsKey($id) -or -not $prev[$id].IsActive) {
            $newHires.Add($c)
            continue
        }
        $p = $prev[$id]
        $diff = @(foreach ($f in $TrackedFields) {
                if ([string]$p.$f -ne [string]$c.$f) { [pscustomobject]@{ Field = $f; Old = $p.$f; New = $c.$f } }
            })
        if ($diff.Count) { $changes.Add([pscustomobject]@{ Employee = $c; Changes = $diff }) }
    }

    foreach ($id in $prev.Keys) {
        $p = $prev[$id]
        if (-not $p.IsActive) { continue }
        if ($cur.ContainsKey($id) -and $cur[$id].IsActive) { continue }
        $c = if ($cur.ContainsKey($id)) { $cur[$id] } else { $null }
        $terms.Add([pscustomobject]@{
                Employee = $p
                TermDate = if ($c) { $c.TermDate } else { $null }
                Reason   = if ($c) { "Status '$($c.Status)' in Paycom" } else { 'No longer on the Current Employees report' }
            })
    }

    [pscustomobject]@{
        FirstRun     = $false
        NewHires     = @($newHires | Sort-Object LastName, PreferredName)
        Terminations = @($terms | Sort-Object { $_.Employee.LastName })
        Changes      = @($changes | Sort-Object { $_.Employee.LastName })
    }
}

#endregion

#region Directory reconciliation

function Test-DirectoryUserInScope {
    <#
        Only accounts that Paycom is the source of truth for should be reconciled:
        enabled members in the configured domains, minus service / shared / admin
        accounts. Without this, every Eirtech (etas.ie) user would show up as orphaned.
    #>
    param($User, [hashtable]$Scope)
    if (-not $User.accountEnabled) { return $false }
    if ($User.PSObject.Properties['userType'] -and $User.userType -and $User.userType -ne 'Member') { return $false }
    $upn = [string]$User.userPrincipalName
    if ($upn -match '#EXT#') { return $false }
    if ((Get-ConfigValue $Scope 'Domains') -and (@((Get-ConfigValue $Scope 'Domains')) -notcontains (Get-EmailDomain $upn))) { return $false }
    if ((Get-ConfigValue $Scope 'ExcludeUpns') -and (@((Get-ConfigValue $Scope 'ExcludeUpns') | ForEach-Object { $_.ToLowerInvariant() }) -contains $upn.ToLowerInvariant())) { return $false }
    foreach ($pattern in @((Get-ConfigValue $Scope 'ExcludePatterns'))) {
        if ($pattern -and ($upn -match $pattern -or [string]$User.displayName -match $pattern)) { return $false }
    }
    if ((Get-ConfigValue $Scope 'RequireLicense') -and -not @($User.assignedLicenses).Count) { return $false }
    return $true
}

$script:DirectoryProperties = @('id', 'displayName', 'givenName', 'surname', 'userPrincipalName', 'mail', 'employeeId', 'accountEnabled',
    'userType', 'assignedLicenses', 'department', 'jobTitle', 'onPremisesSyncEnabled', 'signInActivity')

function Add-DirectoryUserDefaults {
    # Graph omits properties that weren't selected (and exports may be partial); add them as
    # $null so strict-mode property access works everywhere downstream.
    param([object[]]$DirectoryUsers)
    foreach ($u in @($DirectoryUsers)) {
        foreach ($p in $script:DirectoryProperties) {
            if (-not $u.PSObject.Properties[$p]) { Add-Member -InputObject $u -NotePropertyName $p -NotePropertyValue $null }
        }
    }
}

function New-DirectoryIndex {
    param([object[]]$DirectoryUsers)
    $byEmpId = @{}; $byEmail = @{}; $byName = @{}
    $addName = {
        param($key, $u)
        if (-not $key) { return }
        if (-not $byName.ContainsKey($key)) { $byName[$key] = New-Object Collections.Generic.List[object] }
        if (-not $byName[$key].Contains($u)) { $byName[$key].Add($u) }
    }
    foreach ($u in @($DirectoryUsers)) {
        if ($u.employeeId) { $byEmpId[[string]$u.employeeId] = $u }
        foreach ($addr in @($u.mail, $u.userPrincipalName)) {
            if ($addr) { $byEmail[([string]$addr).ToLowerInvariant()] = $u }
        }
        & $addName (Get-NameKey "$($u.givenName)$($u.surname)") $u
        & $addName (Get-NameKey $u.displayName) $u
        & $addName (Get-NameKey ([string]$u.userPrincipalName).Split('@')[0]) $u
    }
    [pscustomobject]@{ ByEmployeeId = $byEmpId; ByEmail = $byEmail; ByName = $byName }
}

function Find-DirectoryMatch {
    <#
        Match order, strongest first: Entra employeeId = Paycom employee code, then
        work email, then an unambiguous name match. Name matches are reported so
        the employeeId can be stamped and future runs match exactly.
    #>
    param($Employee, $Index)
    if ($Index.ByEmployeeId.ContainsKey($Employee.EmployeeId)) {
        return [pscustomobject]@{ User = $Index.ByEmployeeId[$Employee.EmployeeId]; MatchedBy = 'EmployeeId' }
    }
    if ($Employee.Email -and $Index.ByEmail.ContainsKey($Employee.Email.ToLowerInvariant())) {
        return [pscustomobject]@{ User = $Index.ByEmail[$Employee.Email.ToLowerInvariant()]; MatchedBy = 'Email' }
    }
    $keys = @("$($Employee.PreferredName)$($Employee.LastName)", "$($Employee.FirstName)$($Employee.LastName)") |
        ForEach-Object { Get-NameKey $_ } | Select-Object -Unique
    foreach ($k in $keys) {
        if ($Index.ByName.ContainsKey($k)) {
            # Skip accounts already bound to a different employee code.
            $candidates = @($Index.ByName[$k] | Where-Object { -not $_.employeeId -or $_.employeeId -eq $Employee.EmployeeId })
            if ($candidates.Count -eq 1) { return [pscustomobject]@{ User = $candidates[0]; MatchedBy = 'Name' } }
        }
    }
    return $null
}

function Compare-RosterToDirectory {
    <#
        Reconciles the full current roster against Entra ID.
          NoAccount       - active employee, no matching account
          TermedButEnabled- inactive in Paycom, account still enabled
          Orphaned        - enabled in-scope account with no active employee behind it
          IdBackfill      - matched by email/name and employeeId not set in Entra yet
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Roster,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DirectoryUsers,
        [hashtable]$Scope = @{}
    )
    Add-DirectoryUserDefaults $DirectoryUsers
    $index = New-DirectoryIndex $DirectoryUsers
    $matched = @{}   # directory id -> $true when an ACTIVE employee owns it
    $matchList = New-Object Collections.Generic.List[object]
    $noAccount = New-Object Collections.Generic.List[object]
    $termedEnabled = New-Object Collections.Generic.List[object]
    $backfill = New-Object Collections.Generic.List[object]

    foreach ($e in (ConvertTo-RosterIndex $Roster).Values) {
        $m = Find-DirectoryMatch $e $index
        if (-not $m) {
            if ($e.IsActive) { $noAccount.Add($e) }
            continue
        }
        $matchList.Add([pscustomobject]@{ Employee = $e; User = $m.User; MatchedBy = $m.MatchedBy })
        if ($e.IsActive) {
            $matched[$m.User.id] = $true
            if ($m.MatchedBy -ne 'EmployeeId' -and -not $m.User.employeeId) {
                $backfill.Add([pscustomobject]@{ Employee = $e; User = $m.User; MatchedBy = $m.MatchedBy })
            }
        }
        elseif ($m.User.accountEnabled) {
            $termedEnabled.Add([pscustomobject]@{ Employee = $e; User = $m.User; MatchedBy = $m.MatchedBy })
        }
    }

    $termedIds = @{}
    foreach ($t in $termedEnabled) { $termedIds[$t.User.id] = $true }
    $orphaned = @($DirectoryUsers | Where-Object {
            -not $matched.ContainsKey($_.id) -and -not $termedIds.ContainsKey($_.id) -and (Test-DirectoryUserInScope $_ $Scope)
        } | Sort-Object displayName)

    [pscustomobject]@{
        Index            = $index
        Matches          = $matchList.ToArray()
        NoAccount        = @($noAccount | Sort-Object LastName, PreferredName)
        TermedButEnabled = $termedEnabled.ToArray()
        Orphaned         = $orphaned
        IdBackfill       = $backfill.ToArray()
    }
}

function New-UpnCandidate {
    # Follows the existing First.Last@domain convention, adding a number on collision.
    param(
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [Parameter(Mandatory)][string]$Domain,
        [string[]]$ExistingUpns = @()
    )
    $taken = @{}
    foreach ($u in $ExistingUpns) { if ($u) { $taken[$u.ToLowerInvariant()] = $true } }
    $base = '{0}.{1}' -f (Get-AsciiName $FirstName), (Get-AsciiName $LastName)
    $base = $base.Trim('.', '-')
    $candidate = "$base@$Domain"
    $n = 2
    while ($taken.ContainsKey($candidate.ToLowerInvariant())) {
        $candidate = "$base$n@$Domain"; $n++
    }
    return $candidate
}

function Test-OnboardingEligible {
    # Not every Paycom employee needs an M365 account (e.g. shop floor). Config decides.
    param($Employee, [hashtable]$Rules)
    if (-not $Rules) { return $true }
    foreach ($pair in @(@('IncludeDepartments', 'Department'), @('IncludeJobTitles', 'JobTitle'))) {
        $pattern = $Rules[$pair[0]]
        if ($pattern -and ([string]$Employee.($pair[1]) -notmatch $pattern)) { return $false }
    }
    foreach ($pair in @(@('ExcludeDepartments', 'Department'), @('ExcludeJobTitles', 'JobTitle'))) {
        $pattern = $Rules[$pair[0]]
        if ($pattern -and ([string]$Employee.($pair[1]) -match $pattern)) { return $false }
    }
    return $true
}

function Test-RosterSafety {
    <#
        Guards against acting on a bad export (wrong report, partial download,
        filter mistake) which would otherwise look like a mass termination.
    #>
    param([object[]]$Previous, [object[]]$Current, $Diff, [hashtable]$Safety)
    $problems = New-Object Collections.Generic.List[string]
    $curActive = @($Current | Where-Object IsActive).Count
    if ((Get-ConfigValue $Safety 'MinRosterRows') -and $curActive -lt (Get-ConfigValue $Safety 'MinRosterRows')) {
        $problems.Add("Roster has only $curActive active employees (minimum $((Get-ConfigValue $Safety 'MinRosterRows'))).")
    }
    if ($Previous) {
        $prevActive = @($Previous | Where-Object IsActive).Count
        if ($prevActive -gt 0 -and (Get-ConfigValue $Safety 'MaxShrinkPercent')) {
            $shrink = [math]::Round((($prevActive - $curActive) / $prevActive) * 100, 1)
            if ($shrink -gt (Get-ConfigValue $Safety 'MaxShrinkPercent')) {
                $problems.Add("Active headcount dropped $shrink% ($prevActive -> $curActive), above the $((Get-ConfigValue $Safety 'MaxShrinkPercent'))% limit.")
            }
        }
        if ((Get-ConfigValue $Safety 'MaxTerminations') -and @($Diff.Terminations).Count -gt (Get-ConfigValue $Safety 'MaxTerminations')) {
            $problems.Add("$(@($Diff.Terminations).Count) terminations detected, above the limit of $((Get-ConfigValue $Safety 'MaxTerminations')).")
        }
    }
    [pscustomobject]@{ IsSafe = ($problems.Count -eq 0); Problems = $problems.ToArray() }
}

#endregion

#region Reporting

function New-HtmlTable {
    param([object[]]$Rows, [string[]]$Columns, [string]$Empty = 'None')
    if (-not @($Rows).Count) { return "<p style='color:#666'><i>$Empty</i></p>" }
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append("<table style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px'><tr>")
    foreach ($c in $Columns) { [void]$sb.Append("<th style='text-align:left;border-bottom:2px solid #444;padding:4px 10px'>$(ConvertTo-HtmlText $c)</th>") }
    [void]$sb.Append('</tr>')
    foreach ($r in $Rows) {
        [void]$sb.Append('<tr>')
        foreach ($c in $Columns) { [void]$sb.Append("<td style='border-bottom:1px solid #ddd;padding:4px 10px'>$(ConvertTo-HtmlText $r.$c)</td>") }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function New-LifecycleReport {
    <# Builds the weekly summary email / archived HTML report. #>
    param(
        [Parameter(Mandatory)]$Result,
        [string]$CompanyName = 'IAC'
    )
    $d = $Result.Diff; $r = $Result.Reconciliation
    $lic = { param($u) @($u.assignedLicenses).Count }
    $lastSeen = {
        param($u)
        # lastSignInDateTime also counts failed attempts, so prefer the last successful sign-in.
        $activity = Get-ConfigValue $u 'signInActivity'
        $when = Get-ConfigValue $activity 'lastSuccessfulSignInDateTime'
        if (-not $when) {
            # Not recorded before Dec 2023: fall back to the later of interactive / non-interactive.
            $dates = @('lastSignInDateTime', 'lastNonInteractiveSignInDateTime') | ForEach-Object { Get-ConfigValue $activity $_ } |
                Where-Object { $_ } | ForEach-Object { [datetime]$_ } | Sort-Object -Descending
            if ($dates) { $when = @($dates)[0] }
        }
        if ($when) { ([datetime]$when).ToString('yyyy-MM-dd') } else { '' }
    }

    $hires = foreach ($e in $d.NewHires) {
        $plan = $Result.Plan.Onboard | Where-Object { $_.Employee.EmployeeId -eq $e.EmployeeId } | Select-Object -First 1
        $m = $r.Matches | Where-Object { $_.Employee.EmployeeId -eq $e.EmployeeId } | Select-Object -First 1
        [pscustomobject]@{
            'Employee #' = $e.EmployeeId; Name = $e.DisplayName; Department = $e.Department; Title = $e.JobTitle
            Manager = $e.Manager; 'Start date' = $e.HireDate
            Account = if ($m) { "$($m.User.userPrincipalName) (exists)" } elseif ($plan) { "$($plan.Upn) ($($plan.Status))" } else { 'Not eligible / manual' }
        }
    }
    $terms = foreach ($t in $d.Terminations) {
        $plan = $Result.Plan.Offboard | Where-Object { $_.Employee.EmployeeId -eq $t.Employee.EmployeeId } | Select-Object -First 1
        $m = Find-DirectoryMatch $t.Employee $r.Index
        [pscustomobject]@{
            'Employee #' = $t.Employee.EmployeeId; Name = $t.Employee.DisplayName; Department = $t.Employee.Department
            'Term date' = $t.TermDate; Reason = $t.Reason
            Account = if ($m) { $m.User.userPrincipalName } else { 'No account found' }
            Action = if ($plan) { $plan.Status } elseif ($m -and -not $m.User.accountEnabled) { 'Already disabled' } else { '-' }
        }
    }
    $changes = foreach ($c in $d.Changes) {
        [pscustomobject]@{
            'Employee #' = $c.Employee.EmployeeId; Name = $c.Employee.DisplayName
            Changes = ($c.Changes | ForEach-Object { "$($_.Field): '$($_.Old)' -> '$($_.New)'" }) -join '; '
        }
    }
    $termedEnabled = foreach ($t in $r.TermedButEnabled) {
        [pscustomobject]@{ 'Employee #' = $t.Employee.EmployeeId; Name = $t.Employee.DisplayName; 'Term date' = $t.Employee.TermDate; Account = $t.User.userPrincipalName; Licenses = & $lic $t.User }
    }
    # Leavers who dropped off the report are also orphans; they're already listed under terminations.
    $plannedIds = @($Result.Plan.Offboard | ForEach-Object { $_.User.id })
    $orphans = foreach ($u in ($r.Orphaned | Where-Object { $plannedIds -notcontains $_.id })) {
        [pscustomobject]@{ Account = $u.userPrincipalName; Name = $u.displayName; Department = $u.department; Licenses = & $lic $u; 'Last sign-in' = & $lastSeen $u }
    }
    $missing = foreach ($e in $r.NoAccount) {
        [pscustomobject]@{ 'Employee #' = $e.EmployeeId; Name = $e.DisplayName; Department = $e.Department; Title = $e.JobTitle; 'Hire date' = $e.HireDate }
    }
    $backfill = foreach ($b in $r.IdBackfill) {
        [pscustomobject]@{ 'Employee #' = $b.Employee.EmployeeId; Name = $b.Employee.DisplayName; Account = $b.User.userPrincipalName; 'Matched by' = $b.MatchedBy; Status = $b.Status }
    }

    $h2 = "style='font-family:Segoe UI,Arial,sans-serif;font-size:16px;margin:22px 0 6px'"
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append("<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222'>")
    [void]$sb.Append("<h1 style='font-size:20px;margin:0 0 4px'>$(ConvertTo-HtmlText $CompanyName) user lifecycle report</h1>")
    [void]$sb.Append("<p style='margin:0;color:#555'>Roster: $(ConvertTo-HtmlText $Result.RosterFile) &middot; $(@($Result.Current | Where-Object IsActive).Count) active employees &middot; run $(ConvertTo-HtmlText $Result.RunAt.ToString('yyyy-MM-dd HH:mm')) &middot; mode: $(ConvertTo-HtmlText $Result.Mode)</p>")

    if (-not $Result.Safety.IsSafe) {
        [void]$sb.Append("<div style='background:#fde7e7;border:1px solid #c00;padding:10px;margin:14px 0'><b>Safety check failed - no tickets were created and no accounts were changed.</b><ul>")
        foreach ($p in $Result.Safety.Problems) { [void]$sb.Append("<li>$(ConvertTo-HtmlText $p)</li>") }
        [void]$sb.Append('</ul>Check the export in Paycom and re-run with the correct file. The previous snapshot was kept as the baseline.</div>')
    }
    if ($d.FirstRun) {
        [void]$sb.Append("<div style='background:#eef4ff;border:1px solid #36c;padding:10px;margin:14px 0'>First run: this roster is now the baseline. New hires and terminations are reported from next week's file; the directory reconciliation below is already live.</div>")
    }

    [void]$sb.Append("<table style='margin:12px 0;font-size:14px'><tr>")
    foreach ($kv in @(@('New hires', @($d.NewHires).Count), @('Terminations', @($d.Terminations).Count), @('Changes', @($d.Changes).Count),
            @('Termed but enabled', @($r.TermedButEnabled).Count), @('Orphaned accounts', @($orphans).Count), @('No account', @($r.NoAccount).Count))) {
        [void]$sb.Append("<td style='padding:6px 14px;border:1px solid #ccc;text-align:center'><div style='font-size:22px;font-weight:600'>$($kv[1])</div><div style='color:#555'>$($kv[0])</div></td>")
    }
    [void]$sb.Append('</tr></table>')

    $check = if ($Result.PSObject.Properties['RequestCheck']) { $Result.RequestCheck } else { $null }
    if ($check) {
        $gapHires = foreach ($e in $check.HiresWithoutRequest) {
            [pscustomobject]@{ 'Employee #' = $e.EmployeeId; Name = $e.DisplayName; Department = $e.Department; Title = $e.JobTitle; Manager = $e.Manager; 'Start date' = $e.HireDate }
        }
        $gapTerms = foreach ($t in $check.TerminationsWithoutRequest) {
            [pscustomobject]@{ 'Employee #' = $t.Employee.EmployeeId; Name = $t.Employee.DisplayName; Department = $t.Employee.Department; Manager = $t.Employee.Manager; 'Term date' = $t.TermDate }
        }
        [void]$sb.Append("<h2 $h2>Hired in Paycom with no new-hire request</h2><p style='color:#555;margin:0 0 6px'>Process gap: ask the manager to submit the form so IT can set them up.</p>" + (New-HtmlTable $gapHires @('Employee #', 'Name', 'Department', 'Title', 'Manager', 'Start date')))
        [void]$sb.Append("<h2 $h2>Terminated in Paycom with no termination request</h2><p style='color:#555;margin:0 0 6px'>Process gap: this person left without IT being told. Check their account below.</p>" + (New-HtmlTable $gapTerms @('Employee #', 'Name', 'Department', 'Manager', 'Term date')))
    }
    [void]$sb.Append("<h2 $h2>New hires</h2>" + (New-HtmlTable $hires @('Employee #', 'Name', 'Department', 'Title', 'Manager', 'Start date', 'Account')))
    [void]$sb.Append("<h2 $h2>Terminations</h2>" + (New-HtmlTable $terms @('Employee #', 'Name', 'Department', 'Term date', 'Reason', 'Account', 'Action')))
    [void]$sb.Append("<h2 $h2>Job / department / manager changes</h2>" + (New-HtmlTable $changes @('Employee #', 'Name', 'Changes')))
    [void]$sb.Append("<h2 $h2>Terminated in Paycom but account still enabled</h2>" + (New-HtmlTable $termedEnabled @('Employee #', 'Name', 'Term date', 'Account', 'Licenses')))
    [void]$sb.Append("<h2 $h2>Enabled accounts with no active employee (review)</h2><p style='color:#555;margin:0 0 6px'>Not actioned automatically. Add service or shared accounts to the exclusion list in config.</p>" + (New-HtmlTable $orphans @('Account', 'Name', 'Department', 'Licenses', 'Last sign-in')))
    [void]$sb.Append("<h2 $h2>Active employees with no account</h2><p style='color:#555;margin:0 0 6px'>Expected for roles that don't need one. New hires in this list get an onboarding ticket.</p>" + (New-HtmlTable $missing @('Employee #', 'Name', 'Department', 'Title', 'Hire date')))
    [void]$sb.Append("<h2 $h2>Employee ID backfill</h2><p style='color:#555;margin:0 0 6px'>Accounts matched by email or name. Stamping the Paycom employee code on them makes future matches exact.</p>" + (New-HtmlTable $backfill @('Employee #', 'Name', 'Account', 'Matched by', 'Status')))
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function New-LifecycleTickets {
    <#
        One ticket per joiner / leaver / mover. They're sent as email to the Desk365
        support mailbox, which turns each one into a ticket.
    #>
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][hashtable]$TicketConfig)
    $list = New-Object Collections.Generic.List[object]
    $li = { param($items) ($items | ForEach-Object { "<li>$(ConvertTo-HtmlText $_)</li>" }) -join '' }
    $kv = {
        param($pairs)
        '<table style="font-family:Segoe UI,Arial,sans-serif;font-size:13px">' +
        (($pairs | ForEach-Object { "<tr><td style='padding:2px 12px 2px 0;color:#555'>$(ConvertTo-HtmlText $_[0])</td><td>$(ConvertTo-HtmlText $_[1])</td></tr>" }) -join '') + '</table>'
    }

    # With the request form in place, only raise tickets for what nobody filed a request for.
    $check = if ($Result.PSObject.Properties['RequestCheck']) { $Result.RequestCheck } else { $null }
    $skipHires = @(); $skipTerms = @()
    if ($check -and (Get-ConfigValue $TicketConfig 'OnlyForGaps')) { $skipHires = @($check.CoveredHireIds); $skipTerms = @($check.CoveredTerminationIds) }

    foreach ($e in $Result.Diff.NewHires) {
        if ($skipHires -contains $e.EmployeeId) { continue }
        $plan = $Result.Plan.Onboard | Where-Object { $_.Employee.EmployeeId -eq $e.EmployeeId } | Select-Object -First 1
        $m = $Result.Reconciliation.Matches | Where-Object { $_.Employee.EmployeeId -eq $e.EmployeeId } | Select-Object -First 1
        $acct = if ($m) { "$($m.User.userPrincipalName) (already exists)" } elseif ($plan) { "$($plan.Upn) - $($plan.Status)" } else { 'Not auto-created (not eligible by config). Create manually if needed.' }
        $start = if ($e.HireDate) { $e.HireDate.ToString('yyyy-MM-dd') } else { 'unknown' }
        $body = "<p>Paycom shows a new hire. Please complete onboarding before the start date.</p>" +
            (& $kv @(@('Name', $e.DisplayName), @('Employee #', $e.EmployeeId), @('Start date', $start), @('Department', $e.Department),
                    @('Title', $e.JobTitle), @('Manager', $e.Manager), @('Location', $e.Location), @('Account', $acct))) +
            "<p><b>Checklist</b></p><ul>$(& $li (Get-ConfigValue $TicketConfig 'OnboardingChecklist'))</ul>"
        $list.Add([pscustomobject]@{ Type = 'Onboarding'; Subject = "[Onboarding] $($e.DisplayName) - starts $start"; Body = $body })
    }

    foreach ($t in $Result.Diff.Terminations) {
        if ($skipTerms -contains $t.Employee.EmployeeId) { continue }
        $e = $t.Employee
        $plan = $Result.Plan.Offboard | Where-Object { $_.Employee.EmployeeId -eq $e.EmployeeId } | Select-Object -First 1
        $m = Find-DirectoryMatch $e $Result.Reconciliation.Index
        $done = if ($plan -and $plan.Log) { "<p><b>Automated steps</b></p><ul>$(& $li $plan.Log)</ul>" } else { '' }
        $term = if ($t.TermDate) { $t.TermDate.ToString('yyyy-MM-dd') } else { 'not on report' }
        $body = "<p>Paycom shows this employee is no longer active. Please complete offboarding.</p>" +
            (& $kv @(@('Name', $e.DisplayName), @('Employee #', $e.EmployeeId), @('Term date', $term), @('Reason', $t.Reason),
                    @('Department', $e.Department), @('Manager', $e.Manager), @('Account', $(if ($m) { $m.User.userPrincipalName } else { 'No account found' })),
                    @('Account status', $(if ($plan) { $plan.Status } elseif ($m -and -not $m.User.accountEnabled) { 'Already disabled' } elseif ($m) { 'STILL ENABLED - disable now' } else { '-' })))) +
            $done + "<p><b>Checklist</b></p><ul>$(& $li (Get-ConfigValue $TicketConfig 'OffboardingChecklist'))</ul>"
        $list.Add([pscustomobject]@{ Type = 'Offboarding'; Subject = "[Offboarding] $($e.DisplayName) - $term"; Body = $body })
    }

    if ((Get-ConfigValue $TicketConfig 'CreateChangeTickets')) {
        foreach ($c in $Result.Diff.Changes) {
            if (-not @($c.Changes | Where-Object { @((Get-ConfigValue $TicketConfig 'ChangeTicketFields')) -contains $_.Field }).Count) { continue }
            $rows = $c.Changes | ForEach-Object { @($_.Field, "'$($_.Old)' -> '$($_.New)'") }
            $body = "<p>Paycom shows a role change. Review access (groups, shared mailboxes, app roles) for the new role and update the Entra profile.</p>" +
                (& $kv (@(, @('Name', $c.Employee.DisplayName)) + @(, @('Employee #', $c.Employee.EmployeeId)) + $rows))
            $list.Add([pscustomobject]@{ Type = 'Change'; Subject = "[Access review] $($c.Employee.DisplayName) - role change"; Body = $body })
        }
    }
    foreach ($t in $list) { $t.Body = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222'>$($t.Body)</div>" }
    return $list.ToArray()
}

#endregion

#region Microsoft Graph

$script:GraphInvoker = $null
# Used in @odata.id references. graph.microsoft.us for GCC High / DoD tenants.
$script:GraphResourceBase = 'https://graph.microsoft.com'

function Get-LifecycleGraphBase {
    # Base URL for @odata.id references in the current cloud.
    return $script:GraphResourceBase
}

function Set-LifecycleGraphInvoker {
    # Lets tests (or a different transport) replace the Graph call.
    param([scriptblock]$Invoker)
    $script:GraphInvoker = $Invoker
}

function Invoke-LifecycleGraph {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        $Body,
        [string]$OutputFilePath,
        [hashtable]$Headers
    )
    if ($script:GraphInvoker) { return & $script:GraphInvoker $Method $Uri $Body $OutputFilePath }
    $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Headers) { $params.Headers = $Headers }
    if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10); $params.ContentType = 'application/json' }
    if ($OutputFilePath) { $params.OutputFilePath = $OutputFilePath }
    return Invoke-MgGraphRequest @params
}

function Invoke-LifecycleGraphPaged {
    param([Parameter(Mandatory)][string]$Uri, [hashtable]$Headers)
    $next = $Uri
    while ($next) {
        $page = Invoke-LifecycleGraph -Method GET -Uri $next -Headers $Headers
        foreach ($item in @($page.value)) { $item }
        $next = if ($page.ContainsKey('@odata.nextLink')) { $page['@odata.nextLink'] } else { $null }
    }
}

function ConvertTo-LifecycleObject {
    # Invoke-MgGraphRequest returns hashtables; the reconciliation code expects objects.
    param($Value)
    if ($Value -is [System.Collections.IDictionary]) {
        $o = [ordered]@{}
        foreach ($k in $Value.Keys) { $o[$k] = ConvertTo-LifecycleObject $Value[$k] }
        return [pscustomobject]$o
    }
    if ($Value -is [System.Collections.IList] -and $Value -isnot [string]) {
        return , @($Value | ForEach-Object { ConvertTo-LifecycleObject $_ })
    }
    return $Value
}

function Connect-LifecycleGraph {
    param([Parameter(Mandatory)][hashtable]$Graph)
    if ($script:GraphInvoker) { return }   # a test / replacement transport is in use
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $params = @{ TenantId = $Graph.TenantId; ClientId = $Graph.ClientId; NoWelcome = $true }
    if ((Get-ConfigValue $Graph 'Environment')) {
        $params.Environment = (Get-ConfigValue $Graph 'Environment')
        if ((Get-ConfigValue $Graph 'Environment') -match '^USGov') { $script:GraphResourceBase = 'https://graph.microsoft.us' }
    }
    if ((Get-ConfigValue $Graph 'CertificateThumbprint')) { $params.CertificateThumbprint = (Get-ConfigValue $Graph 'CertificateThumbprint') }
    elseif ((Get-ConfigValue $Graph 'UseManagedIdentity')) { $params = @{ Identity = $true; NoWelcome = $true } }
    else { throw 'Graph config needs CertificateThumbprint (app-only) or UseManagedIdentity = $true.' }
    Connect-MgGraph @params | Out-Null
}

function Get-LifecycleDirectoryUsers {
    param([switch]$IncludeSignInActivity)
    $select = 'id,displayName,givenName,surname,userPrincipalName,mail,employeeId,accountEnabled,userType,assignedLicenses,department,jobTitle,onPremisesSyncEnabled,createdDateTime'
    $top = 999
    # With signInActivity selected, Graph caps pages at 500 users.
    if ($IncludeSignInActivity) { $select += ',signInActivity'; $top = 500 }
    @(Invoke-LifecycleGraphPaged -Uri "/v1.0/users?`$select=$select&`$top=$top" | ForEach-Object { ConvertTo-LifecycleObject $_ })
}

function Get-SharePointRosterFile {
    <# Downloads the newest CSV from the SharePoint drop folder. Returns $null if none. #>
    param([Parameter(Mandatory)][string]$DriveId, [Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Destination)
    $folderPath = ($Folder.Trim('/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    $items = @(Invoke-LifecycleGraphPaged -Uri "/v1.0/drives/$DriveId/root:/$folderPath`:/children?`$select=id,name,lastModifiedDateTime,file" |
        Where-Object { $_.file -and $_.name -match '\.csv$' } | Sort-Object { [datetime]$_.lastModifiedDateTime } -Descending)
    if (-not $items.Count) { return $null }
    $target = Join-Path $Destination $items[0].name
    Invoke-LifecycleGraph -Method GET -Uri "/v1.0/drives/$DriveId/items/$($items[0].id)/content" -OutputFilePath $target | Out-Null
    return [pscustomobject]@{ Path = $target; Name = $items[0].name; Modified = [datetime]$items[0].lastModifiedDateTime }
}

function Send-LifecycleMail {
    param(
        [Parameter(Mandatory)][string]$From,
        [Parameter(Mandatory)][string[]]$To,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Html
    )
    $body = @{
        message         = @{
            subject      = $Subject
            body         = @{ contentType = 'HTML'; content = $Html }
            toRecipients = @($To | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
        }
        saveToSentItems = $true
    }
    Invoke-LifecycleGraph -Method POST -Uri "/v1.0/users/$From/sendMail" -Body $body | Out-Null
}

function Invoke-LifecycleOffboarding {
    <#
        Contain first, clean up later: block sign-in and kill sessions right away.
        Licenses and the mailbox are left alone so a tech can convert the mailbox
        to shared and hand it to the manager (the ticket covers that).
    #>
    param([Parameter(Mandatory)]$User, [Parameter(Mandatory)][hashtable]$Config,
        # For a hire who never started: there's no mailbox to keep, so licence groups go too.
        [switch]$IncludeLicenceGroups)
    $log = New-Object Collections.Generic.List[string]
    $step = {
        param([string]$Name, [scriptblock]$Action)
        try { & $Action; $log.Add("OK: $Name") } catch { $log.Add("FAILED: $Name - $($_.Exception.Message)") }
    }
    $id = $User.id
    $stamp = (Get-Date).ToString('yyyy-MM-dd')

    if ((Get-ConfigValue $Config 'DirectoryMode') -eq 'Hybrid' -and $User.onPremisesSyncEnabled) {
        & $step 'Disable on-prem AD account' {
            Import-Module ActiveDirectory -ErrorAction Stop
            $ad = Get-ADUser -Filter "UserPrincipalName -eq '$($User.userPrincipalName)'" -ErrorAction Stop
            if (-not $ad) { throw 'AD account not found' }
            Disable-ADAccount -Identity $ad -ErrorAction Stop
            Set-ADUser -Identity $ad -Description "Disabled by Paycom lifecycle $stamp" -ErrorAction Stop
            if ((Get-ConfigValue $Config 'Offboarding.DisabledUsersOU')) { Move-ADObject -Identity $ad -TargetPath (Get-ConfigValue $Config 'Offboarding.DisabledUsersOU') -ErrorAction Stop }
        }
    }
    else {
        & $step 'Block sign-in' { Invoke-LifecycleGraph -Method PATCH -Uri "/v1.0/users/$id" -Body @{ accountEnabled = $false } | Out-Null }
    }
    & $step 'Revoke sign-in sessions' { Invoke-LifecycleGraph -Method POST -Uri "/v1.0/users/$id/revokeSignInSessions" | Out-Null }

    if ((Get-ConfigValue $Config 'Offboarding.RemoveGroupMemberships') -or $IncludeLicenceGroups) {
        $groups = @()
        try {
            $groups = @(Invoke-LifecycleGraphPaged -Uri "/v1.0/users/$id/memberOf/microsoft.graph.group?`$select=id,displayName,groupTypes,onPremisesSyncEnabled,assignedLicenses" |
                ForEach-Object { ConvertTo-LifecycleObject $_ })
        }
        catch { $log.Add("FAILED: Read group memberships - $($_.Exception.Message)") }
        foreach ($g in $groups) {
            # Keep licence groups (mailbox must be converted first), dynamic groups and synced groups.
            if (@($g.groupTypes) -contains 'DynamicMembership' -or $g.onPremisesSyncEnabled) { continue }
            if (@($g.assignedLicenses).Count -and -not $IncludeLicenceGroups) { continue }
            if (@((Get-ConfigValue $Config 'Offboarding.KeepGroupIds')) -contains $g.id) { continue }
            & $step "Remove from group '$($g.displayName)'" { Invoke-LifecycleGraph -Method DELETE -Uri "/v1.0/groups/$($g.id)/members/$id/`$ref" | Out-Null }
        }
    }
    if ((Get-ConfigValue $Config 'Offboarding.AddToGroupId')) {
        & $step 'Add to offboarded users group' {
            Invoke-LifecycleGraph -Method POST -Uri "/v1.0/groups/$((Get-ConfigValue $Config 'Offboarding.AddToGroupId'))/members/`$ref" -Body @{ '@odata.id' = "$script:GraphResourceBase/v1.0/directoryObjects/$id" } | Out-Null
        }
    }
    return $log.ToArray()
}

function Invoke-LifecycleOnboarding {
    <#
        Cloud-only account creation. Licences come from department group membership
        (group-based licensing). The random password is discarded; on day one the tech
        issues a Temporary Access Pass so no password is ever emailed.
    #>
    param([Parameter(Mandatory)]$Employee, [Parameter(Mandatory)][string]$Upn, [Parameter(Mandatory)][hashtable]$Config, $ManagerUser,
        [string[]]$ExtraGroupIds = @())
    $log = New-Object Collections.Generic.List[string]
    $o = Get-ConfigValue $Config 'Onboarding'
    $body = @{
        accountEnabled    = $true
        displayName       = $Employee.DisplayName
        givenName         = $Employee.PreferredName
        surname           = $Employee.LastName
        userPrincipalName = $Upn
        mailNickname      = $Upn.Split('@')[0]
        usageLocation     = $(if ((Get-ConfigValue $o 'UsageLocation')) { (Get-ConfigValue $o 'UsageLocation') } else { 'US' })
        passwordProfile   = @{ forceChangePasswordNextSignIn = $true; password = (New-RandomPassword) }
    }
    # The Paycom employee code may not exist yet for a form-driven hire; the weekly audit backfills it.
    foreach ($pair in @(@('employeeId', 'EmployeeId'), @('department', 'Department'), @('jobTitle', 'JobTitle'), @('officeLocation', 'Location'))) {
        if ($Employee.($pair[1])) { $body[$pair[0]] = $Employee.($pair[1]) }
    }
    if ($Employee.HireDate) { $body.employeeHireDate = $Employee.HireDate.ToString('yyyy-MM-ddT00:00:00Z') }
    if ((Get-ConfigValue $o 'CompanyName')) { $body.companyName = (Get-ConfigValue $o 'CompanyName') }

    $created = Invoke-LifecycleGraph -Method POST -Uri '/v1.0/users' -Body $body
    $newId = $created.id
    $log.Add("OK: Created $Upn")

    if ($ManagerUser) {
        try {
            Invoke-LifecycleGraph -Method PUT -Uri "/v1.0/users/$newId/manager/`$ref" -Body @{ '@odata.id' = "$script:GraphResourceBase/v1.0/users/$($ManagerUser.id)" } | Out-Null
            $log.Add("OK: Manager set to $($ManagerUser.userPrincipalName)")
        }
        catch { $log.Add("FAILED: Set manager - $($_.Exception.Message)") }
    }
    $groupIds = @((Get-ConfigValue $o 'DefaultGroupIds')) + @($ExtraGroupIds)
    if ((Get-ConfigValue $o 'DepartmentGroups')) {
        foreach ($pattern in (Get-ConfigValue $o 'DepartmentGroups').Keys) {
            if ([string]$Employee.Department -match $pattern) { $groupIds += @((Get-ConfigValue $o 'DepartmentGroups')[$pattern]) }
        }
    }
    foreach ($gid in ($groupIds | Where-Object { $_ } | Select-Object -Unique)) {
        try {
            Invoke-LifecycleGraph -Method POST -Uri "/v1.0/groups/$gid/members/`$ref" -Body @{ '@odata.id' = "$script:GraphResourceBase/v1.0/directoryObjects/$newId" } | Out-Null
            $log.Add("OK: Added to group $gid")
        }
        catch { $log.Add("FAILED: Add to group $gid - $($_.Exception.Message)") }
    }
    return $log.ToArray()
}

function Set-LifecycleEmployeeId {
    param([Parameter(Mandatory)]$User, [Parameter(Mandatory)][string]$EmployeeId)
    Invoke-LifecycleGraph -Method PATCH -Uri "/v1.0/users/$($User.id)" -Body @{ employeeId = $EmployeeId } | Out-Null
}

#endregion

Export-ModuleMember -Function Get-ConfigValue, Get-AsciiName, Add-DirectoryUserDefaults, Get-NameKey, ConvertTo-RosterDate, Import-PaycomRoster, ConvertTo-RosterIndex,
    Compare-PaycomRoster, Test-DirectoryUserInScope, New-DirectoryIndex, Find-DirectoryMatch, Compare-RosterToDirectory,
    New-UpnCandidate, Test-OnboardingEligible, Test-RosterSafety, New-LifecycleReport, New-LifecycleTickets,
    Set-LifecycleGraphInvoker, Get-LifecycleGraphBase, Invoke-LifecycleGraph, Invoke-LifecycleGraphPaged, ConvertTo-LifecycleObject, Connect-LifecycleGraph,
    Get-LifecycleDirectoryUsers, Get-SharePointRosterFile, Send-LifecycleMail, Invoke-LifecycleOffboarding,
    Invoke-LifecycleOnboarding, Set-LifecycleEmployeeId, New-RandomPassword
