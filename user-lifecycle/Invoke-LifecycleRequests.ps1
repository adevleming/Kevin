#Requires -Version 5.1
<#
.SYNOPSIS
    Processes approved employee lifecycle requests from the SharePoint list.

.DESCRIPTION
    Run every 15 minutes (Task Scheduler or an Azure Automation schedule). For each
    request with Status 'Ready for IT' or 'Scheduled':
      - checks the requester is authorised and the request was approved
      - decides the action: create account / create contact / offboard / remove contact /
        update profile, or wait until the end of the employee's last day
      - with -Apply, does it and writes Status, the account name and a log back to the item,
        then emails the requester, the manager and IT

    Without -Apply it only prints what it would do.

.EXAMPLE
    ./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1              # dry run
    ./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1 -Apply       # scheduled run
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    # Offline testing: list items as JSON (Graph listItem shape) instead of reading SharePoint.
    [string]$RequestsJsonPath,
    # Offline testing: Entra users as JSON instead of calling Graph.
    [string]$DirectoryJsonPath,
    # Offline testing: site settings list items as JSON instead of reading SharePoint.
    [string]$SiteSettingsJsonPath,
    [switch]$Apply,
    [switch]$NoEmail,
    [datetime]$NowUtc = [datetime]::UtcNow
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'PaycomLifecycle.psm1')
Import-Module (Join-Path $PSScriptRoot 'LifecycleRequests.psm1')
if ($NowUtc.Kind -eq [DateTimeKind]::Local) { $NowUtc = $NowUtc.ToUniversalTime() }

$config = Import-PowerShellDataFile -Path $ConfigPath
$configDir = Split-Path -Parent (Resolve-Path $ConfigPath)
$statePath = $config.StatePath
if (-not [IO.Path]::IsPathRooted($statePath)) { $statePath = Join-Path $configDir $statePath }
$logDir = Join-Path $statePath 'request-logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

# One run at a time: a slow run must not overlap the next scheduled one.
$lockPath = Join-Path $statePath 'requests.lock'
$lock = $null
try { $lock = [IO.File]::Open($lockPath, 'OpenOrCreate', 'ReadWrite', 'None') }
catch { Write-Warning 'Another run is in progress; exiting.'; return }

try {
    $offline = [bool]$RequestsJsonPath
    if (-not $offline -or -not $DirectoryJsonPath -or $Apply -or -not $NoEmail) { Connect-LifecycleGraph -Graph $config.Graph }

    #region Site settings (badge office and start-day contacts, kept in a list HR can edit)
    try {
        $settingsItems = if ($SiteSettingsJsonPath) { @(Get-Content $SiteSettingsJsonPath -Raw | ConvertFrom-Json) }
        elseif (-not $offline -and (Get-ConfigValue $config 'Requests.SiteSettingsListId')) { @(Get-LifecycleSiteSettingsItems -Config $config) }
        else { $null }
        if ($null -ne $settingsItems) {
            $merged = Merge-LifecycleSiteSettings -Config $config -Items $settingsItems
            $config.Sites = $merged.Sites
            foreach ($w in $merged.Warnings) { Write-Warning $w }
        }
    }
    catch { Write-Warning "Couldn't read the site settings list; using config.psd1 for badge offices and start-day contacts - $($_.Exception.Message)" }
    #endregion

    #region Load requests and directory
    $personCache = @{}
    $resolver = if ($offline) { $null } else { { param($id) Resolve-LifecyclePersonEmail -SiteId $config.Requests.SiteId -LookupId $id -Cache $personCache } }
    $open = @('Ready for IT', 'Scheduled', 'In progress')
    $items = if ($offline) {
        @(Get-Content $RequestsJsonPath -Raw | ConvertFrom-Json)
    }
    else {
        @(Get-LifecycleRequestItems -Config $config -Status $open) +
        @(Get-LifecycleRequestItems -Config $config -Field 'HireOutcome' -Values @('No-show / not starting', 'Pending start'))
    }
    $seen = @{}
    $all = @(foreach ($it in $items) {
            $r = ConvertFrom-LifecycleListItem -Item $it -Config $config -ResolvePerson $resolver
            if ($seen.ContainsKey($r.Id)) { continue }
            $seen[$r.Id] = $true
            $r
        })
    # Work: open requests, plus completed new hires marked as no-shows.
    $requests = @($all | Where-Object {
            ($open -contains $_.Status) -or ($_.RequestType -eq 'New hire' -and $_.Status -eq 'Completed' -and $_.HireOutcome -eq 'No-show / not starting')
        })
    $pendingStarts = @($all | Where-Object { $_.RequestType -eq 'New hire' -and $_.Status -eq 'Completed' -and $_.HireOutcome -eq 'Pending start' })

    $directory = if ($DirectoryJsonPath) { @(Get-Content $DirectoryJsonPath -Raw | ConvertFrom-Json) } else { Get-LifecycleDirectoryUsers }
    Add-DirectoryUserDefaults $directory
    $index = New-DirectoryIndex $directory
    $existingUpns = New-Object Collections.Generic.List[string]
    foreach ($u in $directory) { $existingUpns.Add([string]$u.userPrincipalName); if ($u.mail) { $existingUpns.Add([string]$u.mail) } }
    #endregion

    # Offline (testing) there's no directory to ask, so everyone counts as a member.
    $memberCheck = if ($offline) { { param($who, $groups) $true } } else { { param($who, $groups) Test-LifecycleGroupMembership -UserEmail $who -GroupIds $groups } }
    $verifyHistory = [bool](Get-ConfigValue $config 'Requests.VerifyApprovalHistory')

    # Employee Status Notification (the email HR used to send by hand) and, for terminations,
    # the site's badge office. Sent once per request, when the automation first accepts it.
    $sendStatusNotice = {
        param($req, [string]$kind, $user)
        $to = @(Get-ConfigValue $config 'Notifications.EmployeeStatusTo')
        if ($kind -eq 'Term') { $to += @(Get-ConfigValue (Get-LifecycleConfigEntry (Get-ConfigValue $config 'Sites') $req.Site) 'BadgeOfficeEmails') }
        if ($kind -eq 'No-show') { $to += @(Get-ConfigValue $config 'Notifications.HrTo') }
        $to = @($to | Where-Object { $_ } | Select-Object -Unique)
        if (-not $to.Count -or $NoEmail) { return $false }
        $n = New-LifecycleStatusNotice -Request $req -Kind $kind -Config $config -User $user
        try { Send-LifecycleMail -From $config.Mail.From -To $to -Subject $n.Subject -Html $n.Body; return $true }
        catch { Write-Warning "Request $($req.Id): status email failed - $($_.Exception.Message)"; return $false }
    }
    $noticeKind = @{ 'New hire' = 'New Hire'; 'Termination' = 'Term'; 'Change' = 'Change' }
    $maxOffboard = if (Get-ConfigValue $config 'Requests.MaxOffboardPerRun') { [int](Get-ConfigValue $config 'Requests.MaxOffboardPerRun') } else { 5 }
    $maxPerDay = if (Get-ConfigValue $config 'Requests.MaxOffboardPerDay') { [int](Get-ConfigValue $config 'Requests.MaxOffboardPerDay') } else { 15 }
    $offboarded = 0
    # Terminations already carried out in the last 24 hours (from earlier runs' logs), so the cap
    # can't be sidestepped 5 at a time every 15 minutes.
    $offboardedToday = 0
    foreach ($f in (Get-ChildItem -Path $logDir -Filter 'run-*.json' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTimeUtc -gt $NowUtc.AddHours(-24) })) {
        try { $offboardedToday += @((Get-Content $f.FullName -Raw | ConvertFrom-Json) | Where-Object { @('Offboard', 'RemoveContact', 'ReverseHire') -contains $_.Action -and $_.Status -ne 'Needs IT review' }).Count } catch { }
    }
    $exchangeFailed = $false
    $exchangeConnected = $false
    $results = New-Object Collections.Generic.List[object]

    foreach ($req in ($requests | Sort-Object { [int]($_.Id -replace '\D', '0') })) {
        $row = [pscustomobject]@{
            Id = $req.Id; Type = $req.RequestType; Person = $(if ($req.DisplayName) { $req.DisplayName } else { $req.EmployeeEmail })
            Action = $null; RunAtUtc = $null; Reason = $null; Status = $req.Status; Upn = $null; Log = @()
        }
        $log = $null; $claim = $null; $subject = $null; $notified = @{}
        try {
            # No location on a termination/change: use the employee's office location from Entra.
            if (-not $req.Site -and $req.EmployeeEmail -and $index.ByEmail.ContainsKey($req.EmployeeEmail.ToLowerInvariant())) {
                $office = [string]$index.ByEmail[$req.EmployeeEmail.ToLowerInvariant()].officeLocation
                $match = @((Get-ConfigValue $config 'Sites') | Where-Object { $_.OfficeLocation -eq $office -or $_.Name -eq $office }) | Select-Object -First 1
                if ($match) { $req.Site = $match.Name }
            }
            if ($req.Status -eq 'In progress') {
                # A previous run claimed this request and never finished (crash, reboot, lost connection).
                # Re-running blindly could create a duplicate account, so a person checks first.
                $decision = [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = 'A previous run stopped part-way through this request. Check what was already done (account, contact, mailbox) before setting it back to Ready for IT.' }
            }
            else {
                $allowed = Test-LifecycleRequestAllowed -Request $req -Config $config -IsMemberOf $memberCheck
                if ($allowed.Allowed -and $verifyHistory) {
                    $versions = if ($offline) { $req.Versions } else { Get-LifecycleRequestVersions -Config $config -ItemId $req.Id }
                    $allowed = Test-LifecycleRequestProvenance -Versions $versions -Config $config
                }
                $decision = if ($allowed.Allowed) { Get-LifecycleRequestAction -Request $req -NowUtc $NowUtc -Config $config }
                else { [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = $allowed.Reason } }
            }

            if (@('Offboard', 'RemoveContact', 'ReverseHire') -contains $decision.Action) {
                if ($offboarded -ge $maxOffboard) {
                    $decision = [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = "More than $maxOffboard terminations in one run - held for a person to check." }
                }
                elseif ($offboardedToday + $offboarded -ge $maxPerDay) {
                    $decision = [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = "More than $maxPerDay terminations in 24 hours - held for a person to check." }
                }
                else { $offboarded++ }
            }
            $row.Action = $decision.Action; $row.RunAtUtc = $decision.RunAtUtc; $row.Reason = $decision.Reason

            if (-not $Apply) { $results.Add($row); continue }

            # First time an approved request is accepted: tell Payroll / the badge office.
            $notified = @{}
            $subject = if ($req.EmployeeEmail -and $index.ByEmail.ContainsKey($req.EmployeeEmail.ToLowerInvariant())) { $index.ByEmail[$req.EmployeeEmail.ToLowerInvariant()] } else { $null }
            if ($decision.Action -ne 'Review' -and -not $req.NotifiedAt -and $noticeKind.ContainsKey($req.RequestType) -and @('ReverseHire', 'CancelHire') -notcontains $decision.Action) {
                if (& $sendStatusNotice $req $noticeKind[$req.RequestType] $subject) { $notified.NotifiedAt = $NowUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') }
            }

            if ($decision.Action -eq 'Wait') {
                $fields = @{} + $notified
                if ($req.Status -ne 'Scheduled') {
                    $when = [TimeZoneInfo]::ConvertTimeFromUtc($decision.RunAtUtc, [TimeZoneInfo]::FindSystemTimeZoneById((Get-SiteTimeZoneId $req.Site $config)))
                    $line = "Scheduled: $($decision.Reason) Runs at $($when.ToString('yyyy-MM-dd HH:mm')) site time."
                    $fields.Status = 'Scheduled'; $fields.ITLog = (Add-LifecycleLogLines $req.ITLog @($line) $NowUtc)
                    $row.Status = 'Scheduled'
                }
                if ($fields.Count) { Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields $fields }
                $results.Add($row); continue
            }

            $log = $req.ITLog
            $claim = $null
            if ($decision.Action -eq 'Review') {
                $outcome = [pscustomobject]@{ Status = 'Needs IT review'; Upn = $null; Log = @("REVIEW: $($decision.Reason)") }
            }
            else {
                # Claim the request first so an interrupted run can't repeat it.
                $claim = @{ Status = 'In progress' } + $notified
                $started = "Started: $($decision.Action)"
                if ($decision.Action -eq 'CreateUser' -and -not $req.ITUpn) {
                    # Record the account name up front so a re-run can tell it was already created.
                    $planned = New-UpnCandidate -FirstName $req.PreferredName -LastName $req.LastName -Domain $config.Onboarding.Domain -ExistingUpns $existingUpns
                    $claim.ITUpn = $planned; $started += " as $planned"
                }
                $log = Add-LifecycleLogLines $log @($started) $NowUtc
                $claim.ITLog = $log
                Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields $claim
                if ((Test-RequestNeedsExchange $decision.Action $config) -and -not $exchangeConnected -and -not $exchangeFailed) {
                    # A failed Exchange connection must not stop sign-in blocking: the Exchange steps
                    # then fail on their own and the request goes to review.
                    try { Connect-LifecycleExchange -Exchange $config.Exchange; $exchangeConnected = $true }
                    catch { $exchangeFailed = $true; Write-Warning "Exchange Online connection failed: $($_.Exception.Message)" }
                }
                $outcome = Invoke-LifecycleRequestAction -Request $req -Action $decision.Action -Index $index -Config $config -ExistingUpns $existingUpns
            }

            $fields = @{ Status = $outcome.Status; ITLog = (Add-LifecycleLogLines $log $outcome.Log $NowUtc); ProcessedAt = $NowUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') }
            if ($decision.Action -eq 'Review' -and $notified.Count) { $fields += $notified }
            if ($outcome.Upn) { $fields.ITUpn = $outcome.Upn }
            elseif ($decision.Action -eq 'CreateUser' -and $claim -and $claim.ContainsKey('ITUpn')) { $fields.ITUpn = '' }   # planned name, never created
            Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields $fields
            $row.Status = $outcome.Status; $row.Upn = $outcome.Upn; $row.Log = $outcome.Log
            if (@('Reversed - did not start', 'Cancelled') -contains $outcome.Status) { & $sendStatusNotice $req 'No-show' $null | Out-Null }

            if (-not $NoEmail) {
                $notice = New-LifecycleRequestNotice -Request $req -Outcome $outcome -Action $decision.Action -ListUrl $config.Requests.ListUrl -User $subject
                $to = @($req.RequesterEmail, $req.ManagerEmail, $config.Mail.ReportTo) | Where-Object { $_ } | Select-Object -Unique
                try { Send-LifecycleMail -From $config.Mail.From -To $to -Subject $notice.Subject -Html $notice.Body }
                catch { Write-Warning "Request $($req.Id): notice email failed - $($_.Exception.Message)" }
            }
        }
        catch {
            # One bad request must not stop the rest. Leave it for a person, with the error.
            $err = $_.Exception.Message
            Write-Warning "Request $($req.Id): $err"
            $row.Status = 'Needs IT review'; $row.Log = @("FAILED: $err")
            if ($Apply) {
                # Keep the log of whatever already ran (e.g. the claim) and tell IT.
                $soFar = if ($log) { $log } else { $req.ITLog }
                try { Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields @{ Status = 'Needs IT review'; ITLog = (Add-LifecycleLogLines $soFar @("FAILED: $err") $NowUtc) } }
                catch { Write-Warning "Request $($req.Id): couldn't record the failure - $($_.Exception.Message)" }
                if (-not $NoEmail -and $config.Mail.ReportTo) {
                    try {
                        Send-LifecycleMail -From $config.Mail.From -To @($config.Mail.ReportTo) -Subject "[IT] Lifecycle request $($req.Id) failed" `
                            -Html "<p>Request #$([System.Net.WebUtility]::HtmlEncode($req.Id)) ($([System.Net.WebUtility]::HtmlEncode($req.RequestType))) failed and is now <b>Needs IT review</b>:</p><pre>$([System.Net.WebUtility]::HtmlEncode($err))</pre>"
                    }
                    catch { }
                }
            }
        }
        $results.Add($row)
    }

    #region Start-day check: "did everyone start?"
    $startChecks = New-Object Collections.Generic.List[object]
    foreach ($batch in (Get-LifecycleStartCheckBatches -Requests $pendingStarts -NowUtc $NowUtc -Config $config)) {
        $ask = New-Object Collections.Generic.List[object]
        foreach ($r in $batch.Requests) {
            # Anyone whose new account has already signed in has obviously started.
            $started = $false
            if ($Apply -and $r.ITUpn -and $index.ByEmail.ContainsKey($r.ITUpn.ToLowerInvariant())) {
                try {
                    $u = $index.ByEmail[$r.ITUpn.ToLowerInvariant()]
                    $activity = Get-FieldValue (Invoke-LifecycleGraph -Method GET -Uri "/v1.0/users/$($u.id)?`$select=signInActivity") 'signInActivity'
                    if (Get-FieldValue $activity 'lastSuccessfulSignInDateTime') {
                        Update-LifecycleRequestItem -Config $config -ItemId $r.Id -Fields @{ HireOutcome = 'Started'; ITLog = (Add-LifecycleLogLines $r.ITLog @('Marked Started: their account has signed in.') $NowUtc) }
                        $started = $true
                    }
                }
                catch { Write-Warning "Request $($r.Id): couldn't check sign-in activity - $($_.Exception.Message)" }
            }
            if (-not $started) { $ask.Add($r) }
        }
        $startChecks.Add([pscustomobject]@{ Site = $batch.Site; Asked = @($ask | ForEach-Object Id); AutoStarted = @($batch.Requests | Where-Object { $ask -notcontains $_ } | ForEach-Object Id) })
        foreach ($r in $batch.Requests) {
            $results.Add([pscustomobject]@{
                    Id = $r.Id; Type = 'New hire'; Person = $r.DisplayName; Action = $(if ($ask -contains $r) { 'AskIfStarted' } else { 'MarkedStarted' })
                    RunAtUtc = $null; Reason = "Start date $($r.StartDate.ToString('yyyy-MM-dd')) at $($batch.Site)"; Status = $r.Status; Upn = $r.ITUpn; Log = @()
                })
        }
        if (-not $Apply -or -not $ask.Count) { continue }
        if (-not $NoEmail) {
            $notice = New-LifecycleStartCheckNotice -Site $batch.Site -Requests $ask.ToArray() -ListUrl $config.Requests.ListUrl
            $site = Get-LifecycleConfigEntry (Get-ConfigValue $config 'Sites') $batch.Site
            $to = @($ask | ForEach-Object { $_.RequesterEmail }) + @(Get-ConfigValue $site 'OrientationContacts') + @($config.Mail.ReportTo)
            $to = @($to | Where-Object { $_ } | Select-Object -Unique)
            try { Send-LifecycleMail -From $config.Mail.From -To $to -Subject $notice.Subject -Html $notice.Body }
            catch { Write-Warning "Start-day check for $($batch.Site) failed - $($_.Exception.Message)"; continue }
        }
        foreach ($r in $ask) {
            Update-LifecycleRequestItem -Config $config -ItemId $r.Id -Fields @{ StartCheckSentAt = $NowUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') }
        }
    }
    #endregion

    if ($Apply -and $results.Count) {
        $results | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $logDir ("run-{0}.json" -f $NowUtc.ToString('yyyyMMdd-HHmmss'))) -Encoding UTF8
    }
    if ($startChecks.Count) { Write-Verbose ("Start-day checks: " + (($startChecks | ForEach-Object { "$($_.Site): asked $($_.Asked.Count), auto-started $($_.AutoStarted.Count)" }) -join '; ')) }
    if ($exchangeConnected -and (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue)) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
    $results.ToArray()
}
finally {
    if ($lock) { $lock.Dispose() }
}
