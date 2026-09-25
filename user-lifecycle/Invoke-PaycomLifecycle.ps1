#Requires -Version 5.1
<#
.SYNOPSIS
    Weekly Paycom -> Entra ID user lifecycle run.

.DESCRIPTION
    1. Picks up the latest Paycom "IT Current Employees" export (drop folder or SharePoint).
    2. Diffs it against the last processed snapshot -> new hires, terminations, changes.
    3. Reconciles the roster against Entra ID -> termed-but-enabled, orphaned, missing accounts.
    4. Emails a summary report to IT and one Desk365 ticket per joiner / leaver.
    5. With -Apply (and the matching config switches on), blocks sign-in for leavers,
       creates accounts for joiners and stamps employee IDs.

    Without -Apply nothing in Entra ID is changed.

.EXAMPLE
    # Dry run against a local export and a directory dump, no email:
    ./Invoke-PaycomLifecycle.ps1 -ConfigPath ./config.psd1 -RosterPath ./export.csv -DirectoryJsonPath ./users.json -NoEmail

.EXAMPLE
    # Scheduled production run:
    ./Invoke-PaycomLifecycle.ps1 -ConfigPath ./config.psd1 -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    # Override the configured input with a specific export file.
    [string]$RosterPath,
    # Use a JSON dump of Entra users instead of calling Graph (testing / offline review).
    [string]$DirectoryJsonPath,
    # Use a JSON dump of the lifecycle request list instead of reading SharePoint.
    [string]$RequestsJsonPath,
    # Allow changes in Entra ID / AD. Each action type must also be enabled in config.
    [switch]$Apply,
    # Write the report and tickets to disk only; send nothing.
    [switch]$NoEmail,
    # Process the input even if it's the same file as last run.
    [switch]$Force,
    [datetime]$AsOf = (Get-Date).Date
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'PaycomLifecycle.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'LifecycleRequests.psm1') -Force

$config = Import-PowerShellDataFile -Path $ConfigPath
$configDir = Split-Path -Parent (Resolve-Path $ConfigPath)
$statePath = $config.StatePath
if (-not [IO.Path]::IsPathRooted($statePath)) { $statePath = Join-Path $configDir $statePath }
$snapshotDir = Join-Path $statePath 'snapshots'
$reportDir = Join-Path $statePath 'reports'
$inboxDir = Join-Path $statePath 'incoming'
foreach ($dir in $snapshotDir, $reportDir, $inboxDir) { if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null } }
$stateFile = Join-Path $statePath 'state.json'
$state = if (Test-Path $stateFile) { Get-Content $stateFile -Raw | ConvertFrom-Json } else { [pscustomobject]@{ LastRosterHash = $null; LastRunAt = $null } }

$needGraph = (-not $DirectoryJsonPath) -or (-not $NoEmail) -or $Apply -or (-not $RosterPath -and $config.Input.Source -eq 'SharePoint') -or
    (-not $RequestsJsonPath -and $config.Requests -and $config.Requests.ListId)
if ($needGraph) { Connect-LifecycleGraph -Graph $config.Graph }

function Send-Or-Save {
    param([string]$To, [string]$Subject, [string]$Html, [string]$FileName)
    $Html | Set-Content -Path (Join-Path $reportDir $FileName) -Encoding UTF8
    if (-not $NoEmail) { Send-LifecycleMail -From $config.Mail.From -To $To -Subject $Subject -Html $Html }
}

#region 1. Locate the roster
$rosterFile = $null
if ($RosterPath) {
    $rosterFile = (Resolve-Path $RosterPath).Path
}
elseif ($config.Input.Source -eq 'SharePoint') {
    $dl = Get-SharePointRosterFile -DriveId $config.Input.DriveId -Folder $config.Input.Folder -Destination $inboxDir
    if ($dl) { $rosterFile = $dl.Path }
}
else {
    $dropPath = $config.Input.Path
    if (-not [IO.Path]::IsPathRooted($dropPath)) { $dropPath = Join-Path $configDir $dropPath }
    $latest = Get-ChildItem -Path $dropPath -Filter *.csv -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) { $rosterFile = $latest.FullName }
}

$runStamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
if (-not $rosterFile) {
    Write-Warning 'No roster export found.'
    Send-Or-Save -To $config.Mail.ReportTo -Subject "[Action needed] Paycom roster not found - lifecycle run skipped" -FileName "missing-$runStamp.html" `
        -Html "<p>The weekly lifecycle run didn't find a Paycom export in the drop folder. Download the <b>$($config.Input.ReportName)</b> push report as CSV from Paycom (Report Center &gt; Push Reporting) and save it to the drop folder, then re-run.</p>"
    return
}

$hash = (Get-FileHash -Path $rosterFile -Algorithm SHA256).Hash
if ($hash -eq $state.LastRosterHash -and -not $Force) {
    Write-Warning "Roster '$rosterFile' was already processed on $($state.LastRunAt)."
    if ($config.Input.RemindWhenStale) {
        Send-Or-Save -To $config.Mail.ReportTo -Subject "[Action needed] No new Paycom roster this week" -FileName "stale-$runStamp.html" `
            -Html "<p>The newest file in the drop folder is the same export processed on $($state.LastRunAt). Save this week's <b>$($config.Input.ReportName)</b> export to the drop folder so new hires and terminations are picked up.</p>"
    }
    return
}
#endregion

#region 2. Diff against the previous snapshot
$importArgs = @{ Columns = $config.Roster.Columns; AsOf = $AsOf }
if ($config.Roster.ActiveStatusPattern) { $importArgs.ActiveStatusPattern = $config.Roster.ActiveStatusPattern }
$current = @(Import-PaycomRoster -Path $rosterFile @importArgs)
$prevSnapshot = Get-ChildItem -Path $snapshotDir -Filter 'roster-*.csv' -File | Sort-Object Name -Descending | Select-Object -First 1
$previous = if ($prevSnapshot) { @(Import-PaycomRoster -Path $prevSnapshot.FullName @importArgs) } else { $null }
$diff = Compare-PaycomRoster -Previous $previous -Current $current
$safety = Test-RosterSafety -Previous $previous -Current $current -Diff $diff -Safety $config.Safety
#endregion

#region 3. Reconcile against Entra ID
$directory = if ($DirectoryJsonPath) {
    @(Get-Content $DirectoryJsonPath -Raw | ConvertFrom-Json)
}
else {
    Get-LifecycleDirectoryUsers -IncludeSignInActivity:([bool]$config.Graph.IncludeSignInActivity)
}
$recon = Compare-RosterToDirectory -Roster $current -DirectoryUsers $directory -Scope $config.Scope

# Cross-check against the request form: which Paycom hires/terminations nobody filed a request for.
$requestCheck = $null
$requestItems = if ($RequestsJsonPath) { @(Get-Content $RequestsJsonPath -Raw | ConvertFrom-Json) }
elseif ($config.Requests -and $config.Requests.ListId) {
    $personCache = @{}
    $resolver = { param($id) Resolve-LifecyclePersonEmail -SiteId $config.Requests.SiteId -LookupId $id -Cache $personCache }
    @(Get-LifecycleRequestItems -Config $config)
}
else { $null }
if ($null -ne $requestItems) {
    $resolverArg = if ($RequestsJsonPath) { $null } else { $resolver }
    $requests = @($requestItems | ForEach-Object { ConvertFrom-LifecycleListItem -Item $_ -Config $config -ResolvePerson $resolverArg })
    $requestCheck = Compare-RosterToRequests -Diff $diff -Requests $requests -Index $recon.Index
}
#endregion

#region 4. Plan (and optionally apply) actions
$canAct = $Apply -and $safety.IsSafe
$mode = if ($canAct) { 'Apply' } elseif ($Apply) { 'Apply requested - blocked by safety check' } else { 'Report only' }

$offboard = foreach ($t in $diff.Terminations) {
    $m = Find-DirectoryMatch $t.Employee $recon.Index
    if (-not $m -or -not $m.User.accountEnabled) { continue }
    [pscustomobject]@{ Employee = $t.Employee; User = $m.User; Status = 'Pending - disable sign-in'; Log = @() }
}
$offboard = @($offboard)

$existingUpns = New-Object Collections.Generic.List[string]
foreach ($u in $directory) { $existingUpns.Add([string]$u.userPrincipalName); if ($u.mail) { $existingUpns.Add([string]$u.mail) } }
$noAccountIds = @($recon.NoAccount | ForEach-Object { $_.EmployeeId })
$onboard = foreach ($e in $diff.NewHires) {
    if ($noAccountIds -notcontains $e.EmployeeId) { continue }
    if (-not (Test-OnboardingEligible $e $config.Onboarding.Eligibility)) { continue }
    $upn = New-UpnCandidate -FirstName $e.PreferredName -LastName $e.LastName -Domain $config.Onboarding.Domain -ExistingUpns $existingUpns
    $existingUpns.Add($upn)
    $status = if ($config.DirectoryMode -eq 'Hybrid') { 'Create in on-prem AD (manual)' } else { 'Pending - account to be created' }
    [pscustomobject]@{ Employee = $e; Upn = $upn; Status = $status; Log = @() }
}
$onboard = @($onboard)

foreach ($b in $recon.IdBackfill) { Add-Member -InputObject $b -NotePropertyName Status -NotePropertyValue 'Pending' -Force }

if ($canAct -and $config.Offboarding.Enabled) {
    foreach ($o in $offboard) {
        $o.Log = Invoke-LifecycleOffboarding -User $o.User -Config $config
        $o.Status = if ($o.Log -match '^FAILED') { 'Partially done - see ticket' } else { 'Sign-in blocked, sessions revoked' }
    }
}
if ($canAct -and $config.Onboarding.Enabled -and $config.DirectoryMode -ne 'Hybrid') {
    foreach ($o in $onboard) {
        $mgr = if ($o.Employee.ManagerEmail -and $recon.Index.ByEmail.ContainsKey($o.Employee.ManagerEmail.ToLowerInvariant())) { $recon.Index.ByEmail[$o.Employee.ManagerEmail.ToLowerInvariant()] } else { $null }
        try {
            $o.Log = Invoke-LifecycleOnboarding -Employee $o.Employee -Upn $o.Upn -Config $config -ManagerUser $mgr
            $o.Status = 'Created'
        }
        catch { $o.Status = "Create failed: $($_.Exception.Message)"; $o.Log = @("FAILED: $($_.Exception.Message)") }
    }
}
if ($canAct -and $config.BackfillEmployeeId) {
    foreach ($b in $recon.IdBackfill) {
        try { Set-LifecycleEmployeeId -User $b.User -EmployeeId $b.Employee.EmployeeId; $b.Status = 'Stamped' }
        catch { $b.Status = "Failed: $($_.Exception.Message)" }
    }
}
#endregion

#region 5. Report, tickets, state
$result = [pscustomobject]@{
    RunAt          = Get-Date
    Mode           = $mode
    RosterFile     = Split-Path -Leaf $rosterFile
    Current        = $current
    Previous       = $previous
    Diff           = $diff
    Reconciliation = $recon
    Safety         = $safety
    Plan           = [pscustomobject]@{ Offboard = $offboard; Onboard = $onboard }
    RequestCheck   = $requestCheck
    Tickets        = @()
}

$reportHtml = New-LifecycleReport -Result $result -CompanyName $config.CompanyName
$plannedIds = @($offboard | ForEach-Object { $_.User.id })
$orphanCount = @($recon.Orphaned | Where-Object { $plannedIds -notcontains $_.id }).Count
$counts = "$(@($diff.NewHires).Count) new, $(@($diff.Terminations).Count) termed, $orphanCount orphaned"
$subjectPrefix = if ($safety.IsSafe) { '' } else { '[SAFETY CHECK FAILED] ' }
Send-Or-Save -To $config.Mail.ReportTo -Subject "$subjectPrefix$($config.CompanyName) user lifecycle - $($AsOf.ToString('yyyy-MM-dd')) ($counts)" -Html $reportHtml -FileName "lifecycle-$runStamp.html"

if ($safety.IsSafe) {
    $result.Tickets = @(New-LifecycleTickets -Result $result -TicketConfig $config.Tickets)
    $i = 0
    foreach ($t in $result.Tickets) {
        $i++
        $to = if ($config.Tickets.SendTo) { $config.Tickets.SendTo } else { $config.Mail.ReportTo }
        Send-Or-Save -To $to -Subject $t.Subject -Html $t.Body -FileName ("ticket-$runStamp-{0:D2}-{1}.html" -f $i, $t.Type.ToLowerInvariant())
    }
    # Only a roster that passed the safety check becomes the next baseline.
    Copy-Item -Path $rosterFile -Destination (Join-Path $snapshotDir "roster-$runStamp.csv")
    $keep = if ($config.KeepSnapshots) { [int]$config.KeepSnapshots } else { 26 }
    Get-ChildItem $snapshotDir -Filter 'roster-*.csv' | Sort-Object Name -Descending | Select-Object -Skip $keep | Remove-Item -Force
    @{ LastRosterHash = $hash; LastRunAt = (Get-Date).ToString('s'); LastRosterFile = (Split-Path -Leaf $rosterFile) } |
        ConvertTo-Json | Set-Content -Path $stateFile -Encoding UTF8
}
else {
    Write-Warning ("Safety check failed: " + ($safety.Problems -join ' '))
}

$result
#endregion
