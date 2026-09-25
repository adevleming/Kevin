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

    #region Load requests and directory
    $personCache = @{}
    $resolver = if ($offline) { $null } else { { param($id) Resolve-LifecyclePersonEmail -SiteId $config.Requests.SiteId -LookupId $id -Cache $personCache } }
    $items = if ($offline) {
        @(Get-Content $RequestsJsonPath -Raw | ConvertFrom-Json)
    }
    else {
        @(Get-LifecycleRequestItems -Config $config -Status @('Ready for IT', 'Scheduled', 'In progress'))
    }
    $requests = @($items | ForEach-Object { ConvertFrom-LifecycleListItem -Item $_ -Config $config -ResolvePerson $resolver } |
        Where-Object { @('Ready for IT', 'Scheduled', 'In progress') -contains $_.Status })

    $directory = if ($DirectoryJsonPath) { @(Get-Content $DirectoryJsonPath -Raw | ConvertFrom-Json) } else { Get-LifecycleDirectoryUsers }
    Add-DirectoryUserDefaults $directory
    $index = New-DirectoryIndex $directory
    $existingUpns = New-Object Collections.Generic.List[string]
    foreach ($u in $directory) { $existingUpns.Add([string]$u.userPrincipalName); if ($u.mail) { $existingUpns.Add([string]$u.mail) } }
    #endregion

    $memberCheck = if ($offline) { $null } else { { param($email, $groups) Test-LifecycleGroupMembership -UserEmail $email -GroupIds $groups } }
    $maxOffboard = if ($config.Requests.MaxOffboardPerRun) { [int]$config.Requests.MaxOffboardPerRun } else { 5 }
    $offboarded = 0
    $exchangeConnected = $false
    $results = New-Object Collections.Generic.List[object]

    foreach ($req in ($requests | Sort-Object { [int]($_.Id -replace '\D', '0') })) {
        $row = [pscustomobject]@{
            Id = $req.Id; Type = $req.RequestType; Person = $(if ($req.DisplayName) { $req.DisplayName } else { $req.EmployeeEmail })
            Action = $null; RunAtUtc = $null; Reason = $null; Status = $req.Status; Upn = $null; Log = @()
        }
        try {
            if ($req.Status -eq 'In progress') {
                # A previous run claimed this request and never finished (crash, reboot, lost connection).
                # Re-running blindly could create a duplicate account, so a person checks first.
                $decision = [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = 'A previous run stopped part-way through this request. Check what was already done (account, contact, mailbox) before setting it back to Ready for IT.' }
            }
            else {
                $allowed = Test-LifecycleRequestAllowed -Request $req -Config $config -IsMemberOf $memberCheck
                $decision = if ($allowed.Allowed) { Get-LifecycleRequestAction -Request $req -NowUtc $NowUtc -Config $config }
                else { [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = $allowed.Reason } }
            }

            if (@('Offboard', 'RemoveContact') -contains $decision.Action) {
                if ($offboarded -ge $maxOffboard) {
                    $decision = [pscustomobject]@{ Action = 'Review'; RunAtUtc = $null; Reason = "More than $maxOffboard terminations in one run - held for a person to check." }
                }
                else { $offboarded++ }
            }
            $row.Action = $decision.Action; $row.RunAtUtc = $decision.RunAtUtc; $row.Reason = $decision.Reason

            if (-not $Apply) { $results.Add($row); continue }

            if ($decision.Action -eq 'Wait') {
                if ($req.Status -ne 'Scheduled') {
                    $when = [TimeZoneInfo]::ConvertTimeFromUtc($decision.RunAtUtc, [TimeZoneInfo]::FindSystemTimeZoneById((Get-SiteTimeZoneId $req.Site $config)))
                    $line = "Scheduled: $($decision.Reason) Runs at $($when.ToString('yyyy-MM-dd HH:mm')) site time."
                    Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields @{ Status = 'Scheduled'; ITLog = (Add-LifecycleLogLines $req.ITLog @($line) $NowUtc) }
                    $row.Status = 'Scheduled'
                }
                $results.Add($row); continue
            }

            $log = $req.ITLog
            if ($decision.Action -eq 'Review') {
                $outcome = [pscustomobject]@{ Status = 'Needs IT review'; Upn = $null; Log = @("REVIEW: $($decision.Reason)") }
            }
            else {
                # Claim the request first so an interrupted run can't repeat it.
                $claim = @{ Status = 'In progress' }
                $started = "Started: $($decision.Action)"
                if ($decision.Action -eq 'CreateUser' -and -not $req.ITUpn) {
                    # Record the account name up front so a re-run can tell it was already created.
                    $planned = New-UpnCandidate -FirstName $req.PreferredName -LastName $req.LastName -Domain $config.Onboarding.Domain -ExistingUpns $existingUpns
                    $claim.ITUpn = $planned; $started += " as $planned"
                }
                $log = Add-LifecycleLogLines $log @($started) $NowUtc
                $claim.ITLog = $log
                Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields $claim
                if ((Test-RequestNeedsExchange $decision.Action $config) -and -not $exchangeConnected) {
                    Connect-LifecycleExchange -Exchange $config.Exchange
                    $exchangeConnected = $true
                }
                $outcome = Invoke-LifecycleRequestAction -Request $req -Action $decision.Action -Index $index -Config $config -ExistingUpns $existingUpns
            }

            $fields = @{ Status = $outcome.Status; ITLog = (Add-LifecycleLogLines $log $outcome.Log $NowUtc); ProcessedAt = $NowUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') }
            if ($outcome.Upn) { $fields.ITUpn = $outcome.Upn }
            Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields $fields
            $row.Status = $outcome.Status; $row.Upn = $outcome.Upn; $row.Log = $outcome.Log

            if (-not $NoEmail) {
                $notice = New-LifecycleRequestNotice -Request $req -Outcome $outcome -Action $decision.Action -ListUrl $config.Requests.ListUrl
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
                try { Update-LifecycleRequestItem -Config $config -ItemId $req.Id -Fields @{ Status = 'Needs IT review'; ITLog = (Add-LifecycleLogLines $req.ITLog @("FAILED: $err") $NowUtc) } }
                catch { Write-Warning "Request $($req.Id): couldn't record the failure - $($_.Exception.Message)" }
            }
        }
        $results.Add($row)
    }

    if ($Apply -and $results.Count) {
        $results | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $logDir ("run-{0}.json" -f $NowUtc.ToString('yyyyMMdd-HHmmss'))) -Encoding UTF8
    }
    if ($exchangeConnected -and (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue)) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
    $results.ToArray()
}
finally {
    if ($lock) { $lock.Dispose() }
}
