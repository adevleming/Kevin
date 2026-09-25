#Requires -Version 5.1
<#
    LifecycleRequests - the form-driven side of user lifecycle automation.

    Hiring managers / HR submit New hire, Termination and Change requests through a
    SharePoint list form pinned in Teams. A Power Automate flow sends each one to HR
    for approval and sets Status = 'Ready for IT'. This module picks those up, works
    out which IT action is due (now, or at the end of an employee's last day), carries
    it out through Microsoft Graph / Exchange Online, and writes the outcome back to
    the list item.

    Decision and normalisation functions are pure so they can be tested offline.
    Everything that touches a service goes through Invoke-LifecycleGraph or
    Invoke-LifecycleExchange, both of which can be swapped for a mock.
#>
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'PaycomLifecycle.psm1')

$script:Status = @{
    Submitted       = 'Submitted'
    PendingApproval = 'Pending approval'
    Ready           = 'Ready for IT'
    Scheduled       = 'Scheduled'
    InProgress      = 'In progress'
    Completed       = 'Completed'
    Review          = 'Needs IT review'
    Rejected        = 'Rejected'
    Cancelled       = 'Cancelled'
}
$script:RequestTypes = @('New hire', 'Termination', 'Change')

#region Helpers

function Get-FieldValue {
    # Reads a property from a Graph hashtable or a JSON/PSCustomObject, $null if absent.
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function ConvertTo-LifecycleBool {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    return @('true', 'yes', '1') -contains ([string]$Value).Trim().ToLowerInvariant()
}

function Get-LifecycleConfigEntry {
    # Config lists (Sites, AccessTypes) are arrays of hashtables keyed by Name.
    param([object[]]$List, [string]$Name)
    if (-not $Name) { return $null }
    foreach ($entry in @($List)) { if ($entry -and (Get-ConfigValue $entry 'Name') -eq $Name) { return $entry } }
    return $null
}

function Get-SiteTimeZoneId {
    param([string]$Site, [hashtable]$Config)
    $entry = Get-LifecycleConfigEntry (Get-ConfigValue $Config 'Sites') $Site
    if ($entry -and (Get-ConfigValue $entry 'TimeZone')) { return (Get-ConfigValue $entry 'TimeZone') }
    if ((Get-ConfigValue $Config 'Requests.DefaultTimeZone')) { return (Get-ConfigValue $Config 'Requests.DefaultTimeZone') }
    return 'Pacific Standard Time'
}

function Get-SiteLocalTimeUtc {
    # "18:00 on the employee's last day, at their site" as a UTC instant.
    param([Parameter(Mandatory)][datetime]$Date, [string]$Time = '00:00', [string]$Site, [Parameter(Mandatory)][hashtable]$Config)
    $tz = [TimeZoneInfo]::FindSystemTimeZoneById((Get-SiteTimeZoneId $Site $Config))
    $local = [datetime]::SpecifyKind($Date.Date.Add([timespan]::Parse($Time)), [DateTimeKind]::Unspecified)
    return [TimeZoneInfo]::ConvertTimeToUtc($local, $tz)
}

function ConvertFrom-ListDate {
    <#
        SharePoint stores a date-only column as midnight in the SITE's regional time zone
        and Graph returns it in UTC (e.g. 2026-09-28T07:00:00Z for a Pacific site), so the
        calendar date has to be recovered by converting back to that time zone.
    #>
    param($Value, [string]$TimeZoneId = 'Pacific Standard Time')
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -and -not $Value.Trim()) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return $Value.Date }
        $utc = $Value.ToUniversalTime()
    }
    elseif ($Value -is [DateTimeOffset]) { $utc = $Value.UtcDateTime }
    else {
        $s = ([string]$Value).Trim()
        if ($s -match '^\d{4}-\d{2}-\d{2}$') { return [datetime]::ParseExact($s, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
        $utc = [DateTimeOffset]::Parse($s, [Globalization.CultureInfo]::InvariantCulture).UtcDateTime
    }
    $tz = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    return [TimeZoneInfo]::ConvertTimeFromUtc($utc, $tz).Date
}

#endregion

#region Exchange Online

$script:ExchangeInvoker = $null

function Set-LifecycleExchangeInvoker {
    param([scriptblock]$Invoker)
    $script:ExchangeInvoker = $Invoker
}

function Invoke-LifecycleExchange {
    # Runs an Exchange Online cmdlet by name. Tests replace this with a mock.
    param([Parameter(Mandatory)][string]$Command, [hashtable]$Parameters = @{})
    if ($script:ExchangeInvoker) { return & $script:ExchangeInvoker $Command $Parameters }
    $p = @{} + $Parameters
    $p.ErrorAction = 'Stop'
    return & $Command @p
}

function Connect-LifecycleExchange {
    param([Parameter(Mandatory)][hashtable]$Exchange)
    if ($script:ExchangeInvoker) { return }   # a test / replacement transport is in use
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    if ((Get-ConfigValue $Exchange 'UseManagedIdentity')) {
        Connect-ExchangeOnline -ManagedIdentity -Organization $Exchange.Organization -ShowBanner:$false -ErrorAction Stop
    }
    else {
        Connect-ExchangeOnline -AppId $Exchange.AppId -CertificateThumbprint $Exchange.CertificateThumbprint -Organization $Exchange.Organization -ShowBanner:$false -ErrorAction Stop
    }
}

function Test-RequestNeedsExchange {
    param([string]$Action, [hashtable]$Config)
    if (@('CreateContact', 'RemoveContact') -contains $Action) { return $true }
    return ($Action -eq 'Offboard' -and [bool](Get-ConfigValue $Config 'Offboarding.ConvertMailboxToShared'))
}

#endregion

#region SharePoint list

function Get-LifecycleRequestListColumns {
    <#
        Column definitions for the "Employee Lifecycle Requests" list, in Microsoft Graph
        columnDefinition format. Choices for sites, departments and access types come
        from config so the form and the automation can't drift apart.
    #>
    param([Parameter(Mandatory)][hashtable]$Config)
    $choice = {
        param($name, $display, $choices, [bool]$textEntry = $false, [string]$displayAs = 'dropDownMenu')
        @{ name = $name; displayName = $display; choice = @{ choices = @($choices); displayAs = $displayAs; allowTextEntry = $textEntry } }
    }
    $text = { param($name, $display) @{ name = $name; displayName = $display; text = @{} } }
    $multi = { param($name, $display) @{ name = $name; displayName = $display; text = @{ allowMultipleLines = $true; linesForEditing = 6 } } }
    $date = { param($name, $display) @{ name = $name; displayName = $display; dateTime = @{ format = 'dateOnly'; displayAs = 'default' } } }
    $person = { param($name, $display) @{ name = $name; displayName = $display; personOrGroup = @{ allowMultipleSelection = $false; chooseFromType = 'peopleOnly' } } }

    $status = & $choice 'Status' 'Status' @($script:Status.Submitted, $script:Status.PendingApproval, $script:Status.Ready, $script:Status.Scheduled,
        $script:Status.InProgress, $script:Status.Completed, $script:Status.Review, $script:Status.Rejected, $script:Status.Cancelled)
    $status.indexed = $true
    $status.defaultValue = @{ value = $script:Status.Submitted }

    $requestType = & $choice 'RequestType' 'Request type' $script:RequestTypes
    $requestType.required = $true

    @(
        $requestType
        $status
        & $choice 'AccessType' 'Computer access' @((Get-ConfigValue $Config 'AccessTypes') | ForEach-Object { $_.Name })
        & $text 'FirstName' 'Legal first name'
        & $text 'PreferredName' 'Preferred first name'
        & $text 'LastName' 'Last name'
        & $person 'Employee' 'Employee'
        & $choice 'Site' 'Site' @((Get-ConfigValue $Config 'Sites') | ForEach-Object { $_.Name }) $true
        & $choice 'Department' 'Department' @(Get-ConfigValue $Config 'Departments') $true
        & $text 'JobTitle' 'Job title'
        & $person 'Manager' 'Manager'
        & $date 'StartDate' 'Start date'
        & $date 'LastDay' 'Last day worked'
        & $choice 'TerminationType' 'Termination type' @('Voluntary', 'Involuntary')
        @{ name = 'DisableImmediately'; displayName = 'Disable access immediately'; boolean = @{} }
        & $person 'MailboxDelegate' 'Give mailbox and files to'
        & $date 'EffectiveDate' 'Change effective date'
        & $text 'PersonalEmail' 'Personal email'
        & $text 'MobilePhone' 'Mobile phone'
        & $choice 'Equipment' 'Equipment needed' @((Get-ConfigValue $Config 'Requests.EquipmentChoices')) $false 'checkBoxes'
        & $multi 'Notes' 'Notes for HR / IT'
        & $text 'PaycomEmployeeId' 'Paycom employee code'
        & $text 'ApprovedBy' 'Approved by'
        & $text 'ITUpn' 'Account created'
        & $multi 'ITLog' 'IT automation log'
        @{ name = 'ProcessedAt'; displayName = 'Processed at'; dateTime = @{ format = 'dateTime'; displayAs = 'default' } }
    )
}

function Resolve-LifecyclePersonEmail {
    # Person columns come back from Graph as <Name>LookupId, an ID in the site's hidden
    # User Information List. Resolve it to the person's email / UPN.
    param([Parameter(Mandatory)][string]$SiteId, $LookupId, [hashtable]$Cache = @{})
    if (-not $LookupId) { return $null }
    $key = [string]$LookupId
    if ($Cache.ContainsKey($key)) { return $Cache[$key] }
    $item = Invoke-LifecycleGraph -Method GET -Uri "/v1.0/sites/$SiteId/lists/User%20Information%20List/items/$key`?expand=fields"
    $f = Get-FieldValue $item 'fields'
    $email = Get-FieldValue $f 'EMail'
    if (-not $email) { $email = Get-FieldValue $f 'UserName' }
    $Cache[$key] = $email
    return $email
}

function ConvertFrom-LifecycleListItem {
    <# Turns a Graph listItem (with expanded fields) into a flat request object. #>
    param([Parameter(Mandatory)]$Item, [Parameter(Mandatory)][hashtable]$Config, [scriptblock]$ResolvePerson)
    $f = Get-FieldValue $Item 'fields'
    $get = {
        param($n)
        $v = Get-FieldValue $f $n
        if ($v -is [string]) { return $v.Trim() }
        return $v
    }
    $person = {
        param($name)
        $lookup = Get-FieldValue $f "${name}LookupId"
        if ($lookup -and $ResolvePerson) {
            $resolved = & $ResolvePerson $lookup
            if ($resolved) { return ([string]$resolved).Trim() }
        }
        $typed = Get-FieldValue $f "${name}Email"
        if ($typed) { return ([string]$typed).Trim() }
        return $null
    }
    $tzId = if ((Get-ConfigValue $Config 'Requests.SiteTimeZone')) { (Get-ConfigValue $Config 'Requests.SiteTimeZone') } else { 'Pacific Standard Time' }
    $first = & $get 'FirstName'
    $preferred = & $get 'PreferredName'
    $last = & $get 'LastName'
    $usedFirst = if ($preferred) { $preferred } else { $first }
    $modifiedBy = Get-FieldValue (Get-FieldValue $Item 'lastModifiedBy') 'user'
    # createdBy.user.email is usually present but not guaranteed; fall back to the Author person field.
    $requester = Get-FieldValue (Get-FieldValue (Get-FieldValue $Item 'createdBy') 'user') 'email'
    if (-not $requester -and $ResolvePerson) {
        $author = Get-FieldValue $f 'AuthorLookupId'
        if ($author) { $requester = & $ResolvePerson $author }
    }

    [pscustomobject]@{
        Id                   = [string](Get-FieldValue $Item 'id')
        RequestType          = & $get 'RequestType'
        Status               = & $get 'Status'
        AccessType           = & $get 'AccessType'
        FirstName            = $first
        PreferredName        = $usedFirst
        LastName             = $last
        DisplayName          = "$usedFirst $last".Trim()
        EmployeeEmail        = & $person 'Employee'
        Site                 = & $get 'Site'
        Department           = & $get 'Department'
        JobTitle             = & $get 'JobTitle'
        ManagerEmail         = & $person 'Manager'
        StartDate            = ConvertFrom-ListDate (& $get 'StartDate') $tzId
        LastDay              = ConvertFrom-ListDate (& $get 'LastDay') $tzId
        EffectiveDate        = ConvertFrom-ListDate (& $get 'EffectiveDate') $tzId
        TerminationType      = & $get 'TerminationType'
        DisableImmediately   = ConvertTo-LifecycleBool (& $get 'DisableImmediately')
        MailboxDelegateEmail = & $person 'MailboxDelegate'
        PersonalEmail        = & $get 'PersonalEmail'
        MobilePhone          = & $get 'MobilePhone'
        Notes                = & $get 'Notes'
        PaycomEmployeeId     = & $get 'PaycomEmployeeId'
        ApprovedBy           = & $get 'ApprovedBy'
        ITLog                = & $get 'ITLog'
        ITUpn                = & $get 'ITUpn'
        RequesterEmail       = $requester
        LastModifiedByEmail  = Get-FieldValue $modifiedBy 'email'
    }
}

function Get-LifecycleRequestItems {
    <# Raw list items, optionally only those in the given statuses (one query per status). #>
    param([Parameter(Mandatory)][hashtable]$Config, [string[]]$Status)
    $base = "/v1.0/sites/$($Config.Requests.SiteId)/lists/$($Config.Requests.ListId)/items?expand=fields&`$top=200"
    if (-not $Status) { return @(Invoke-LifecycleGraphPaged -Uri $base) }
    $headers = @{ Prefer = 'HonorNonIndexedQueriesWarningMayFailRandomly' }
    foreach ($s in $Status) {
        $filter = [uri]::EscapeDataString("fields/Status eq '$s'")
        Invoke-LifecycleGraphPaged -Uri "$base&`$filter=$filter" -Headers $headers
    }
}

function Update-LifecycleRequestItem {
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$ItemId, [Parameter(Mandatory)][hashtable]$Fields)
    Invoke-LifecycleGraph -Method PATCH -Uri "/v1.0/sites/$($Config.Requests.SiteId)/lists/$($Config.Requests.ListId)/items/$ItemId/fields" -Body $Fields | Out-Null
}

function Add-LifecycleLogLines {
    # Appends timestamped lines to the ITLog text kept on the request.
    param([string]$Existing, [string[]]$Lines, [datetime]$NowUtc = [datetime]::UtcNow)
    $stamp = $NowUtc.ToString('yyyy-MM-dd HH:mm') + ' UTC'
    $new = ($Lines | Where-Object { $_ } | ForEach-Object { "[$stamp] $_" }) -join "`n"
    if ($Existing) { return "$Existing`n$new" }
    return $new
}

#endregion

#region Decisions

function Test-LifecycleRequestAllowed {
    <#
        Defence in depth on top of list permissions and the approval flow: the requester
        must belong to an authorised group, and anything other than an immediate
        termination must carry an approval.
    #>
    param([Parameter(Mandatory)]$Request, [Parameter(Mandatory)][hashtable]$Config, [scriptblock]$IsMemberOf)
    $r = $Config.Requests
    if (@((Get-ConfigValue $r 'AuthorizedGroupIds')).Count -and $IsMemberOf) {
        if (-not $Request.RequesterEmail) { return [pscustomobject]@{ Allowed = $false; Reason = 'Requester unknown.' } }
        if (-not (& $IsMemberOf $Request.RequesterEmail @((Get-ConfigValue $r 'AuthorizedGroupIds')))) {
            return [pscustomobject]@{ Allowed = $false; Reason = "Requester $($Request.RequesterEmail) is not in an authorised hiring group." }
        }
    }
    # Requests are locked read-only for the requester once submitted. If someone outside the
    # trusted editors (HR approvers, the flow's account) changed it last, a person should look.
    # Edits made by this automation (app-only) carry no email and are fine.
    $trusted = @(Get-ConfigValue $r 'TrustedEditors' | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
    if ($trusted.Count -and $Request.LastModifiedByEmail -and ($trusted -notcontains $Request.LastModifiedByEmail.ToLowerInvariant())) {
        return [pscustomobject]@{ Allowed = $false; Reason = "Last edited by $($Request.LastModifiedByEmail), who isn't a trusted editor. Check what changed (version history) before running it." }
    }
    $immediateTerm = $Request.RequestType -eq 'Termination' -and ($Request.DisableImmediately -or $Request.TerminationType -eq 'Involuntary')
    if ((Get-ConfigValue $r 'RequireApproval') -and -not $immediateTerm -and -not $Request.ApprovedBy) {
        return [pscustomobject]@{ Allowed = $false; Reason = 'Marked Ready for IT without a recorded approval.' }
    }
    return [pscustomobject]@{ Allowed = $true; Reason = '' }
}

function Get-LifecycleRequestAction {
    <#
        Decides what IT should do for a request that's Ready for IT (or Scheduled):
          CreateUser | CreateContact | Offboard | RemoveContact | UpdateProfile | Wait | Review
    #>
    param([Parameter(Mandatory)]$Request, [Parameter(Mandatory)][datetime]$NowUtc, [Parameter(Mandatory)][hashtable]$Config)
    if ($NowUtc.Kind -eq [DateTimeKind]::Local) { $NowUtc = $NowUtc.ToUniversalTime() }
    $result = { param($a, $why, $runAt) [pscustomobject]@{ Action = $a; RunAtUtc = $runAt; Reason = $why } }
    $access = Get-LifecycleConfigEntry (Get-ConfigValue $Config 'AccessTypes') $Request.AccessType
    $isContact = [bool]($access -and (Get-ConfigValue $access 'Contact'))
    $cutoff = if ((Get-ConfigValue $Config 'Requests.TerminationCutoff')) { (Get-ConfigValue $Config 'Requests.TerminationCutoff') } else { '18:00' }

    switch ($Request.RequestType) {
        'New hire' {
            if (-not $Request.FirstName -or -not $Request.LastName) { return & $result 'Review' 'First and last name are required.' $null }
            if (-not $access) { return & $result 'Review' "Unknown computer access type '$($Request.AccessType)'." $null }
            if ($isContact) { return & $result 'CreateContact' 'Contact only - no account.' $null }
            if ((Get-ConfigValue $Config 'DirectoryMode') -eq 'Hybrid') { return & $result 'Review' 'Accounts are mastered in on-prem AD: create this one in AD.' $null }
            return & $result 'CreateUser' 'New account.' $null
        }
        'Termination' {
            $a = if ($isContact) { 'RemoveContact' } else { 'Offboard' }
            if ($Request.DisableImmediately -or $Request.TerminationType -eq 'Involuntary') { return & $result $a 'Immediate termination.' $null }
            if (-not $Request.LastDay) { return & $result $a 'No last day given - treated as immediate.' $null }
            $runAt = Get-SiteLocalTimeUtc -Date $Request.LastDay -Time $cutoff -Site $Request.Site -Config $Config
            if ($NowUtc -lt $runAt) { return & $result 'Wait' "Access ends at $cutoff on the last day ($($Request.LastDay.ToString('yyyy-MM-dd')))." $runAt }
            return & $result $a 'Last day has ended.' $runAt
        }
        'Change' {
            $runAt = $null
            if ($Request.EffectiveDate) {
                $runAt = Get-SiteLocalTimeUtc -Date $Request.EffectiveDate -Time '00:00' -Site $Request.Site -Config $Config
                if ($NowUtc -lt $runAt) { return & $result 'Wait' "Change takes effect on $($Request.EffectiveDate.ToString('yyyy-MM-dd'))." $runAt }
            }
            if ($isContact) { return & $result 'Review' 'Contact-only change - update the contact manually.' $runAt }
            return & $result 'UpdateProfile' 'Update title / department / manager.' $runAt
        }
        default { return & $result 'Review' "Unknown request type '$($Request.RequestType)'." $null }
    }
}

#endregion

#region Actions

function Invoke-LifecycleRequestAction {
    <#
        Carries out one decided action. Returns Status (Completed / Needs IT review), the
        account UPN when one was created, and the log lines written back to the request.
    #>
    param(
        [Parameter(Mandatory)]$Request,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)]$Index,
        [Parameter(Mandatory)][hashtable]$Config,
        [Collections.Generic.List[string]]$ExistingUpns = (New-Object Collections.Generic.List[string])
    )
    $log = New-Object Collections.Generic.List[string]
    $upn = $null
    $site = Get-LifecycleConfigEntry (Get-ConfigValue $Config 'Sites') $Request.Site
    $access = Get-LifecycleConfigEntry (Get-ConfigValue $Config 'AccessTypes') $Request.AccessType
    $office = if ($site -and (Get-ConfigValue $site 'OfficeLocation')) { (Get-ConfigValue $site 'OfficeLocation') } else { $Request.Site }
    $findUser = {
        param($email)
        if ($email -and $Index.ByEmail.ContainsKey($email.ToLowerInvariant())) { return $Index.ByEmail[$email.ToLowerInvariant()] }
        return $null
    }
    $finish = {
        $status = if (@($log | Where-Object { $_ -match '^(FAILED|REVIEW)' }).Count) { $script:Status.Review } else { $script:Status.Completed }
        [pscustomobject]@{ Status = $status; Upn = $upn; Log = $log.ToArray() }
    }

    switch ($Action) {
        'CreateUser' {
            # A re-run of a request whose account was already made (e.g. after an interrupted run).
            if ($Request.ITUpn -and $Index.ByEmail.ContainsKey($Request.ITUpn.ToLowerInvariant())) {
                $log.Add("REVIEW: $($Request.ITUpn) already exists from an earlier attempt at this request. Check it and mark the request Completed.")
                return & $finish
            }
            if ($Request.PaycomEmployeeId -and $Index.ByEmployeeId.ContainsKey($Request.PaycomEmployeeId)) {
                $existing = $Index.ByEmployeeId[$Request.PaycomEmployeeId]
                $state = if ($existing.accountEnabled) { 'enabled' } else { 'disabled' }
                $log.Add("REVIEW: Paycom employee $($Request.PaycomEmployeeId) already has an account ($($existing.userPrincipalName), $state). If this is a rehire, re-enable it and reset MFA.")
                return & $finish
            }
            $nameKey = Get-NameKey "$($Request.PreferredName)$($Request.LastName)"
            if ($Index.ByName.ContainsKey($nameKey)) {
                $same = ($Index.ByName[$nameKey] | ForEach-Object { $_.userPrincipalName }) -join ', '
                $log.Add("NOTE: Existing account(s) with the same name: $same. Confirm this is a different person, not a rehire.")
            }
            $upn = New-UpnCandidate -FirstName $Request.PreferredName -LastName $Request.LastName -Domain $Config.Onboarding.Domain -ExistingUpns $ExistingUpns
            $employee = [pscustomobject]@{
                EmployeeId = $Request.PaycomEmployeeId; DisplayName = $Request.DisplayName; PreferredName = $Request.PreferredName
                LastName = $Request.LastName; Department = $Request.Department; JobTitle = $Request.JobTitle
                Location = $office; HireDate = $Request.StartDate
            }
            $groups = @()
            if ($access -and (Get-ConfigValue $access 'GroupIds')) { $groups += @((Get-ConfigValue $access 'GroupIds')) }
            if ($site -and (Get-ConfigValue $site 'GroupIds')) { $groups += @((Get-ConfigValue $site 'GroupIds')) }
            $manager = & $findUser $Request.ManagerEmail
            if ($Request.ManagerEmail -and -not $manager) { $log.Add("NOTE: Manager $($Request.ManagerEmail) not found in Entra ID; set the manager manually.") }
            try {
                foreach ($line in (Invoke-LifecycleOnboarding -Employee $employee -Upn $upn -Config $Config -ManagerUser $manager -ExtraGroupIds $groups)) { $log.Add($line) }
                $ExistingUpns.Add($upn)
            }
            catch {
                $log.Add("FAILED: Create $upn - $($_.Exception.Message)")
                $upn = $null
            }
            return & $finish
        }

        'CreateContact' {
            if (-not $Request.PersonalEmail) {
                $log.Add('OK: No personal email given, so no address-book contact was created. This request is the record.')
                return & $finish
            }
            $id = $Request.PersonalEmail
            $existing = $null
            try { $existing = Invoke-LifecycleExchange 'Get-MailContact' @{ Identity = $id } } catch { $existing = $null }
            if ($existing) { $log.Add("OK: Contact for $id already exists.") }
            else {
                try {
                    Invoke-LifecycleExchange 'New-MailContact' @{
                        Name = $Request.DisplayName; DisplayName = $Request.DisplayName; FirstName = $Request.PreferredName
                        LastName = $Request.LastName; ExternalEmailAddress = $id
                    } | Out-Null
                    $log.Add("OK: Created contact $($Request.DisplayName) <$id>")
                }
                catch {
                    $log.Add("FAILED: Create contact - $($_.Exception.Message)")
                    return & $finish
                }
            }
            $contact = @{ Identity = $id }
            foreach ($pair in @(@('Department', $Request.Department), @('Title', $Request.JobTitle), @('Office', $office),
                    @('MobilePhone', $Request.MobilePhone), @('Company', (Get-ConfigValue $Config 'Onboarding.CompanyName')))) {
                if ($pair[1]) { $contact[$pair[0]] = $pair[1] }
            }
            try { Invoke-LifecycleExchange 'Set-Contact' $contact | Out-Null; $log.Add('OK: Contact details set') }
            catch { $log.Add("FAILED: Set contact details - $($_.Exception.Message)") }
            $attrs = @{ Identity = $id; CustomAttribute2 = "Lifecycle request $($Request.Id)" }
            if ($Request.PaycomEmployeeId) { $attrs.CustomAttribute1 = $Request.PaycomEmployeeId }
            try { Invoke-LifecycleExchange 'Set-MailContact' $attrs | Out-Null }
            catch { $log.Add("FAILED: Tag contact - $($_.Exception.Message)") }
            if ($site -and (Get-ConfigValue $site 'ContactGroups')) {
                foreach ($g in @((Get-ConfigValue $site 'ContactGroups'))) {
                    try { Invoke-LifecycleExchange 'Add-DistributionGroupMember' @{ Identity = $g; Member = $id } | Out-Null; $log.Add("OK: Added to $g") }
                    catch { $log.Add("FAILED: Add to $g - $($_.Exception.Message)") }
                }
            }
            return & $finish
        }

        'Offboard' {
            if (-not $Request.EmployeeEmail) { $log.Add('REVIEW: No employee selected on the request.'); return & $finish }
            $user = & $findUser $Request.EmployeeEmail
            if (-not $user) { $log.Add("REVIEW: No Entra account found for $($Request.EmployeeEmail)."); return & $finish }
            $upnLower = ([string]$user.userPrincipalName).ToLowerInvariant()
            if (@((Get-ConfigValue $Config 'Offboarding.ProtectedUpns') | ForEach-Object { ([string]$_).ToLowerInvariant() }) -contains $upnLower) {
                $log.Add("REVIEW: $($user.userPrincipalName) is on the protected list; offboard it by hand.")
                return & $finish
            }
            if (-not $user.accountEnabled) { $log.Add("OK: $($user.userPrincipalName) was already disabled.") }
            else { foreach ($line in (Invoke-LifecycleOffboarding -User $user -Config $Config)) { $log.Add($line) } }
            if ((Get-ConfigValue $Config 'Offboarding.ConvertMailboxToShared')) {
                try {
                    Invoke-LifecycleExchange 'Set-Mailbox' @{ Identity = $user.userPrincipalName; Type = 'Shared' } | Out-Null
                    $log.Add('OK: Mailbox converted to shared')
                }
                catch { $log.Add("FAILED: Convert mailbox to shared - $($_.Exception.Message)") }
                if ($Request.MailboxDelegateEmail) {
                    try {
                        Invoke-LifecycleExchange 'Add-MailboxPermission' @{
                            Identity = $user.userPrincipalName; User = $Request.MailboxDelegateEmail; AccessRights = 'FullAccess'
                            InheritanceType = 'All'; AutoMapping = $true
                        } | Out-Null
                        $log.Add("OK: Mailbox access given to $($Request.MailboxDelegateEmail)")
                    }
                    catch { $log.Add("FAILED: Mailbox access for $($Request.MailboxDelegateEmail) - $($_.Exception.Message)") }
                }
            }
            $upn = $user.userPrincipalName
            return & $finish
        }

        'RemoveContact' {
            $id = if ($Request.PersonalEmail) { $Request.PersonalEmail } else { $Request.DisplayName }
            if (-not $id) { $log.Add('REVIEW: No contact name or email on the request.'); return & $finish }
            $existing = $null
            try { $existing = Invoke-LifecycleExchange 'Get-MailContact' @{ Identity = $id } } catch { $existing = $null }
            if (-not $existing) { $log.Add("OK: No contact found for $id; nothing to remove."); return & $finish }
            try { Invoke-LifecycleExchange 'Remove-MailContact' @{ Identity = $id; Confirm = $false } | Out-Null; $log.Add("OK: Removed contact $id") }
            catch { $log.Add("FAILED: Remove contact $id - $($_.Exception.Message)") }
            return & $finish
        }

        'UpdateProfile' {
            $user = & $findUser $Request.EmployeeEmail
            if (-not $user) { $log.Add("REVIEW: No Entra account found for $($Request.EmployeeEmail)."); return & $finish }
            $body = @{}
            if ($Request.Department) { $body.department = $Request.Department }
            if ($Request.JobTitle) { $body.jobTitle = $Request.JobTitle }
            if ($Request.Site) { $body.officeLocation = $office }
            if ($Request.PaycomEmployeeId -and -not $user.employeeId) { $body.employeeId = $Request.PaycomEmployeeId }
            if ($body.Count) {
                try {
                    Invoke-LifecycleGraph -Method PATCH -Uri "/v1.0/users/$($user.id)" -Body $body | Out-Null
                    $log.Add("OK: Updated $((($body.Keys | Sort-Object) -join ', '))")
                }
                catch { $log.Add("FAILED: Update profile - $($_.Exception.Message)") }
            }
            if ($Request.ManagerEmail) {
                $manager = & $findUser $Request.ManagerEmail
                if (-not $manager) { $log.Add("REVIEW: Manager $($Request.ManagerEmail) not found in Entra ID.") }
                else {
                    try {
                        Invoke-LifecycleGraph -Method PUT -Uri "/v1.0/users/$($user.id)/manager/`$ref" -Body @{ '@odata.id' = "$(Get-LifecycleGraphBase)/v1.0/users/$($manager.id)" } | Out-Null
                        $log.Add("OK: Manager set to $($manager.userPrincipalName)")
                    }
                    catch { $log.Add("FAILED: Set manager - $($_.Exception.Message)") }
                }
            }
            if ($Request.Site) { $log.Add('NOTE: Site changed - review site groups and shared mailboxes (see the Desk365 ticket).') }
            $upn = $user.userPrincipalName
            return & $finish
        }

        default {
            $log.Add("REVIEW: No automated action for '$Action'.")
            return & $finish
        }
    }
}

function New-LifecycleRequestNotice {
    <# The completion email to the requester, manager and IT. #>
    param([Parameter(Mandatory)]$Request, [Parameter(Mandatory)]$Outcome, [Parameter(Mandatory)][string]$Action, [string]$ListUrl)
    $name = if ($Request.DisplayName) { $Request.DisplayName } elseif ($Outcome.Upn) { $Outcome.Upn } else { $Request.EmployeeEmail }
    $done = $Outcome.Status -eq $script:Status.Completed
    $what = switch ($Action) {
        'CreateUser' { if ($done) { "The account <b>$(ConvertTo-LifecycleHtml $Outcome.Upn)</b> is ready. On their first day IT will give $(ConvertTo-LifecycleHtml $name) a one-time sign-in code to set up their password and MFA. No password is sent by email." } else { 'The account could not be created automatically. IT has been notified.' } }
        'CreateContact' { "$(ConvertTo-LifecycleHtml $name) has been added to the company address book as a contact (no computer account)." }
        'Offboard' { "Access for $(ConvertTo-LifecycleHtml $name) has been turned off." }
        'RemoveContact' { "$(ConvertTo-LifecycleHtml $name) has been removed from the address book." }
        'UpdateProfile' { "The directory profile for $(ConvertTo-LifecycleHtml $name) has been updated." }
        default { 'IT needs to review this request.' }
    }
    $subject = if ($done) { "[IT] $($Request.RequestType) completed: $name" } else { "[IT] $($Request.RequestType) needs IT review: $name" }
    $lines = ($Outcome.Log | ForEach-Object { "<li>$(ConvertTo-LifecycleHtml $_)</li>" }) -join ''
    $link = if ($ListUrl) { "<p><a href='$ListUrl'>Open the request list</a></p>" } else { '' }
    $body = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222'><p>$what</p><p style='color:#555'>Request #$(ConvertTo-LifecycleHtml $Request.Id)</p><ul>$lines</ul>$link</div>"
    [pscustomobject]@{ Subject = $subject; Body = $body }
}

function ConvertTo-LifecycleHtml {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Test-LifecycleGroupMembership {
    # True if the user is a (transitive) member of any of the groups.
    param([Parameter(Mandatory)][string]$UserEmail, [Parameter(Mandatory)][string[]]$GroupIds)
    $res = Invoke-LifecycleGraph -Method POST -Uri "/v1.0/users/$([uri]::EscapeDataString($UserEmail))/checkMemberGroups" -Body @{ groupIds = @($GroupIds) }
    return @(Get-FieldValue $res 'value').Count -gt 0
}

#endregion

#region Paycom audit

function Compare-RosterToRequests {
    <#
        For the weekly Paycom audit: which hires and terminations Paycom shows that nobody
        filed a request for. Matching is on Paycom employee code, then the person's account
        (termination and change requests pick the account, not a typed name), then name.
    #>
    param([Parameter(Mandatory)]$Diff, [object[]]$Requests = @(), $Index)
    $live = @($Requests | Where-Object { @($script:Status.Rejected, $script:Status.Cancelled) -notcontains $_.Status })
    $match = {
        param($employee, [string]$type)
        $addresses = @($employee.Email)
        if ($Index) {
            $m = Find-DirectoryMatch $employee $Index
            if ($m) { $addresses += @($m.User.userPrincipalName, $m.User.mail) }
        }
        $addresses = @($addresses | Where-Object { $_ })
        foreach ($r in $live) {
            if ($r.RequestType -ne $type) { continue }
            if ($r.PaycomEmployeeId -and $r.PaycomEmployeeId -eq $employee.EmployeeId) { return $true }
            if ($r.EmployeeEmail -and $addresses -contains $r.EmployeeEmail) { return $true }
            $rk = Get-NameKey "$($r.PreferredName)$($r.LastName)"
            $rk2 = Get-NameKey "$($r.FirstName)$($r.LastName)"
            foreach ($ek in @((Get-NameKey "$($employee.PreferredName)$($employee.LastName)"), (Get-NameKey "$($employee.FirstName)$($employee.LastName)"))) {
                if ($ek -and ($ek -eq $rk -or $ek -eq $rk2)) { return $true }
            }
        }
        return $false
    }
    $gapHires = @($Diff.NewHires | Where-Object { -not (& $match $_ 'New hire') })
    $gapTerms = @($Diff.Terminations | Where-Object { -not (& $match $_.Employee 'Termination') })
    [pscustomobject]@{
        HiresWithoutRequest        = $gapHires
        TerminationsWithoutRequest = $gapTerms
        CoveredHireIds             = @($Diff.NewHires | Where-Object { $gapHires -notcontains $_ } | ForEach-Object EmployeeId)
        CoveredTerminationIds      = @($Diff.Terminations | Where-Object { $gapTerms -notcontains $_ } | ForEach-Object { $_.Employee.EmployeeId })
    }
}

#endregion

Export-ModuleMember -Function Get-FieldValue, ConvertTo-LifecycleBool, Get-LifecycleConfigEntry, Get-SiteTimeZoneId, Get-SiteLocalTimeUtc,
    ConvertFrom-ListDate, Set-LifecycleExchangeInvoker, Invoke-LifecycleExchange, Connect-LifecycleExchange, Test-RequestNeedsExchange,
    Get-LifecycleRequestListColumns, Resolve-LifecyclePersonEmail, ConvertFrom-LifecycleListItem, Get-LifecycleRequestItems,
    Update-LifecycleRequestItem, Add-LifecycleLogLines, Test-LifecycleRequestAllowed, Get-LifecycleRequestAction,
    Invoke-LifecycleRequestAction, New-LifecycleRequestNotice, Test-LifecycleGroupMembership, Compare-RosterToRequests
