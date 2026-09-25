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
    Reversed        = 'Reversed - did not start'
}
$script:RequestTypes = @('New hire', 'Termination', 'Change')
# "Did they start?" on a new hire. A no-show makes the automation undo what it set up.
$script:HireOutcome = @{
    Pending = 'Pending start'
    Started = 'Started'
    NoShow  = 'No-show / not starting'
}
# Fields the automation acts on. After approval, only trusted editors (the flow, HR) may change them.
$script:CriticalFields = @('RequestType', 'AccessType', 'FirstName', 'PreferredName', 'LastName', 'EmployeeLookupId', 'EmployeeEmail',
    'Site', 'Department', 'JobTitle', 'ManagerLookupId', 'ManagerEmail', 'StartDate', 'LastDay', 'EffectiveDate', 'TerminationType',
    'DisableImmediately', 'MailboxDelegateLookupId', 'MailboxDelegateEmail', 'PersonalEmail', 'PaycomEmployeeId')

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

function Find-LifecycleMailContact {
    <#
        Exact lookup of a mail contact by its external address. Never calls Get-MailContact
        with an empty identity (that returns every contact). Returns the single match, $null
        if there is none, and throws if the address matches more than one contact.
    #>
    param([string]$ExternalEmail)
    if (-not $ExternalEmail -or -not $ExternalEmail.Trim()) { throw 'No email address to look up.' }
    $escaped = $ExternalEmail.Trim() -replace "'", "''"
    $found = @(Invoke-LifecycleExchange 'Get-MailContact' @{ Filter = "ExternalEmailAddress -eq '$escaped'"; ResultSize = 2 })
    $found = @($found | Where-Object { $_ })
    if ($found.Count -gt 1) { throw "More than one contact has the address $ExternalEmail." }
    if ($found.Count -eq 1) { return $found[0] }
    return $null
}

function Test-RequestNeedsExchange {
    param([string]$Action, [hashtable]$Config)
    if (@('CreateContact', 'RemoveContact', 'ReverseHire') -contains $Action) { return $true }
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

    $yesNo = @('Yes', 'No')
    $status = & $choice 'Status' 'Status' @($script:Status.Submitted, $script:Status.PendingApproval, $script:Status.Ready, $script:Status.Scheduled,
        $script:Status.InProgress, $script:Status.Completed, $script:Status.Review, $script:Status.Rejected, $script:Status.Cancelled,
        $script:Status.Reversed)
    $status.indexed = $true
    $status.defaultValue = @{ value = $script:Status.Submitted }

    $requestType = & $choice 'RequestType' 'Request type' $script:RequestTypes
    $requestType.required = $true

    $outcome = & $choice 'HireOutcome' 'Did they start?' @($script:HireOutcome.Pending, $script:HireOutcome.Started, $script:HireOutcome.NoShow)
    $outcome.indexed = $true
    $outcome.defaultValue = @{ value = $script:HireOutcome.Pending }

    $dateTime = { param($name, $display) @{ name = $name; displayName = $display; dateTime = @{ format = 'dateTime'; displayAs = 'default' } } }
    $cfgList = { param($path) @((Get-ConfigValue $Config $path) | Where-Object { $_ }) }

    @(
        # ---- Every request (requester) ----
        $requestType
        $status
        & $choice 'AccessType' 'Computer access' @((Get-ConfigValue $Config 'AccessTypes') | ForEach-Object { $_.Name })
        & $text 'FirstName' 'Legal first name'
        & $text 'PreferredName' 'Preferred first name'
        & $text 'LastName' 'Last name'
        & $person 'Employee' 'Employee'
        & $choice 'Site' 'Location' @((Get-ConfigValue $Config 'Sites') | ForEach-Object { $_.Name }) $true
        & $choice 'Department' 'Department' @(Get-ConfigValue $Config 'Departments') $true
        & $text 'JobTitle' 'Job title'
        & $person 'Manager' 'Manager'
        & $multi 'Notes' 'Notes for HR / IT'

        # ---- New hire (requester) ----
        & $date 'StartDate' 'Start date (orientation day)'
        & $choice 'Equipment' 'Equipment needed' (& $cfgList 'Requests.EquipmentChoices') $false 'checkBoxes'
        & $text 'PersonalEmail' 'Personal email'
        & $text 'MobilePhone' 'Mobile phone'
        & $person 'Buddy' 'Buddy / ambassador'
        $outcome

        # ---- New hire (HR: New Employee Checklist) ----
        & $text 'ApplicantSource' 'Applicant source'
        & $text 'JobAdId' 'Job ad ID'
        & $choice 'HrNewHireChecklist' 'New Employee Checklist (HR)' (& $cfgList 'Requests.HrNewHireChecklist') $false 'checkBoxes'

        # ---- Termination (requester / General Manager: Separation Checklist) ----
        & $date 'TerminationDate' 'Termination date'
        & $date 'LastDay' 'Last day worked'
        & $choice 'TerminationType' 'Termination type' @('Voluntary', 'Involuntary')
        & $choice 'RehireEligible' 'Rehire eligible' $yesNo
        & $multi 'SeparationReason' 'Reason for separation'
        & $choice 'ProperNotice' 'Proper notice given' $yesNo
        & $choice 'ExitInterview' 'Exit interview' @('Accepts', 'Declines')
        & $choice 'OutgoingMedical' 'Outgoing medical testing' @('Accepts', 'Declines')
        & $text 'OutstandingEquipment' 'Outstanding equipment purchases'
        & $choice 'AmexAdvances' 'Outstanding advances / AMEX card collected' @('Yes (Notify Finance)', 'N/A')
        @{ name = 'DisableImmediately'; displayName = 'Disable access immediately'; boolean = @{} }
        & $person 'MailboxDelegate' 'Give mailbox and files to'
        # Collected on the last day (the General Manager, site admin or HR ticks these afterwards)
        & $choice 'BadgeCollected' 'ID badge collected' $yesNo
        & $choice 'RepairmanCertificate' 'Repairman certificate collected' @('Yes', 'No', 'N/A')
        & $choice 'ItemsReturned' 'Equipment / PPE returned' (& $cfgList 'Requests.ReturnItems') $false 'checkBoxes'

        # ---- Termination (HR / Payroll: Separation Checklist) ----
        & $date 'FinalPaycheckDate' 'Final paycheck date'
        & $choice 'PtoDue' 'PTO due' @('Yes', 'No', 'N/A')
        & $choice 'ExitInterviewDone' 'Exit interview completed' $yesNo
        & $choice 'BenefitsStatus' 'Benefits ending' (& $cfgList 'Requests.BenefitsChoices') $false 'checkBoxes'
        & $date 'BenefitsEndDate' 'Benefits date of termination'
        & $choice 'HrExitChecklist' 'HR exit checklist' (& $cfgList 'Requests.HrExitChecklist') $false 'checkBoxes'
        & $choice 'PayrollExitChecklist' 'Payroll exit checklist' (& $cfgList 'Requests.PayrollExitChecklist') $false 'checkBoxes'

        # ---- Change ----
        & $date 'EffectiveDate' 'Change effective date'

        # ---- HR / automation (hidden from the requester's form) ----
        & $text 'PaycomEmployeeId' 'Paycom employee code'
        & $text 'ApprovedBy' 'Approved by'
        & $text 'ITUpn' 'Account created'
        & $multi 'ITLog' 'IT automation log'
        & $dateTime 'ProcessedAt' 'Processed at'
        & $dateTime 'NotifiedAt' 'Employee status email sent'
        & $dateTime 'StartCheckSentAt' 'Start-day check sent'
    )
}

function New-LifecycleRequestList {
    <#
        Creates the request list on a site, makes sure Status is indexed, and makes the unused
        built-in Title column optional. Returns the created list (id, webUrl).
    #>
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$SiteId, [string]$DisplayName = 'Employee Lifecycle Requests')
    $body = @{ displayName = $DisplayName; list = @{ template = 'genericList' }; columns = @(Get-LifecycleRequestListColumns -Config $Config) }
    $list = Invoke-LifecycleGraph -Method POST -Uri "/v1.0/sites/$SiteId/lists" -Body $body
    $listId = Get-FieldValue $list 'id'
    $colUri = "/v1.0/sites/$SiteId/lists/$listId/columns"
    $cols = @(Get-FieldValue (Invoke-LifecycleGraph -Method GET -Uri "$colUri`?`$select=id,name,indexed,required") 'value')
    $byName = @{}
    foreach ($c in $cols) { $byName[[string](Get-FieldValue $c 'name')] = $c }
    $warnings = New-Object Collections.Generic.List[string]
    foreach ($indexed in 'Status', 'HireOutcome') {
        if ($byName.ContainsKey($indexed) -and -not (Get-FieldValue $byName[$indexed] 'indexed')) {
            try { Invoke-LifecycleGraph -Method PATCH -Uri "$colUri/$(Get-FieldValue $byName[$indexed] 'id')" -Body @{ indexed = $true } | Out-Null }
            catch { $warnings.Add("Couldn't index $indexed ($($_.Exception.Message)). Do it in List settings > Indexed columns.") }
        }
    }
    if ($byName.ContainsKey('Title')) {
        try { Invoke-LifecycleGraph -Method PATCH -Uri "$colUri/$(Get-FieldValue $byName['Title'] 'id')" -Body @{ required = $false } | Out-Null }
        catch { $warnings.Add("Couldn't make Title optional ($($_.Exception.Message)). Do it in List settings > Title > Require = No.") }
    }
    [pscustomobject]@{ Id = $listId; WebUrl = Get-FieldValue $list 'webUrl'; Warnings = $warnings.ToArray() }
}

#region Site settings list (badge office and start-day contacts, edited without touching config)

function Get-LifecycleSiteSettingsColumns {
    # Title holds the site code (AMA, FTW...), matching Sites[].Code in config.
    @(
        @{ name = 'SiteName'; displayName = 'Site'; description = 'For reference; the site code in Title is what counts.'; text = @{} }
        @{ name = 'BadgeOfficeEmails'; displayName = 'Badge office email(s)'
            description = 'Told about every termination at this site so the badge is returned. One address per line; can be outside IAC.'
            text = @{ allowMultipleLines = $true; textType = 'plain' } }
        @{ name = 'StartDayContacts'; displayName = 'Start-day email to'
            description = 'Asked "did everyone start?" on a new hire''s start date, with the hiring manager. One IAC address per line.'
            text = @{ allowMultipleLines = $true; textType = 'plain' } }
    )
}

function New-LifecycleSiteSettingsList {
    <#
        Creates the site settings list and seeds one row per site in config, with the addresses
        config holds today. Returns the list (id, webUrl).
    #>
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$SiteId, [string]$DisplayName = 'Lifecycle Site Settings')
    $body = @{ displayName = $DisplayName; list = @{ template = 'genericList' }; columns = @(Get-LifecycleSiteSettingsColumns) }
    $list = Invoke-LifecycleGraph -Method POST -Uri "/v1.0/sites/$SiteId/lists" -Body $body
    $listId = Get-FieldValue $list 'id'
    foreach ($s in @(Get-ConfigValue $Config 'Sites')) {
        $fields = @{
            Title             = [string](Get-ConfigValue $s 'Code')
            SiteName          = [string](Get-ConfigValue $s 'Name')
            BadgeOfficeEmails = (@(Get-ConfigValue $s 'BadgeOfficeEmails') | Where-Object { $_ }) -join "`n"
            StartDayContacts  = (@(Get-ConfigValue $s 'OrientationContacts') | Where-Object { $_ }) -join "`n"
        }
        Invoke-LifecycleGraph -Method POST -Uri "/v1.0/sites/$SiteId/lists/$listId/items" -Body @{ fields = $fields } | Out-Null
    }
    [pscustomobject]@{ Id = $listId; WebUrl = Get-FieldValue $list 'webUrl' }
}

function Get-LifecycleSiteSettingsItems {
    param([Parameter(Mandatory)][hashtable]$Config)
    $listId = Get-ConfigValue $Config 'Requests.SiteSettingsListId'
    if (-not $listId) { return }
    Invoke-LifecycleGraphPaged -Uri "/v1.0/sites/$($Config.Requests.SiteId)/lists/$listId/items?expand=fields&`$top=200"
}

function ConvertTo-LifecycleEmailList {
    # Splits a free-text cell (lines, commas or semicolons) into addresses; anything that isn't one goes to Invalid.
    param([string]$Text)
    $valid = New-Object Collections.Generic.List[string]; $invalid = New-Object Collections.Generic.List[string]
    foreach ($part in ([string]$Text -split '[\s;,]+')) {
        $p = $part.Trim().Trim('<', '>')
        if (-not $p) { continue }
        if ($p -match '^[^@\s<>"]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') { if (-not $valid.Contains($p.ToLowerInvariant())) { $valid.Add($p.ToLowerInvariant()) } }
        else { $invalid.Add($p) }
    }
    [pscustomobject]@{ Valid = $valid.ToArray(); Invalid = $invalid.ToArray() }
}

function Merge-LifecycleSiteSettings {
    <#
        Overlays the site settings list onto config.Sites and returns the merged Sites (copies;
        config itself is not changed) plus warnings for IT.
          - A row counts only if a trusted editor saved it last (Requests.SettingsEditors, or
            TrustedEditors if that isn't set). Otherwise config's values stay.
          - Start-day contacts get new hires' names and a link to the list, so they must be in
            Scope.Domains. Badge offices can be external (the airport).
          - An empty cell means "nobody", so people can be removed as well as added.
    #>
    param([Parameter(Mandatory)][hashtable]$Config, [object[]]$Items)
    $warnings = New-Object Collections.Generic.List[string]
    $editors = @(Get-ConfigValue $Config 'Requests.SettingsEditors')
    if (-not @($editors | Where-Object { $_ }).Count) { $editors = @(Get-ConfigValue $Config 'Requests.TrustedEditors') }
    $editors = @($editors | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $domains = @(Get-ConfigValue $Config 'Scope.Domains' | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })

    $rows = @{}
    foreach ($item in @($Items)) {
        if (-not $item) { continue }
        $f = Get-FieldValue $item 'fields'
        $code = ([string](Get-FieldValue $f 'Title')).Trim()
        if (-not $code) { continue }
        $editor = [string](Get-FieldValue (Get-FieldValue (Get-FieldValue $item 'lastModifiedBy') 'user') 'email')
        if (-not $editor -or $editors -notcontains $editor.ToLowerInvariant()) {
            $warnings.Add("Site settings for $code were last saved by '$editor', who isn't a settings editor; using config for that site until an editor saves the row.")
            continue
        }
        if ($rows.ContainsKey($code.ToLowerInvariant())) { $warnings.Add("Site settings has more than one row for $code; using the first."); continue }
        $rows[$code.ToLowerInvariant()] = $f
    }

    $known = @{}
    $sites = @(foreach ($s in @(Get-ConfigValue $Config 'Sites')) {
            $copy = @{}; foreach ($k in $s.Keys) { $copy[$k] = $s[$k] }
            $code = ([string](Get-ConfigValue $s 'Code')).ToLowerInvariant()
            $known[$code] = $true
            if ($code -and $rows.ContainsKey($code)) {
                $f = $rows[$code]
                $badge = ConvertTo-LifecycleEmailList (Get-FieldValue $f 'BadgeOfficeEmails')
                $start = ConvertTo-LifecycleEmailList (Get-FieldValue $f 'StartDayContacts')
                foreach ($bad in @($badge.Invalid) + @($start.Invalid)) { $warnings.Add("Site settings for $($s.Code): '$bad' isn't an email address; ignored.") }
                $internal = @($start.Valid | Where-Object { $domains -contains ($_ -split '@')[-1] })
                foreach ($ext in @($start.Valid | Where-Object { $internal -notcontains $_ })) {
                    $warnings.Add("Site settings for $($s.Code): start-day contact $ext is outside $($domains -join ', '); ignored.")
                }
                $copy['BadgeOfficeEmails'] = @($badge.Valid)
                $copy['OrientationContacts'] = $internal
            }
            $copy
        })
    foreach ($code in $rows.Keys) { if (-not $known.ContainsKey($code)) { $warnings.Add("Site settings has a row for '$code', which isn't a site in config; ignored.") } }
    [pscustomobject]@{ Sites = $sites; Warnings = $warnings.ToArray() }
}

#endregion

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
        HireOutcome          = $(if (& $get 'HireOutcome') { & $get 'HireOutcome' } else { $script:HireOutcome.Pending })
        TerminationDate      = ConvertFrom-ListDate (& $get 'TerminationDate') $tzId
        RehireEligible       = & $get 'RehireEligible'
        BuddyEmail           = & $person 'Buddy'
        NotifiedAt           = & $get 'NotifiedAt'
        StartCheckSentAt     = & $get 'StartCheckSentAt'
        Versions             = @(Get-FieldValue $Item 'versions')
        RequesterEmail       = $requester
        RequesterId          = Get-FieldValue (Get-FieldValue (Get-FieldValue $Item 'createdBy') 'user') 'id'
        CreatedAt            = $(if (Get-FieldValue $Item 'createdDateTime') { [DateTimeOffset]::Parse([string](Get-FieldValue $Item 'createdDateTime'), [Globalization.CultureInfo]::InvariantCulture).UtcDateTime } else { $null })
        LastModifiedByEmail  = Get-FieldValue $modifiedBy 'email'
    }
}

function Get-LifecycleRequestItems {
    <#
        Raw list items, optionally only those where an (indexed) field has one of the given
        values: one query per value, since list filters take a single indexed field.
    #>
    param([Parameter(Mandatory)][hashtable]$Config, [string[]]$Status, [string]$Field = 'Status', [string[]]$Values)
    $base = "/v1.0/sites/$($Config.Requests.SiteId)/lists/$($Config.Requests.ListId)/items?expand=fields&`$top=200"
    if ($Status) { $Values = $Status; $Field = 'Status' }
    if (-not $Values) { return @(Invoke-LifecycleGraphPaged -Uri $base) }
    $headers = @{ Prefer = 'HonorNonIndexedQueriesWarningMayFailRandomly' }
    foreach ($v in $Values) {
        $filter = [uri]::EscapeDataString("fields/$Field eq '$($v -replace "'", "''")'")
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
    $groups = @((Get-ConfigValue $r 'AuthorizedGroupIds') | Where-Object { $_ })
    # Checked when a request is first accepted; a scheduled termination still runs if the
    # requester has since left the team.
    if ($Request.Status -eq $script:Status.Ready) {
        if (-not $groups.Count) { return [pscustomobject]@{ Allowed = $false; Reason = 'Requests.AuthorizedGroupIds isn''t set, so the requester can''t be checked.' } }
        if (-not $IsMemberOf) { return [pscustomobject]@{ Allowed = $false; Reason = 'No way to check the requester''s group membership.' } }
        $who = if ($Request.RequesterId) { $Request.RequesterId } else { $Request.RequesterEmail }
        if (-not $who) { return [pscustomobject]@{ Allowed = $false; Reason = 'Requester unknown.' } }
        if (-not (& $IsMemberOf $who $groups)) {
            return [pscustomobject]@{ Allowed = $false; Reason = "Requester $($Request.RequesterEmail) is not in an authorised hiring group." }
        }
    }
    $immediateTerm = $Request.RequestType -eq 'Termination' -and ($Request.DisableImmediately -or $Request.TerminationType -eq 'Involuntary')
    if ((Get-ConfigValue $r 'RequireApproval') -and -not $immediateTerm -and -not $Request.ApprovedBy) {
        return [pscustomobject]@{ Allowed = $false; Reason = 'Marked Ready for IT without a recorded approval.' }
    }
    return [pscustomobject]@{ Allowed = $true; Reason = '' }
}

function Get-LifecycleRequestVersions {
    <# The item's version history (oldest first), with field values for each version. #>
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][string]$ItemId)
    @(Invoke-LifecycleGraphPaged -Uri "/v1.0/sites/$($Config.Requests.SiteId)/lists/$($Config.Requests.ListId)/items/$ItemId/versions?`$expand=fields")
}

function Test-LifecycleRequestProvenance {
    <#
        Checks the item's version history, which ordinary users can't alter:
          1. the change to 'Ready for IT' was made by a trusted editor (the flow's account or HR);
          2. nobody else changed a field the automation acts on after that.
        Requesters can still edit their own request afterwards (no-show, badge collected, etc.);
        only the fields in CriticalFields are protected. Edits by this automation carry no
        user email and never touch those fields.
    #>
    param([object[]]$Versions, [Parameter(Mandatory)][hashtable]$Config)
    $result = { param([bool]$ok, [string]$why) [pscustomobject]@{ Allowed = $ok; Reason = $why } }
    $trusted = @(Get-ConfigValue $Config 'Requests.TrustedEditors' | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
    if (-not $trusted.Count) { return & $result $false 'Requests.TrustedEditors is empty, so approvals can''t be verified.' }
    $versions = @($Versions | Where-Object { $_ })
    if (-not $versions.Count) { return & $result $false 'Couldn''t read the request''s version history, so its approval can''t be verified.' }
    if (@($versions | Where-Object { $null -eq (Get-FieldValue $_ 'fields') }).Count) { return & $result $false 'The version history came back without field values, so the approval can''t be verified.' }

    $when = {
        param($v)
        $t = Get-FieldValue $v 'lastModifiedDateTime'
        if ($t -is [datetime]) { return $t.ToUniversalTime() }
        if ($t) { return [DateTimeOffset]::Parse([string]$t, [Globalization.CultureInfo]::InvariantCulture).UtcDateTime }
        return [datetime]::MinValue
    }
    $versions = @($versions | Sort-Object { & $when $_ }, { [double]((Get-FieldValue $_ 'id') -replace '[^\d.]', '') })
    $field = { param($v, $n) $x = Get-FieldValue (Get-FieldValue $v 'fields') $n; if ($null -eq $x) { '' } elseif ($x -is [datetime]) { $x.ToUniversalTime().ToString('o') } else { [string]$x } }
    $editor = { param($v) ([string](Get-FieldValue (Get-FieldValue (Get-FieldValue $v 'lastModifiedBy') 'user') 'email')).ToLowerInvariant() }

    $approval = -1
    for ($i = 0; $i -lt $versions.Count; $i++) {
        $isReady = (& $field $versions[$i] 'Status') -eq $script:Status.Ready
        $wasReady = $i -gt 0 -and (& $field $versions[$i - 1] 'Status') -eq $script:Status.Ready
        if ($isReady -and -not $wasReady) { $approval = $i }
    }
    if ($approval -lt 0) { return & $result $false 'No change to Ready for IT found in the version history.' }
    $approver = & $editor $versions[$approval]
    if ($approval -eq 0) { return & $result $false "The request was created already marked Ready for IT (by $approver), skipping approval." }
    if ($trusted -notcontains $approver) {
        $who = if ($approver) { $approver } else { 'an app' }
        return & $result $false "It was set to Ready for IT by $who, who isn't a trusted approver."
    }
    for ($i = $approval + 1; $i -lt $versions.Count; $i++) {
        $changed = @($script:CriticalFields | Where-Object { (& $field $versions[$i - 1] $_) -ne (& $field $versions[$i] $_) })
        $by = & $editor $versions[$i]
        if ($changed.Count -and ($trusted -notcontains $by)) {
            $who = if ($by) { $by } else { 'an app or account with no email' }
            return & $result $false "Changed after approval by $who ($($changed -join ', ')). HR needs to re-approve."
        }
    }
    return & $result $true ''
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
            if ($Request.HireOutcome -eq $script:HireOutcome.NoShow) {
                if ($Request.Status -eq $script:Status.Completed) { return & $result 'ReverseHire' 'Marked as a no-show: undo what was set up.' $null }
                return & $result 'CancelHire' 'Marked as a no-show before anything was set up.' $null
            }
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
    $successStatus = switch ($Action) { 'ReverseHire' { $script:Status.Reversed } 'CancelHire' { $script:Status.Cancelled } default { $script:Status.Completed } }
    # Accounts the automation must not touch: protected, or outside the domains Paycom covers
    # (e.g. Eirtech etas.ie users, guests).
    $outOfBounds = {
        param($u)
        $upnL = ([string]$u.userPrincipalName).ToLowerInvariant()
        if (@((Get-ConfigValue $Config 'Offboarding.ProtectedUpns') | ForEach-Object { ([string]$_).ToLowerInvariant() }) -contains $upnL) {
            return "$($u.userPrincipalName) is on the protected list; handle it by hand."
        }
        $domains = @((Get-ConfigValue $Config 'Scope.Domains') | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
        if ($domains.Count -and ($domains -notcontains (($upnL -split '@')[-1]))) { return "$($u.userPrincipalName) isn't in a domain this automation manages." }
        if ((Get-FieldValue $u 'userType') -and (Get-FieldValue $u 'userType') -ne 'Member') { return "$($u.userPrincipalName) is a guest account." }
        return $null
    }
    $finish = {
        $status = if (@($log | Where-Object { $_ -match '^(FAILED|REVIEW)' }).Count) { $script:Status.Review } else { $successStatus }
        [pscustomobject]@{ Status = $status; Upn = $upn; Log = $log.ToArray() }
    }
    $removeContact = {
        param([bool]$MissingIsFine)
        # Only ever by exact external address: a name could match the wrong person.
        $id = $Request.PersonalEmail
        if (-not $id) { $log.Add('REVIEW: No personal email on the request, so the contact can''t be identified safely. Remove it by hand.'); return }
        try { $existing = Find-LifecycleMailContact $id }
        catch { $log.Add("REVIEW: $($_.Exception.Message)"); return }
        if (-not $existing) {
            if ($MissingIsFine) { $log.Add("OK: No contact found for $id; nothing to remove.") }
            else { $log.Add("REVIEW: No contact found for $id. Check the address; the person may still be in the address book.") }
            return
        }
        # Only contacts this automation created (tagged on creation), unless configured otherwise.
        $tag = [string](Get-FieldValue $existing 'CustomAttribute2')
        if ($tag -notmatch '^Lifecycle request' -and -not (Get-ConfigValue $Config 'Exchange.RemoveUntaggedContacts')) {
            $log.Add("REVIEW: The contact for $id wasn't created by this automation; remove it by hand if that's right.")
            return
        }
        $guid = [string](Get-FieldValue $existing 'Guid')
        if (-not $guid) { $log.Add("REVIEW: Contact for $id has no GUID; remove it by hand."); return }
        try { Invoke-LifecycleExchange 'Remove-MailContact' @{ Identity = $guid; Confirm = $false } | Out-Null; $log.Add("OK: Removed contact $id") }
        catch { $log.Add("FAILED: Remove contact $id - $($_.Exception.Message)") }
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
                # So a duplicate request later in the same run sees this account.
                $made = [pscustomobject]@{ id = $null; userPrincipalName = $upn; mail = $upn; employeeId = $Request.PaycomEmployeeId; accountEnabled = $true; displayName = $Request.DisplayName }
                $Index.ByEmail[$upn.ToLowerInvariant()] = $made
                if ($Request.PaycomEmployeeId) { $Index.ByEmployeeId[$Request.PaycomEmployeeId] = $made }
                if (-not $Index.ByName.ContainsKey($nameKey)) { $Index.ByName[$nameKey] = New-Object Collections.Generic.List[object] }
                $Index.ByName[$nameKey].Add($made)
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
            try { $existing = Find-LifecycleMailContact $id }
            catch { $log.Add("REVIEW: $($_.Exception.Message)"); return & $finish }
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
                    $add = @{ Identity = $g; Member = $id }
                    if (Get-ConfigValue $Config 'Exchange.BypassGroupOwnerCheck') { $add.BypassSecurityGroupManagerCheck = $true }
                    try { Invoke-LifecycleExchange 'Add-DistributionGroupMember' $add | Out-Null; $log.Add("OK: Added to $g") }
                    catch { $log.Add("FAILED: Add to $g - $($_.Exception.Message)") }
                }
            }
            return & $finish
        }

        'Offboard' {
            if (-not $Request.EmployeeEmail) { $log.Add('REVIEW: No employee selected on the request.'); return & $finish }
            $user = & $findUser $Request.EmployeeEmail
            if (-not $user) { $log.Add("REVIEW: No Entra account found for $($Request.EmployeeEmail)."); return & $finish }
            $why = & $outOfBounds $user
            if ($why) { $log.Add("REVIEW: $why"); return & $finish }

            # An immediate termination skips HR approval, so the requester must be the employee's
            # manager in Entra, or HR/IT. Otherwise HR approves it first.
            $immediate = $Request.DisableImmediately -or $Request.TerminationType -eq 'Involuntary'
            $humanApproved = $Request.ApprovedBy -and $Request.ApprovedBy -notmatch '^Immediate'
            $manager = $null
            try { $manager = Invoke-LifecycleGraph -Method GET -Uri "/v1.0/users/$($user.id)/manager?`$select=id,mail,userPrincipalName" } catch { $manager = $null }
            $managerAddrs = @(@((Get-FieldValue $manager 'mail'), (Get-FieldValue $manager 'userPrincipalName')) | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
            $trusted = @((Get-ConfigValue $Config 'Requests.TrustedEditors') | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
            $requester = ([string]$Request.RequesterEmail).ToLowerInvariant()
            if ($immediate -and -not $humanApproved -and ($managerAddrs -notcontains $requester) -and ($trusted -notcontains $requester)) {
                $log.Add("REVIEW: Immediate termination requested by $($Request.RequesterEmail), who isn't $($user.userPrincipalName)'s manager or HR. HR needs to approve it.")
                return & $finish
            }

            $blocked = $false
            if (-not $user.accountEnabled) { $log.Add("OK: $($user.userPrincipalName) was already disabled."); $blocked = $true }
            else {
                foreach ($line in (Invoke-LifecycleOffboarding -User $user -Config $Config)) { $log.Add($line) }
                $blocked = [bool]($log | Where-Object { $_ -match '^OK: (Block sign-in|Disable on-prem AD account)' })
            }
            if (Get-ConfigValue $Config 'Offboarding.ConvertMailboxToShared') {
                if (-not $blocked) {
                    $log.Add('REVIEW: Sign-in could not be blocked, so the mailbox was left alone.')
                }
                else {
                    try {
                        Invoke-LifecycleExchange 'Set-Mailbox' @{ Identity = $user.userPrincipalName; Type = 'Shared' } | Out-Null
                        $log.Add('OK: Mailbox converted to shared')
                    }
                    catch { $log.Add("FAILED: Convert mailbox to shared - $($_.Exception.Message)") }
                    $delegate = $Request.MailboxDelegateEmail
                    if ($delegate -and $immediate -and -not $humanApproved -and ($managerAddrs -notcontains $delegate.ToLowerInvariant())) {
                        $log.Add("NOTE: Mailbox access for $delegate wasn't granted automatically (immediate termination, and they aren't the employee's manager). HR can grant it after review.")
                        $delegate = $null
                    }
                    if ($delegate) {
                        try {
                            Invoke-LifecycleExchange 'Add-MailboxPermission' @{
                                Identity = $user.userPrincipalName; User = $delegate; AccessRights = 'FullAccess'
                                InheritanceType = 'All'; AutoMapping = $true
                            } | Out-Null
                            $log.Add("OK: Mailbox access given to $delegate")
                        }
                        catch { $log.Add("FAILED: Mailbox access for $delegate - $($_.Exception.Message)") }
                    }
                }
            }
            $upn = $user.userPrincipalName
            return & $finish
        }

        'RemoveContact' {
            & $removeContact $false
            return & $finish
        }

        'CancelHire' {
            $log.Add('OK: Marked as a no-show before anything was set up; nothing to undo.')
            return & $finish
        }

        'ReverseHire' {
            # Undo a hire who never started: the account (or contact) this request created.
            $isContact = [bool](Get-ConfigValue $access 'Contact')
            if ($isContact) { & $removeContact $true; return & $finish }
            if (-not $Request.ITUpn) { $log.Add('OK: No account had been created for this request; nothing to undo.'); return & $finish }
            $user = & $findUser $Request.ITUpn
            if (-not $user) { $log.Add("OK: $($Request.ITUpn) no longer exists; nothing to undo."); return & $finish }
            $why = & $outOfBounds $user
            if ($why) { $log.Add("REVIEW: $why"); return & $finish }
            # If the account has been used, they may have started after all.
            try {
                $activity = Get-FieldValue (Invoke-LifecycleGraph -Method GET -Uri "/v1.0/users/$($user.id)?`$select=signInActivity") 'signInActivity'
                $lastOk = Get-FieldValue $activity 'lastSuccessfulSignInDateTime'
                if ($lastOk) {
                    $log.Add("REVIEW: $($user.userPrincipalName) signed in successfully on $(([datetime]$lastOk).ToString('yyyy-MM-dd HH:mm')) UTC. Check they really didn't start before removing access.")
                    return & $finish
                }
            }
            catch { $log.Add("NOTE: Couldn't check sign-in activity ($($_.Exception.Message)); continuing.") }
            foreach ($line in (Invoke-LifecycleOffboarding -User $user -Config $Config -IncludeLicenceGroups)) { $log.Add($line) }
            if (Get-ConfigValue $Config 'Requests.NoShow.DeleteAccount') {
                try {
                    Invoke-LifecycleGraph -Method DELETE -Uri "/v1.0/users/$($user.id)" | Out-Null
                    $log.Add("OK: Deleted $($user.userPrincipalName) (restorable from Deleted users for 30 days)")
                }
                catch { $log.Add("FAILED: Delete account - $($_.Exception.Message)") }
            }
            $upn = $user.userPrincipalName
            return & $finish
        }

        'UpdateProfile' {
            $user = & $findUser $Request.EmployeeEmail
            if (-not $user) { $log.Add("REVIEW: No Entra account found for $($Request.EmployeeEmail)."); return & $finish }
            $why = & $outOfBounds $user
            if ($why) { $log.Add("REVIEW: $why"); return & $finish }
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
            $log.Add('NOTE: Licence / computer access changes are made by IT from the ticket, not automatically.')
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
    param([Parameter(Mandatory)]$Request, [Parameter(Mandatory)]$Outcome, [Parameter(Mandatory)][string]$Action, [string]$ListUrl, $User)
    # Terminations and changes act on the picked Employee; name them from the directory, not
    # from typed name fields that may be left over from another request type.
    $name = if ($User -and @('Offboard', 'UpdateProfile') -contains $Action) { Get-FieldValue $User 'displayName' }
    elseif ($Request.DisplayName) { $Request.DisplayName } elseif ($Outcome.Upn) { $Outcome.Upn } else { $Request.EmployeeEmail }
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

function Get-LifecycleSiteCode {
    param([string]$Site, [hashtable]$Config)
    $entry = Get-LifecycleConfigEntry (Get-ConfigValue $Config 'Sites') $Site
    $code = Get-ConfigValue $entry 'Code'
    if ($code) { return $code }
    if ($Site -match '\(([A-Z]{2,5})\)') { return $Matches[1] }
    return $Site
}

function New-LifecycleStatusNotice {
    <#
        The "Employee Status Notification" email HR used to send by hand, in the same format:
        subject "<site code> New Hire" / "Term" / "Change" / "No-show", with a table of
        First Name | Last Name | Title | Effective Date | Location.
    #>
    param([Parameter(Mandatory)]$Request, [Parameter(Mandatory)][ValidateSet('New Hire', 'Term', 'Change', 'No-show')][string]$Kind,
        [Parameter(Mandatory)][hashtable]$Config, $User)
    # For a picked employee, the directory is the truth: typed name fields may be left over.
    $fromUser = [bool]$User -and @('Term', 'Change') -contains $Kind
    $first = if ($fromUser) { Get-FieldValue $User 'givenName' } elseif ($Request.PreferredName) { $Request.PreferredName } else { '' }
    $last = if ($fromUser) { Get-FieldValue $User 'surname' } elseif ($Request.LastName) { $Request.LastName } else { '' }
    $title = if ($Kind -eq 'Change' -and $Request.JobTitle) { $Request.JobTitle } elseif ($fromUser) { Get-FieldValue $User 'jobTitle' } elseif ($Request.JobTitle) { $Request.JobTitle } else { '' }
    $effective = switch ($Kind) {
        'New Hire' { $Request.StartDate }
        'No-show' { $Request.StartDate }
        'Term' { if ($Request.TerminationDate) { $Request.TerminationDate } else { $Request.LastDay } }
        'Change' { $Request.EffectiveDate }
    }
    $effectiveText = if ($effective) { $effective.ToString('MM/dd/yyyy') } else { '' }
    $code = Get-LifecycleSiteCode $Request.Site $Config
    $cell = "style='border:1px solid #999;padding:4px 10px'"
    $row = @($first, $last, $title, $effectiveText, $Request.Site) | ForEach-Object { "<td $cell>$(ConvertTo-LifecycleHtml $_)</td>" }
    $head = @('First Name', 'Last Name', 'Title', 'Effective Date', 'Location') | ForEach-Object { "<th $cell>$_</th>" }
    $extra = switch ($Kind) {
        'Term' { "<p>Last day worked: $(if ($Request.LastDay) { $Request.LastDay.ToString('MM/dd/yyyy') }). Termination type: $(ConvertTo-LifecycleHtml $Request.TerminationType).</p>" }
        'No-show' { '<p>This person did not start. Their IT access has been removed. Please reverse the hire in Paycom.</p>' }
        default { '' }
    }
    [pscustomobject]@{
        Subject = "$code $Kind"
        Body    = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px'><table style='border-collapse:collapse'><tr>$($head -join '')</tr><tr>$($row -join '')</tr></table>$extra<p style='color:#666'>Sent automatically from lifecycle request #$(ConvertTo-LifecycleHtml $Request.Id).</p></div>"
    }
}

function Get-LifecycleStartCheckBatches {
    <#
        New hires whose start date has arrived (in their site's time zone, after the check
        time) and who haven't been asked about yet, grouped by site: "did these people start?"
    #>
    param([object[]]$Requests, [Parameter(Mandatory)][datetime]$NowUtc, [Parameter(Mandatory)][hashtable]$Config)
    if ($NowUtc.Kind -eq [DateTimeKind]::Local) { $NowUtc = $NowUtc.ToUniversalTime() }
    $checkTime = [timespan]::Parse($(if (Get-ConfigValue $Config 'Requests.NoShow.CheckTime') { Get-ConfigValue $Config 'Requests.NoShow.CheckTime' } else { '10:00' }))
    $due = foreach ($r in @($Requests)) {
        if ($r.RequestType -ne 'New hire' -or $r.Status -ne $script:Status.Completed) { continue }
        if ($r.HireOutcome -ne $script:HireOutcome.Pending -or $r.StartCheckSentAt -or -not $r.StartDate) { continue }
        $local = [TimeZoneInfo]::ConvertTimeFromUtc($NowUtc, [TimeZoneInfo]::FindSystemTimeZoneById((Get-SiteTimeZoneId $r.Site $Config)))
        if ($r.StartDate.Date -lt $local.Date -or ($r.StartDate.Date -eq $local.Date -and $local.TimeOfDay -ge $checkTime)) { $r }
    }
    @($due | Group-Object Site | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Site = $_.Name; Requests = @($_.Group | Sort-Object LastName, PreferredName) } })
}

function New-LifecycleStartCheckNotice {
    <# One email per site: who was due to start, and how to mark anyone who didn't. #>
    param([Parameter(Mandatory)][string]$Site, [Parameter(Mandatory)][object[]]$Requests, [string]$ListUrl)
    $cell = "style='border-bottom:1px solid #ddd;padding:4px 10px;text-align:left'"
    $rows = foreach ($r in $Requests) {
        $link = if ($ListUrl) { "<a href='$ListUrl/EditForm.aspx?ID=$([uri]::EscapeDataString($r.Id))'>Mark no-show</a>" } else { "Request #$(ConvertTo-LifecycleHtml $r.Id)" }
        $start = if ($r.StartDate) { $r.StartDate.ToString('ddd MM/dd') } else { '' }
        "<tr><td $cell>$(ConvertTo-LifecycleHtml $r.DisplayName)</td><td $cell>$(ConvertTo-LifecycleHtml $r.JobTitle)</td><td $cell>$start</td><td $cell>$(ConvertTo-LifecycleHtml $r.ITUpn)</td><td $cell>$link</td></tr>"
    }
    $n = @($Requests).Count
    $who = if ($n -eq 1) { '1 person was' } else { "$n people were" }
    [pscustomobject]@{
        Subject = "Did everyone start? $who due to start at $Site"
        Body    = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222'><p>$who due to start at <b>$(ConvertTo-LifecycleHtml $Site)</b>. If everyone showed up, there's nothing to do.</p><p>For anyone who <b>didn't start</b>, open their request and set <b>Did they start?</b> to <b>No-show / not starting</b>. IT's automation then removes the access that was set up and lets HR know to reverse the hire in Paycom.</p><table style='border-collapse:collapse;font-size:13px'><tr><th $cell>Name</th><th $cell>Title</th><th $cell>Start</th><th $cell>Account</th><th $cell></th></tr>$($rows -join '')</table></div>"
    }
}

function Test-LifecycleGroupMembership {
    # True if the user (object ID, or UPN) is a transitive member of any of the groups.
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
    param([Parameter(Mandatory)]$Diff, [object[]]$Requests = @(), $Index, [datetime]$NowUtc = [datetime]::UtcNow, [int]$StaleDays = 14)
    # A request counts as coverage if it went ahead or is on its way; one stuck waiting for
    # approval for more than two weeks doesn't.
    $live = @($Requests | Where-Object {
            (@($script:Status.Rejected, $script:Status.Cancelled, $script:Status.Reversed) -notcontains $_.Status) -and
            -not (@($script:Status.Submitted, $script:Status.PendingApproval) -contains $_.Status -and $_.CreatedAt -and ($NowUtc - $_.CreatedAt).TotalDays -gt $StaleDays)
        })
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
            if ($r.ITUpn -and $addresses -contains $r.ITUpn) { return $true }
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

Export-ModuleMember -Function Get-FieldValue, Get-LifecycleSiteSettingsColumns, New-LifecycleSiteSettingsList, Get-LifecycleSiteSettingsItems, ConvertTo-LifecycleEmailList, Merge-LifecycleSiteSettings, ConvertTo-LifecycleBool, Get-LifecycleConfigEntry, Get-SiteTimeZoneId, Get-SiteLocalTimeUtc,
    ConvertFrom-ListDate, Set-LifecycleExchangeInvoker, Invoke-LifecycleExchange, Connect-LifecycleExchange, Find-LifecycleMailContact, Test-RequestNeedsExchange,
    Get-LifecycleRequestListColumns, New-LifecycleRequestList, Resolve-LifecyclePersonEmail, ConvertFrom-LifecycleListItem, Get-LifecycleRequestItems,
    Update-LifecycleRequestItem, Add-LifecycleLogLines, Test-LifecycleRequestAllowed, Get-LifecycleRequestAction,
    Invoke-LifecycleRequestAction, New-LifecycleRequestNotice, Test-LifecycleGroupMembership, Compare-RosterToRequests,
    Get-LifecycleRequestVersions, Test-LifecycleRequestProvenance, Get-LifecycleSiteCode, New-LifecycleStatusNotice,
    Get-LifecycleStartCheckBatches, New-LifecycleStartCheckNotice
