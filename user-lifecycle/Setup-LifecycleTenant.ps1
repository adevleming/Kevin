#Requires -Version 5.1
<#
.SYNOPSIS
    One-time tenant setup for the lifecycle request process. Creates new objects only.

.DESCRIPTION
    Run this on your own PC or on the server that will run the automation, signed in as an
    admin. It creates:

      1. A private Microsoft 365 group + Teams team ("Hiring & Staffing Requests") with the
         owners and members from config. Its SharePoint site hosts everything below.
      2. The "Employee Lifecycle Requests" list on that site (all columns, Status indexed).
      4. The "IT Lifecycle Automation" app registration with a certificate, and its service
         principal. Permissions are requested but NOT consented; you click Grant admin
         consent yourself.
      5. A 'write' grant for that app on the new site only (Sites.Selected).
      6. With -IncludeExchange: the it-automation@ shared mailbox, and an Exchange RBAC for
         Applications assignment so the app can send mail as that one mailbox only.

    It never modifies or deletes anything that already exists. If an object with the same
    name is already there and wasn't created by this script (per the manifest), it stops.
    Every created object's ID is written to state/tenant-setup.json, so a re-run picks up where
    it left off and you know exactly what to remove if you ever want to.

    -WhatIf signs in, runs every read-only check (names free, people resolve) and prints what it
    would create, without creating anything.

.EXAMPLE
    ./Setup-LifecycleTenant.ps1 -ConfigPath ./config.psd1 -WhatIf
    ./Setup-LifecycleTenant.ps1 -ConfigPath ./config.psd1 -IncludeExchange
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    # Also create the sender mailbox and the Exchange RBAC for Applications scope (needs ExchangeOnlineManagement).
    [switch]$IncludeExchange,
    # Use an existing certificate instead of creating one.
    [string]$CertificateThumbprint,
    # Where a new certificate goes. Use Cert:\LocalMachine\My (elevated) if a service account runs the task.
    [string]$CertStoreLocation = 'Cert:\CurrentUser\My'
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'PaycomLifecycle.psm1')
Import-Module (Join-Path $PSScriptRoot 'LifecycleRequests.psm1')

$config = Import-PowerShellDataFile -Path $ConfigPath
$setup = $config.Setup
if (-not $setup) { throw "config has no Setup section. Copy it from config.example.psd1." }
$configDir = Split-Path -Parent (Resolve-Path $ConfigPath)
$statePath = $config.StatePath
if (-not [IO.Path]::IsPathRooted($statePath)) { $statePath = Join-Path $configDir $statePath }
if (-not (Test-Path $statePath)) { New-Item -ItemType Directory -Path $statePath -Force | Out-Null }
$manifestPath = Join-Path $statePath 'tenant-setup.json'
$manifest = @{}
if (Test-Path $manifestPath) {
    $loaded = Get-Content $manifestPath -Raw | ConvertFrom-Json
    foreach ($p in $loaded.PSObject.Properties) { $manifest[$p.Name] = $p.Value }
}
$dryRun = [bool]$WhatIfPreference
function Save-Manifest {
    if ($dryRun) { return }
    $manifest.updated = (Get-Date).ToString('s')
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $manifestPath -Encoding UTF8
}
function Step([string]$Text) { Write-Host ''; Write-Host "== $Text" -ForegroundColor Cyan }
function Plan([string]$Text) { Write-Host "   would create: $Text" -ForegroundColor Yellow }
function Done([string]$Text) { Write-Host "   $Text" -ForegroundColor Green }
function Invoke-WithRetry {
    # New groups, teams and sites take a while to replicate; Graph returns 404 until then.
    param([scriptblock]$Action, [int]$Attempts = 30, [int]$DelaySeconds = 20, [string]$What = 'resource')
    for ($i = 1; $i -le $Attempts; $i++) {
        try { return & $Action }
        catch {
            if ($i -eq $Attempts) { throw }
            Write-Host "   waiting for $What to be ready ($i/$Attempts): $($_.Exception.Message)" -ForegroundColor DarkGray
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}
$graphBase = 'https://graph.microsoft.com'

#region Sign in (delegated, as you)
Step 'Sign in to Microsoft Graph'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$scopes = @('User.Read.All', 'Group.ReadWrite.All', 'Team.Create', 'Sites.FullControl.All', 'Application.ReadWrite.All')
$connect = @{ Scopes = $scopes; NoWelcome = $true }
if ($config.Graph.TenantId -and $config.Graph.TenantId -notmatch '^0{8}-') { $connect.TenantId = $config.Graph.TenantId }
Connect-MgGraph @connect | Out-Null
$ctx = Get-MgContext
Done "Signed in as $($ctx.Account) to tenant $($ctx.TenantId)"
$tenantId = $ctx.TenantId
#endregion

#region Preflight (read-only)
Step 'Check people and names'
$people = @{}
$missing = New-Object Collections.Generic.List[string]
foreach ($addr in (@($setup.Owners) + @($setup.Members) | Where-Object { $_ } | Select-Object -Unique)) {
    try {
        $u = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/users/$([uri]::EscapeDataString($addr))?`$select=id,displayName,userPrincipalName,accountEnabled"
        $people[$addr.ToLowerInvariant()] = $u
        Done "$addr -> $($u.displayName)"
    }
    catch { $missing.Add($addr) }
}
if ($missing.Count) { throw "These addresses don't match an account: $($missing -join ', '). Fix Setup.Owners / Setup.Members and re-run." }
if (@($setup.Owners).Count -eq 0) { throw 'Setup.Owners is empty; the team needs at least one owner.' }

$nick = $setup.MailNickname
$existingGroup = @((Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groups?`$filter=mailNickname eq '$nick'&`$select=id,displayName").value)
if ($existingGroup.Count -and $existingGroup[0].id -ne $manifest.groupId) {
    throw "A group with mail nickname '$nick' already exists ($($existingGroup[0].displayName)) and wasn't created by this script. Pick another Setup.MailNickname; nothing was changed."
}
$existingApp = @((Invoke-MgGraphRequest -Method GET -Uri "/v1.0/applications?`$filter=displayName eq '$($setup.AppName -replace "'", "''")'&`$select=id,appId").value)
if ($existingApp.Count -and $existingApp[0].id -ne $manifest.appObjectId) {
    throw "An app registration named '$($setup.AppName)' already exists and wasn't created by this script. Pick another Setup.AppName; nothing was changed."
}
Done 'Names are free (or were created by an earlier run of this script)'
#endregion

#region 1. Group and team
Step "Team '$($setup.TeamName)'"
if ($manifest.groupId) { Done "Already created: group $($manifest.groupId)" }
elseif ($dryRun) { Plan "private Microsoft 365 group + team '$($setup.TeamName)' with $(@($setup.Owners).Count) owners and $(@($setup.Members).Count) members" }
else {
    $ownerIds = @($setup.Owners | ForEach-Object { $people[$_.ToLowerInvariant()].id })
    $memberIds = @((@($setup.Owners) + @($setup.Members)) | Where-Object { $_ } | Select-Object -Unique | ForEach-Object { $people[$_.ToLowerInvariant()].id })
    $firstBatch = @($memberIds | Select-Object -First ([math]::Max(0, 20 - $ownerIds.Count)))
    $group = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/groups' -Body (@{
            displayName             = $setup.TeamName
            description             = $setup.TeamDescription
            mailNickname            = $nick
            groupTypes              = @('Unified')
            mailEnabled             = $true
            securityEnabled         = $false
            visibility              = 'Private'
            resourceBehaviorOptions = @('WelcomeEmailDisabled')
            'owners@odata.bind'     = @($ownerIds | ForEach-Object { "$graphBase/v1.0/users/$_" })
            'members@odata.bind'    = @($firstBatch | ForEach-Object { "$graphBase/v1.0/users/$_" })
        } | ConvertTo-Json -Depth 5) -ContentType 'application/json'
    $manifest.groupId = $group.id; Save-Manifest
    foreach ($id in ($memberIds | Where-Object { $firstBatch -notcontains $_ })) {
        Invoke-MgGraphRequest -Method POST -Uri "/v1.0/groups/$($group.id)/members/`$ref" -Body (@{ '@odata.id' = "$graphBase/v1.0/directoryObjects/$id" } | ConvertTo-Json) -ContentType 'application/json' | Out-Null
    }
    Done "Created group $($group.id)"
}
if ($manifest.groupId -and -not $manifest.teamCreated -and -not $dryRun) {
    Invoke-WithRetry -What 'the new group' -Action {
        Invoke-MgGraphRequest -Method PUT -Uri "/v1.0/groups/$($manifest.groupId)/team" -Body (@{
                memberSettings = @{ allowCreateUpdateChannels = $false; allowDeleteChannels = $false; allowAddRemoveApps = $false; allowCreateUpdateRemoveTabs = $false }
            } | ConvertTo-Json -Depth 3) -ContentType 'application/json' | Out-Null
    }
    $manifest.teamCreated = $true; Save-Manifest
    Done 'Team created (members can use it but not add channels, tabs or apps)'
}
#endregion

#region 2. Site and list
Step 'SharePoint site and request list'
if (-not $manifest.siteId -and -not $dryRun) {
    $site = Invoke-WithRetry -What 'the team site' -Action { Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groups/$($manifest.groupId)/sites/root?`$select=id,webUrl" }
    $manifest.siteId = $site.id; $manifest.siteUrl = $site.webUrl; Save-Manifest
}
if ($manifest.siteId) { Done "Site: $($manifest.siteUrl)" }

if ($manifest.listId) { Done "Already created: list $($manifest.listId)" }
elseif ($dryRun) { Plan "list 'Employee Lifecycle Requests' with $(@(Get-LifecycleRequestListColumns -Config $config).Count) columns on the team site" }
else {
    # The module's Graph helper uses Invoke-MgGraphRequest, which is signed in as you right now.
    $list = Invoke-WithRetry -What 'the site to accept lists' -Attempts 10 -Action { New-LifecycleRequestList -Config $config -SiteId $manifest.siteId }
    $manifest.listId = $list.Id; $manifest.listUrl = $list.WebUrl; Save-Manifest
    foreach ($w in $list.Warnings) { Write-Warning $w }
    Done "Created list: $($list.WebUrl)"
}

# The weekly Paycom roster is NOT kept on this site: every hiring manager can read the team's
# files, and a replaced CSV could fake terminations. Keep it in an IT-only folder (Input.Path).
#endregion

#region 3. App registration, certificate, service principal
Step "App registration '$($setup.AppName)'"
$graphSp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')?`$select=appRoles"
$exoSp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/servicePrincipals(appId='00000002-0000-0ff1-ce00-000000000000')?`$select=appRoles"
$roleId = {
    param($sp, [string]$value)
    $r = @($sp.appRoles | Where-Object { $_.value -eq $value })
    if (-not $r.Count) { throw "Permission '$value' not found in this tenant." }
    @{ id = $r[0].id; type = 'Role' }
}
# Mail.Send is deliberately NOT requested: RBAC for Applications scopes it to one mailbox instead.
$graphRoles = @('User.ReadWrite.All', 'User.RevokeSessions.All', 'GroupMember.ReadWrite.All', 'Sites.Selected', 'AuditLog.Read.All')
$requiredAccess = @(
    @{ resourceAppId = '00000003-0000-0000-c000-000000000000'; resourceAccess = @($graphRoles | ForEach-Object { & $roleId $graphSp $_ }) }
    @{ resourceAppId = '00000002-0000-0ff1-ce00-000000000000'; resourceAccess = @(& $roleId $exoSp 'Exchange.ManageAsApp') }
)

if ($manifest.appObjectId) { Done "Already created: app $($manifest.appId)" }
elseif ($dryRun) {
    Plan "app registration '$($setup.AppName)' requesting (not consented): $($graphRoles -join ', '), Exchange.ManageAsApp"
    if (-not $CertificateThumbprint) { Plan "self-signed certificate 'CN=$($setup.AppName)' in $CertStoreLocation (2 years)" }
}
else {
    if ($CertificateThumbprint) { $cert = Get-Item (Join-Path $CertStoreLocation $CertificateThumbprint) }
    else {
        # KeyExchange / CSP key so the same certificate works for Exchange Online app-only sign-in.
        $cert = New-SelfSignedCertificate -Subject "CN=$($setup.AppName)" -CertStoreLocation $CertStoreLocation -KeyExportPolicy Exportable `
            -KeySpec KeyExchange -KeyLength 2048 -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2) `
            -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'
        Done "Created certificate $($cert.Thumbprint) in $CertStoreLocation (expires $($cert.NotAfter.ToString('yyyy-MM-dd')))"
    }
    $app = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/applications' -Body (@{
            displayName            = $setup.AppName
            signInAudience         = 'AzureADMyOrg'
            notes                  = "Created by user-lifecycle Setup-LifecycleTenant.ps1 on $((Get-Date).ToString('yyyy-MM-dd')) by $($ctx.Account)."
            requiredResourceAccess = $requiredAccess
            keyCredentials         = @(@{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($cert.RawData); displayName = "CN=$($setup.AppName)" })
        } | ConvertTo-Json -Depth 6) -ContentType 'application/json'
    $manifest.appObjectId = $app.id; $manifest.appId = $app.appId; $manifest.certThumbprint = $cert.Thumbprint; Save-Manifest
    Done "Created app $($app.appId)"
}
if ($manifest.appId -and -not $manifest.spObjectId -and -not $dryRun) {
    $sp = Invoke-WithRetry -What 'the new app' -Attempts 10 -DelaySeconds 10 -Action {
        Invoke-MgGraphRequest -Method POST -Uri '/v1.0/servicePrincipals' -Body (@{ appId = $manifest.appId } | ConvertTo-Json) -ContentType 'application/json'
    }
    $manifest.spObjectId = $sp.id; Save-Manifest
    Done "Created service principal $($sp.id)"
}
#endregion

#region 4. Site grant (Sites.Selected)
Step 'Give the app write access to the new site only'
if ($manifest.sitePermissionId) { Done 'Already granted' }
elseif ($dryRun) { Plan "Sites.Selected 'write' grant for '$($setup.AppName)' on the team site" }
else {
    $perm = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/sites/$($manifest.siteId)/permissions" -Body (@{
            roles               = @('write')
            grantedToIdentities = @(@{ application = @{ id = $manifest.appId; displayName = $setup.AppName } })
        } | ConvertTo-Json -Depth 5) -ContentType 'application/json'
    $manifest.sitePermissionId = $perm.id; Save-Manifest
    Done 'Granted write on the team site'
}
#endregion

#region 5. Exchange: sender mailbox + RBAC for Applications
if ($IncludeExchange) {
    Step 'Exchange Online: sender mailbox and send-as scope'
    $sender = $config.Mail.From
    if ($dryRun) {
        Plan "shared mailbox $sender (if it doesn't exist)"
        Plan "Exchange service principal, management scope and 'Application Mail.Send' assignment limited to $sender"
    }
    else {
        # Graph is already connected; load Exchange second to avoid the known MSAL assembly clash.
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false
        if (-not $manifest.senderMailbox) {
            $mbx = Get-Recipient -Identity $sender -ErrorAction SilentlyContinue
            if ($mbx) { throw "$sender already exists. Point Mail.From at a new address, or remove this check if you meant to reuse it; nothing was changed." }
            New-Mailbox -Shared -Name $setup.SenderDisplayName -DisplayName $setup.SenderDisplayName -PrimarySmtpAddress $sender | Out-Null
            $manifest.senderMailbox = $sender; Save-Manifest
            Done "Created shared mailbox $sender"
        }
        if (-not $manifest.exoServicePrincipal) {
            New-ServicePrincipal -AppId $manifest.appId -ObjectId $manifest.spObjectId -DisplayName $setup.AppName | Out-Null
            $manifest.exoServicePrincipal = $true; Save-Manifest
        }
        $scopeName = "$($setup.AppName) sender"
        if (-not $manifest.exoScope) {
            New-ManagementScope -Name $scopeName -RecipientRestrictionFilter "PrimarySmtpAddress -eq '$sender'" | Out-Null
            $manifest.exoScope = $scopeName; Save-Manifest
        }
        if (-not $manifest.exoRoleAssignment) {
            $ra = New-ManagementRoleAssignment -App $manifest.spObjectId -Role 'Application Mail.Send' -CustomResourceScope $scopeName
            $manifest.exoRoleAssignment = [string]$ra.Name; Save-Manifest
        }
        Done "The app can send mail as $sender only"
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
#endregion

#region Summary
Step 'Done'
if ($dryRun) { Write-Host '   Dry run: nothing was created. Re-run without -WhatIf to create the items above.' -ForegroundColor Yellow; return }
Write-Host "   Manifest of everything created: $manifestPath"
Write-Host ''
Write-Host 'Put these into config.psd1:' -ForegroundColor Cyan
@"
    Graph.TenantId              = '$tenantId'
    Graph.ClientId              = '$($manifest.appId)'
    Graph.CertificateThumbprint = '$($manifest.certThumbprint)'
    Exchange.AppId              = '$($manifest.appId)'
    Exchange.CertificateThumbprint = '$($manifest.certThumbprint)'
    Requests.SiteId             = '$($manifest.siteId)'
    Requests.ListId             = '$($manifest.listId)'
    Requests.ListUrl            = '$($manifest.listUrl)'
    Requests.AuthorizedGroupIds = @('$($manifest.groupId)')
"@ | Write-Host
Write-Host ''
Write-Host 'Your steps (these grant lasting power, so they are yours to click):' -ForegroundColor Cyan
Write-Host "   1. Grant admin consent: Entra admin center > App registrations > $($setup.AppName) > API permissions > Grant admin consent"
Write-Host "      or open https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$($manifest.appId)"
Write-Host "   2. Entra admin center > Roles and admins > Exchange Recipient Administrator > Add assignments > $($setup.AppName)"
Write-Host '      (needed for contacts and shared-mailbox conversion)'
Write-Host '   3. Finish the list settings, form, Teams tab and approval flow: docs/FORM-AND-FLOW.md, sections 3-6'
#endregion
