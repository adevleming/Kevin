# User lifecycle automation (new hires, terminations, role changes)

Automates IT's side of hiring, letting people go, and role changes. It doesn't use Paycom's API,
which costs about $1,000 per payroll cycle.

## How it works

Two parts:

1. **The request form (primary).** Hiring managers and HR submit a **New hire**, **Termination** or
   **Change** request from a tab in Teams. Shelbae (payroll/HR) gets an approval card in Teams with
   everything she needs to set the person up in Paycom. Once approved, IT's automation:
   - creates the account before day one (licence, site and department groups, manager), **or** a
     contact for people who don't need a computer (painters);
   - turns off access at **6pm site time on the last day**, or immediately for involuntary
     terminations; converts the mailbox to shared and gives it to whoever the requester named;
   - updates title, department, site and manager for role changes;
   - writes the result back to the request and emails the requester, the manager and IT;
   - sends the **Employee Status** email to Payroll (employeestatus@iac.aero, HR's usual "AMA
     Term" format) and tells the site badge office about terminations;
   - on the **start date**, asks the hiring manager and site admin "did everyone start?". Marking
     a **no-show** removes the access that was set up and tells HR to reverse the hire in Paycom.

   The form carries HR's **Separation Checklist** and **New Employee Checklist**, so it replaces
   those forms rather than adding another one.
2. **The weekly Paycom audit (backstop).** Paycom only marks someone terminated after their final
   payroll, and its weekly email is just a link to the Report Center. So Paycom can't drive the process,
   but it's still the payroll record. Once a week someone downloads the *IT Current Employees* report,
   and the audit flags:
   - anyone hired or terminated in Paycom that nobody filed a form for;
   - enabled accounts with no employee behind them (unused licences, security risk);
   - terminated people whose account is still on.

```
Teams "Requests" tab ──► SharePoint list ──► Power Automate: HR approval in Teams, Desk365 ticket
                                                     │ Status = Ready for IT
                                                     ▼
                           Invoke-LifecycleRequests.ps1 (every 15 min) ──► Entra ID / Exchange Online
                                                     ▲
Paycom weekly CSV ──► Invoke-PaycomLifecycle.ps1 ────┘  (audit: gaps, orphaned accounts, report to IT)
```

| File | What it is |
|---|---|
| [`docs/FORM-AND-FLOW.md`](docs/FORM-AND-FLOW.md) | Setup and how it works: list settings, form, Teams tab, approval flow, start day and no-shows |
| `Setup-LifecycleTenant.ps1` | One-time, run by you: creates the team, site, list, app and mailbox (new objects only; `-WhatIf` first) |
| `New-LifecycleRequestList.ps1` | Creates just the list on an existing site, if you'd rather not use the setup script |
| `Invoke-LifecycleRequests.ps1` | Scheduled every 15 min: processes approved requests |
| `Invoke-PaycomLifecycle.ps1` | Weekly: the Paycom audit report |
| `LifecycleRequests.psm1`, `PaycomLifecycle.psm1` | The logic behind both |
| Lifecycle Site Settings (SharePoint list) | Badge office and start-day email per site, editable by HR without IT ([section 8](docs/FORM-AND-FLOW.md#8-changing-badge-offices-start-day-contacts-and-whos-on-the-team)) |
| `config.example.psd1` | Copy to `config.psd1` (git-ignored) and fill in |

## Safeguards

- **Only people with hiring ability can submit.** List permissions and the Teams team membership
  control it, and the script checks again that the requester is in an authorised group.
- **Nothing runs without approval,** except immediate terminations, where HR is notified in
  parallel instead of blocking.
- **Approved requests can't be quietly changed.** The script reads the list's version history,
  which ordinary users can't edit. The change to *Ready for IT* must have come from the flow or HR,
  and nobody else may have changed who, when or what access afterwards. Requesters can still update
  checklist items, such as badge collected or a no-show.
- **Terminations can't run away.** At most 5 per run and 15 per day. Protected accounts, and
  accounts outside `iac.aero`, are never touched automatically. Accounts are disabled, never
  deleted (unless you turn that on for no-shows).
- **Immediate terminations are limited.** They skip approval, so they only run if the requester is
  the employee's manager or HR/IT. The mailbox then goes only to the manager.
- **No passwords in email.** New accounts get a random password nobody sees. On day one the tech
  issues a Temporary Access Pass, and the user sets up MFA and a password.
- **The Paycom audit won't act on a bad export.** If the roster looks truncated or headcount drops
  sharply, nothing changes and no tickets go out.

## One-time setup

`Setup-LifecycleTenant.ps1` does most of this: see [docs/FORM-AND-FLOW.md](docs/FORM-AND-FLOW.md#1-run-the-setup-script-creates-new-things-only).
The details below are for reference or for doing it by hand.

1. **Entra app registration** (e.g. `IT Lifecycle Automation`) with a certificate. Grant these
   *application* permissions (admin consent):

   | Permission | Needed for |
   |---|---|
   | `User.ReadWrite.All` | Create accounts, block sign-in, set title/department/manager, stamp employee ID |
   | `User.RevokeSessions.All` | Sign leavers out everywhere (`User.ReadWrite.All` doesn't cover this) |
   | `GroupMember.ReadWrite.All` | Add new accounts to licence / site / department groups (security or Microsoft 365 groups) |
   | `GroupMember.Read.All` | Check the requester is in an authorised group |
   | `Sites.Selected` | The request list and Paycom drop folder. Grant the app **write** on that one site. |
   | `AuditLog.Read.All` | Last sign-in dates in the Paycom audit (optional, Entra ID P1) |
   | Office 365 Exchange Online > `Exchange.ManageAsApp` | Contacts and mailbox conversion. Also assign the app's service principal the **Exchange Recipient Administrator** role. |

   Don't consent Microsoft Graph `Mail.Send` tenant-wide. Scope it to the one sender mailbox with
   Exchange **RBAC for Applications** (Application Access Policies are being retired):
   ```powershell
   New-ServicePrincipal -AppId <app-id> -ObjectId <enterprise-app-object-id> -DisplayName 'IT Lifecycle Automation'
   New-ManagementScope -Name 'IT automation sender' -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'it-automation@iac.aero'"
   New-ManagementRoleAssignment -App <enterprise-app-object-id> -Role 'Application Mail.Send' -CustomResourceScope 'IT automation sender'
   ```
   Accounts that hold admin roles can't be disabled by this app. Those requests go to *Needs IT
   review*, which is intended.
2. **Sender mailbox** `it-automation@iac.aero` (`Mail.From`).
3. **The request list, form, Teams tab and flow:** follow [`docs/FORM-AND-FLOW.md`](docs/FORM-AND-FLOW.md).
4. **Config:** copy `config.example.psd1` to `config.psd1`. Fill in the tenant and app IDs, `Requests`
   (site/list IDs, authorised groups, trusted editors), `Sites` (time zones and group IDs), and
   `AccessTypes` (licence group per access level). Licence and site groups must be security or
   Microsoft 365 groups. Distribution lists only work in `Sites.ContactGroups`, which go through
   Exchange.
5. **Host:** install `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` on the server that
   runs the scripts (plus `ActiveDirectory` if `DirectoryMode = 'Hybrid'`).
   - **Pin the module versions.** Recent versions of these two modules are known to clash when both
     load in one session. The scripts always connect Graph first, which avoids it. Once a pair
     works, pin it (`Install-Module -RequiredVersion`) and test before upgrading either.
   - **Azure Automation:** use a PowerShell **7.4** runtime environment; 7.2 isn't supported after
     Sept 30, 2026. Task Scheduler is still the better fit for the 15-minute schedule.

> The Paycom API sandbox SID and token were emailed in plain text on 9/2. Since IAC isn't using the API,
> ask Paycom to revoke them.

## Run

```powershell
# Requests: dry run first (prints what it would do), then schedule the -Apply run every 15 minutes
./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1
./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1 -Apply

# Paycom audit: weekly, after the CSV is saved to the drop folder
./Invoke-PaycomLifecycle.ps1 -ConfigPath ./config.psd1
```

Use Task Scheduler on a utility server for the 15-minute schedule; Azure Automation can't schedule
more often than hourly.

## The weekly Paycom audit

Paycom's *IT Current Employees* push report (Brian Stamer set it up on 8/12) only emails a link.
Someone with Paycom access clicks **VIEW REPORTS**, downloads the CSV, and saves it to an IT-only
folder (`Input.Path` in config, e.g. a share on the automation server). **Not** the Hiring &
Staffing team site: that file lists every employee, hiring managers can read the team's files,
and a replaced CSV could fake terminations. If nobody does, IT gets a reminder instead of the week
being skipped silently.

What to ask Brian to change on the report:

1. **Delivery:** have the notice go to `it-automation@iac.aero` (or a shared IT mailbox) instead of one
   person.
2. **Columns:** employee code, legal first and last name, preferred name, work email, department,
   position, primary supervisor name and email, location, employee status, hire date, termination
   date. Then set `Roster.Columns` to the exact CSV headers.
3. **Include recent terminations:** *Status = Active* **or** *termination date in the last 30 days*,
   so leavers show with a real date instead of just disappearing.
4. **Format:** CSV.

The audit compares each week's file with the last one and with Entra ID. Where a form already covers
someone, the audit stays quiet. It only raises Desk365 tickets for gaps. It can also stamp the Paycom
employee code onto matched accounts (`BackfillEmployeeId`), so later matches are exact. See
[`docs/sample-report.png`](docs/sample-report.png) for the report built from the test data.

## Scope and limits

- Paycom company `0QT40` is North America only. `Scope.Domains = @('iac.aero')` keeps Eirtech
  (`etas.ie`) accounts out of the audit.
- **Hybrid AD:** with `DirectoryMode = 'Hybrid'`, leavers are disabled in on-prem AD, and new-hire
  requests go to *Needs IT review* for the account to be made in AD.
- Removing licences (after the mailbox is converted), wiping devices, collecting equipment and
  removing third-party app access stay on the Desk365 ticket.

## Next: painter crews and hangar scheduling

The proposal is in [`docs/CREW-SCHEDULING.md`](docs/CREW-SCHEDULING.md):
- a crew roster fed by the lifecycle requests;
- aircraft and assignment lists that GMs edit in Teams;
- a page for the hangar UniFi displays, on the internal network only;
- a weekly email to each painter's personal address, for people without company mailboxes.

It's not built yet; the open questions at the end of that document need answers first.

## Tests

```powershell
pwsh ./tests/Test-LifecycleRequests.ps1     # form-driven requests (mocked Graph and Exchange)
pwsh ./tests/Test-SetupLifecycleTenant.ps1  # setup script against a fake Graph: creates only, never modifies
pwsh ./tests/Test-PaycomLifecycle.ps1       # Paycom audit
```

Offline and self-contained (no Pester, no tenant). The fixtures cover:
- every request type and access level;
- time zones and the last-day cutoff;
- no-shows and the start-day check;
- the Employee Status email format;
- rehires, same-name hires and accented names;
- forged or edited approvals (version history);
- requesters outside the team;
- the per-run termination cap;
- truncated Paycom exports.
