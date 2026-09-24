# Paycom → Entra ID user lifecycle

Automates joiners, leavers and movers using the Paycom report IAC already gets every week. It does
**not** need Paycom's paid API/SFTP integration.

## Why it works this way

The weekly email from `systemmessage@paycomonline.com` ("Scheduled Reports Completed") doesn't list any names.
It only says the **IT Current Employees** Push Report is ready in Paycom's Report Center. That report is a
full list of current employees, so new hires and leavers have to be worked out by comparing it with last
week's copy:

| Signal | How it's detected |
|---|---|
| New hire | Active on this week's report, not active last week (rehires included) |
| Termination | Active last week, now terminated in Paycom **or** gone from the report |
| Mover | Department, title, manager, name or location changed |
| Termed but still enabled | Inactive in Paycom, Entra account still enabled |
| Orphaned account | Enabled `@iac.aero` account with no active employee behind it (service, room and guest accounts are excluded) |

The Entra reconciliation runs on every roster, so the first run already shows stale accounts that are still
using licences.

## What a run does

```
Paycom push report ──► CSV in SharePoint "Paycom Roster" folder
                              │
                    Invoke-PaycomLifecycle.ps1  (weekly, Thursday morning)
                              │
         ┌────────────────────┼─────────────────────────┐
   diff vs last snapshot   reconcile vs Entra ID    safety checks
         └────────────────────┼─────────────────────────┘
                              ▼
   • Summary report email to IT              (always)
   • One Desk365 ticket per joiner/leaver    (email-to-ticket, with checklist)
   • With -Apply: block sign-in + revoke sessions for leavers,
     create accounts for eligible joiners, stamp Paycom employee code on accounts
```

See [`docs/sample-report.png`](docs/sample-report.png) for the report built from the test data.

**Safety checks.** A bad export (wrong report, partial download, filter mistake) would look like a mass
termination. So if the roster has too few people, headcount drops too far, or there are too many leavers in
one run, nothing is changed and no tickets are sent. The report is flagged, and the previous snapshot stays
the baseline.

**Leavers are contained, not deleted.** Sign-in is blocked and sessions are revoked. Licences, licence groups
and the mailbox are left alone so a tech can convert the mailbox to shared for the manager. The ticket
covers the rest.

**No passwords in email.** New accounts get a random password that nobody sees. On day one the tech issues a
Temporary Access Pass, and the user registers MFA and sets their own password.

## 1. Get the export out of Paycom

Pick one:

- **A. Fully automatic (ask your Paycom specialist first).** Ask whether the *IT Current Employees* Push
  Report can deliver the file itself (email attachment or a Paycom-hosted download) instead of just a
  notification. If it can, add a Power Automate flow: *When a new email arrives* (from
  `systemmessage@paycomonline.com`, has attachments) → *Create file* in the SharePoint `Paycom Roster`
  folder.
- **B. One-minute manual step.** When the Thursday email arrives, whoever gets it clicks **VIEW REPORTS**,
  downloads the report as CSV, and saves it to the SharePoint `Paycom Roster` folder. If nobody does, the run
  emails IT a reminder instead of silently skipping the week.

Changes worth asking HR/Paycom to make to the push report (they're free):

1. **Columns.** At minimum: employee code, legal first/last name, preferred name, work email, department,
   position, primary supervisor (name and email), location, employee status, hire date, termination date.
   Then set `Roster.Columns` in config to match the exact CSV headers.
2. **Include recent terminations.** Filter on *Status = Active* **or** *Termination date in the last 30
   days*. Leavers then come through with a real termination date instead of just disappearing. Both cases
   are handled either way.
3. **Run it daily** if Push Reporting allows it. The script ignores a file it has already processed, so a
   daily run costs nothing and cuts leaver detection from up to 7 days to 1.

## 2. One-time Entra setup

1. **App registration** (e.g. `IT Lifecycle Automation`) with a certificate. Grant these *application*
   permissions:

   | Permission | Needed for |
   |---|---|
   | `User.Read.All` | Report-only mode |
   | `User.ReadWrite.All` | Block sign-in, revoke sessions, create users, stamp employee ID |
   | `GroupMember.ReadWrite.All` | Onboarding groups / offboarding group cleanup |
   | `AuditLog.Read.All` | Last sign-in date in the report (optional, Entra ID P1) |
   | `Mail.Send` | Report and ticket emails. **Limit it to the sender mailbox** with `New-ApplicationAccessPolicy` |
   | `Sites.Selected` | Read the SharePoint drop folder (grant the app access to just that site) |

   Accounts with admin roles can't be disabled by this app. They show up as failed in the ticket, which is
   intended.
2. **Sender mailbox**, e.g. `it-automation@iac.aero`, for `Mail.From`.
3. **SharePoint folder.** Create `Paycom Roster` in the IT site's document library. Get the drive ID from
   Graph Explorer: `GET /sites/iacaero.sharepoint.com:/sites/IT:/drives` (use your real site path).
4. **Desk365.** Set `Tickets.SendTo` to the support address Desk365 turns into tickets.
5. Install `Microsoft.Graph.Authentication` on the host that runs the script. It's the only module needed,
   plus `ActiveDirectory` in Hybrid mode.

> The Paycom API sandbox SID/token were emailed in plain text on 9/2. Since IAC isn't going ahead with the
> API, ask Paycom to revoke them.

## 3. Configure and run

```powershell
Copy-Item config.example.psd1 config.psd1   # fill in IDs, domains, column names, groups
# Dry run against a downloaded export: nothing is sent or changed, output goes to state/reports
./Invoke-PaycomLifecycle.ps1 -ConfigPath ./config.psd1 -RosterPath '.\IT Current Employees.csv' -NoEmail
```

Schedule it for Thursday mornings after the 1:30 AM report, either as a Task Scheduler job on a utility
server (certificate auth) or as an Azure Automation runbook (managed identity):

```powershell
pwsh -File C:\Scripts\paycom-lifecycle\Invoke-PaycomLifecycle.ps1 -ConfigPath C:\Scripts\paycom-lifecycle\config.psd1 -Apply
```

`-Apply` only allows changes. Each type of change is also switched off in config until you turn it on.

## 4. Rollout

| Phase | Settings | Goal |
|---|---|---|
| Weeks 1-2 | no `-Apply` | Check the column map, tune `Scope.ExcludePatterns`/`ExcludeUpns` until the orphan list is only real leftovers, clean those up |
| Weeks 3-4 | `-Apply`, `BackfillEmployeeId = $true` | Tickets flow to Desk365. Accounts get the Paycom employee code so matching becomes exact |
| Week 5+ | `Offboarding.Enabled = $true` | Leavers are blocked automatically when the run sees them |
| Later | `Onboarding.Enabled = $true`, department → group map | Accounts ready before day one. Set `Eligibility` so shop-floor roles without M365 don't get accounts |

Keep the existing rule that **HR tells IT the same day about involuntary terminations**. A scheduled report
is a safety net, not a replacement for that.

## Scope and limits

- Paycom company `0QT40` covers North America only. `Scope.Domains = @('iac.aero')` keeps Eirtech
  (`etas.ie`) accounts out of the reconciliation. The same tool could reconcile Eirtech once it has its own
  source of truth.
- **Hybrid AD:** set `DirectoryMode = 'Hybrid'`. Leavers are disabled in on-prem AD (and moved to
  `DisabledUsersOU`); joiners get a pre-filled ticket instead of an automatic account.
- Converting mailboxes, removing licences and wiping devices stay on the ticket checklist. They need
  Exchange/Intune permissions and a person's judgement about timing.

## Tests

```powershell
pwsh ./tests/Test-PaycomLifecycle.ps1
```

Offline and self-contained (no Pester, no Graph). The fixtures cover rehires, same-name new hires,
accented names, leavers who drop off the report, service/room/guest/Eirtech exclusions, a truncated export,
and mocked Graph calls for onboarding and offboarding.
