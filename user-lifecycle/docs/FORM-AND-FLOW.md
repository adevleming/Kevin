# Employee lifecycle requests: setup and how it works

HR and hiring managers request a **new hire**, a **termination** (the Separation Checklist) or a
**change** from a tab in Teams. The request goes to Shelbae (payroll/HR) for approval, and IT's
automation does the rest. It uses a SharePoint list, standard Power Automate connectors and Teams,
so no premium licences are needed.

```
 Hiring manager (Teams "Requests" tab)      Shelbae / HR (Teams Approvals)         IT automation (every 15 min)
 ───────────────────────────────────       ──────────────────────────────         ─────────────────────────────
 + New: New hire / Termination / Change ─► Approve / Reject card ─► Ready for IT ─► • Employee Status email to Payroll
                                            (details to set up in Paycom)            (+ badge office for terms)
                                                                                     • account or painter contact
                                           Desk365 ticket (after approval)          • access off 6pm on last day
                                                                                     • title / manager changes
 Start date: "did everyone start?" email ◄──────────────────────────────────────── • no-shows: access removed,
 Mark a no-show on the request ───────────────────────────────────────────────────►   HR told to reverse in Paycom
```

## What replaced the paper forms and manual emails

| Before | Now |
|---|---|
| **Separation Checklist**, General Manager part (header, rehire, notice, exit interview, medical, AMEX, badge, repairman certificate) | Fields on the Termination request, filled in by the manager |
| **Separation Checklist**, HR part (final paycheck, PTO, benefits, HR and payroll exit checklists) | HR fields on the same request, filled in by HR after approval |
| **New Employee Checklist** (applicant source, job ad ID, the 21 paperwork items) | HR fields on the New hire request |
| Email to employeestatus@iac.aero ("AMA Term", "… New Hire", with the First/Last/Title/Effective Date/Location table) | Sent automatically, once per request, in the same format |
| Email the airport badge office about a termination | Sent automatically for sites with `BadgeOfficeEmails` set |
| Telling IT about a hire or a leaver | The request itself |
| Finding out someone didn't show up for orientation | A "did everyone start?" email on the start date. Marking a no-show removes the access |

The checklist items are in `config.psd1` (`Requests.HrNewHireChecklist`, `HrExitChecklist`,
`PayrollExitChecklist`, `ReturnItems`), so HR can change them. Re-run the list script, or edit
the column's choices in the list, after changing them. One item is still unclear: the
Separation Checklist's payroll line reads **"Remove from Pamir is"**. Check with Payroll what
system that is.

## 1. Run the setup script (creates new things only)

On your PC or the automation server, in PowerShell as a Microsoft 365 admin:

```powershell
Install-Module Microsoft.Graph.Authentication, ExchangeOnlineManagement -Scope CurrentUser
Copy-Item config.example.psd1 config.psd1        # fill in Setup.Members (HR, management)
./Setup-LifecycleTenant.ps1 -ConfigPath ./config.psd1 -WhatIf            # shows what it would create
./Setup-LifecycleTenant.ps1 -ConfigPath ./config.psd1 -IncludeExchange   # creates it
```

It creates:
- the private **Hiring & Staffing Requests** team, with its site and the request list;
- the **IT Lifecycle Automation** app (certificate included, no permissions consented);
- a write grant for that app on the new site only;
- with `-IncludeExchange`, the `it-automation@iac.aero` shared mailbox, with send-as scoped to that
  mailbox.

It stops without changing anything if a group or app with the same name already exists. It records
every object it creates in `state/tenant-setup.json` and prints the values to paste into
`config.psd1`.

Signing in asks you to consent to the Microsoft Graph PowerShell app's delegated permissions for
your own account. That's normal for any Graph PowerShell use.

Then do the two steps the script leaves to you:
1. **Grant admin consent** for the app: Entra admin center > App registrations > IT Lifecycle
   Automation > API permissions > *Grant admin consent*.
2. **Assign the app Exchange Recipient Administrator**: Entra admin center > Roles and admins. It
   needs this for painter contacts and for converting mailboxes to shared.

**Team roles.** Owners (IT, Shelbae, AnnaMarie) see every request and manage who's on the team.
Members are everyone with hiring ability. They submit requests and see only their own.
**Adding someone to the team gives them hiring ability**, and the script checks team membership
before acting on anything.

## 2. List settings

In the list: **Settings (gear) > List settings**.

- **Versioning settings:** keep version history on, with no version limit, or at least 100
  versions. **The automation relies on it.** It checks that the change to *Ready for IT* came from
  the flow or HR. It also checks that nobody else changed who, when or what access after approval.
  Version history can't be edited by ordinary users.
- **Advanced settings > Item-level permissions:**
  - Read access: **Read items that were created by the user**
  - Create and Edit access: **Create items and edit items that were created by the user**
- **Permissions for this list > Stop inheriting permissions**, then:
  - remove the site **Members** group;
  - add **Hiring & Staffing Requests Members** with **Contribute**;
  - keep **Owners** at **Full Control**;
  - add the flow's account (e.g. `flows@iac.aero`) with **Full Control**.

  Contribute can't change list settings. Only Full Control (or Design) sees every request,
  because item-level permissions only apply to people without "Override List Behaviors".
- **Indexed columns:** confirm **Status** and **Did they start?** (HireOutcome) are listed.
- **Validation settings** (optional):
  ```
  =IF([Request type]="New hire",AND(NOT(ISBLANK([Legal first name])),NOT(ISBLANK([Last name])),NOT(ISBLANK([Start date (orientation day)])),NOT(ISBLANK([Computer access]))),IF([Request type]="Termination",OR([Disable access immediately],NOT(ISBLANK([Last day worked]))),TRUE))
  ```
  User message: *New hires need a name, start date and computer access. Terminations need a last
  day worked, or tick "Disable access immediately".*

## 3. The form

Open the list, click **+ New**, then **Edit form (pencil) > Edit columns**.

**Hide** these; the automation and approval fill them in: Title, Status, Approved by, Account
created, IT automation log, Processed at, Employee status email sent, Start-day check sent.

**Order and show/hide.** For each column: **... > Edit conditional formula**. The formulas use the
internal names set by the setup script, e.g. `RequestType`. Hidden-by-formula fields keep their
values.

| Section | Columns | Formula |
|---|---|---|
| Everyone | Request type, Computer access, Location, Notes | *(always shown)* |
| Person | Legal first name, Preferred first name, Last name, Personal email, Mobile phone | `=if([$RequestType] == 'New hire' \|\| [$AccessType] == 'Contact only', 'true', 'false')` |
| Person | Employee (people picker) | `=if([$RequestType] != 'New hire' && [$AccessType] != 'Contact only', 'true', 'false')` |
| New hire / change | Department, Job title, Manager | `=if([$RequestType] == 'Termination', 'false', 'true')` |
| New hire | Start date (orientation day), Equipment needed, Buddy / ambassador | `=if([$RequestType] == 'New hire', 'true', 'false')` |
| New hire, after it's set up | Did they start? | `=if([$RequestType] == 'New hire' && [$Status] == 'Completed', 'true', 'false')` |
| New hire, HR | Applicant source, Job ad ID, Paycom employee code, New Employee Checklist (HR) | `=if([$RequestType] == 'New hire' && [$Status] != 'Submitted', 'true', 'false')` |
| Separation Checklist, manager | Termination date, Last day worked, Termination type, Rehire eligible, Reason for separation, Proper notice given, Exit interview, Outgoing medical testing, Outstanding equipment purchases, Outstanding advances / AMEX card collected, Disable access immediately, Give mailbox and files to | `=if([$RequestType] == 'Termination', 'true', 'false')` |
| Separation Checklist, last day | ID badge collected, Repairman certificate collected, Equipment / PPE returned | `=if([$RequestType] == 'Termination' && [$Status] != 'Submitted', 'true', 'false')` |
| Separation Checklist, HR | Final paycheck date, PTO due, Exit interview completed, Benefits ending, Benefits date of termination, HR exit checklist, Payroll exit checklist, Paycom employee code | `=if([$RequestType] == 'Termination' && [$Status] != 'Submitted', 'true', 'false')` |
| Change | Change effective date | `=if([$RequestType] == 'Change', 'true', 'false')` |

HR sections stay hidden while a request is being filled in (Status *Submitted*) and appear once
it's in progress. SharePoint forms can't hide columns by person, so a manager opening their own
request later can see those sections too.

What each request type means for the person filling it in:

- **New hire, Full or Basic user:** name, location, department, title, manager, start date
  (orientation day), equipment. The account is ready before day one.
- **New hire, Contact only** (painters and others without a computer): name, personal email and
  phone. They're added to the address book and the site's distribution list. No account, no
  licence.
- **Termination:** pick the **Employee**, then fill in the manager's part of the Separation
  Checklist. Access ends at **6pm site time on the last day worked**, or immediately if *Disable
  access immediately* is ticked or it's involuntary. For a contact-only person, type their name and
  personal email instead of picking an employee. After the last day, tick badge, repairman
  certificate and equipment returned.
- **Change:** pick the Employee, then the new title, department, location or manager, and the date
  it takes effect.

## 4. Put it in Teams

1. In the **Hiring & Staffing Requests** team's General channel: **+ (Add a tab) > Lists > Add an
   existing list** > *Employee Lifecycle Requests*. Name the tab **Requests**.
2. Post a pinned message: *To hire, let someone go or change someone's role, open the Requests tab
   and click + New.*
3. Optional: pin the team for its members with a Teams setup policy.

Shelbae needs nothing extra. Approvals appear in her Teams activity feed and in the **Approvals** app.

## 5. The approval flow

Create it in Power Automate (make.powerautomate.com), ideally under a dedicated `flows@iac.aero`
account. That account's email goes in `Requests.TrustedEditors`, along with Shelbae's and
AnnaMarie's.

**Flow: "Lifecycle request - submitted"**

1. **Trigger:** SharePoint, *When an item is created*, on the Employee Lifecycle Requests list.
2. **Compose a summary** for the approval and the ticket. Include everything the automation acts
   on, so HR isn't approving blind:
   - request type and computer access;
   - for terminations and changes, *Employee DisplayName* **and** *Employee Email*; for new hires,
     the typed name;
   - location, department, title, manager;
   - start date, or last day worked;
   - *Disable access immediately*, and **Give mailbox and files to** (*MailboxDelegate Email*);
   - equipment, notes, and *Link to item*.

   Put requester-typed text in as plain text, not HTML.
3. **Update item:** Title = `<Request type> - <name>`, Status = **Pending approval**. Pass Request
   type from the trigger; Update item needs every required column.
4. **Condition:** Request type = Termination **and** (Disable access immediately = true **or**
   Termination type = Involuntary).
   - **Yes, immediate:**
     - *Update item*: Status = **Ready for IT**, Approved by = `Immediate - HR notified`.
     - *Send an email (V2)*: FYI to Shelbae and AnnaMarie.
     - *Send an email (V2)* to the Desk365 address.
   - **No:**
     - *Initialize variable* `asked` = `utcNow()`.
     - *Start and wait for an approval*: Approve/Reject - First to respond, assigned to Shelbae; AnnaMarie.
       Title `Approve: <Request type> - <name>`, Details = summary, Item link = *Link to item*.
     - *Get item* (the same ID).
     - **Condition:** Outcome = Approve **and** *Get item Modified* is less than `asked`. That
       second part means nothing changed while the approval was waiting.
       - **Yes:**
         - *Update item*: Status = **Ready for IT**, Approved by = `<Responses Approver name> <Response date>`.
         - *Send an email (V2)* to the Desk365 address: `[Onboarding|Offboarding|Change] <name>` with the summary.
       - **No, rejected:** *Update item* Status = **Rejected**. Then Teams *Post message in a chat or
         channel* (Flow bot) to *Created By Email*, with the approver's comments.
       - **No, approved but edited meanwhile:** *Update item* Status = **Needs IT review**, and let
         HR know. Then set it back to *Pending approval* by hand to re-run, or build a second
         approval round.

The automation sends the **Employee Status email** to employeestatus@iac.aero and the badge office
itself, so the flow doesn't need to.

Shelbae sets new hires up in Paycom from the approval card. She adds the **Paycom employee code**
to the request when she has it; the automation stamps it on the account. She fills in the HR part
of the checklist on the request at any time.

**Optional:** a flow that posts a Teams message to the requester when Status changes to Completed,
Needs IT review or Reversed. Use a trigger condition on `body/Status/Value`. The automation
already emails the requester, the manager and IT.

## 6. The IT automation

```powershell
./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1           # dry run: prints what it would do
./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1 -Apply    # does it
```

Schedule the `-Apply` run **every 15 minutes** with Task Scheduler on a utility server. Azure
Automation can't schedule more often than hourly. Each run:

- picks up requests that are *Ready for IT*, *Scheduled*, or no-shows;
- holds a request (Status becomes **Needs IT review**) if:
  - it isn't approved;
  - the approval didn't come from the flow or HR;
  - someone else changed a key field after approval;
  - the requester isn't in the team;
- sends the **Employee Status email**, once per request;
- marks the request **In progress**, does the work, then writes Status, the account name and a log
  line back. An interrupted run never repeats a request; it goes to review instead;
- leaves future terminations **Scheduled** until 6pm site time on the last day worked;
- never offboards more than 5 people in one run or 15 in 24 hours, anyone on the protected list, or
  accounts outside `iac.aero`, such as Eirtech users or guests;
- for an **immediate** termination, which skips approval, only acts if the requester is the
  employee's manager in Entra, or HR/IT. The mailbox then goes only to that manager; any other
  named person gets it after HR reviews. If sign-in can't be blocked, the mailbox is left alone;
- only removes address-book contacts that it created itself;
- carries on with the other requests if one fails.

**Statuses:** Submitted > Pending approval > Ready for IT > (Scheduled) > In progress > Completed.
The other outcomes are:
- **Needs IT review:** a person has to look. The IT automation log on the request says why.
- **Rejected** and **Cancelled**.
- **Reversed - did not start**.

To re-run a request after fixing the problem, set it back to *Ready for IT*. That has to be done by
one of the trusted editors (HR, IT, or the flow). A failed request also emails IT with the error.

## 7. Start day and no-shows

On the start date, after 10:00 site time (`Requests.NoShow.CheckTime`), the automation looks at
every new hire due that day at each location:

- **Anyone whose new account has already signed in** is marked *Started* automatically.
- **Everyone else** gets one email per location, to the hiring manager(s) and the site's
  `OrientationContacts` (the site admin): *"10 people were due to start at Amarillo (AMA). For
  anyone who didn't start, open their request and set Did they start? to No-show."* Each person has
  a link straight to their request.

Setting **Did they start? = No-show / not starting** makes the next run undo the hire:
- **Full or Basic user:** the account is turned off, signed out, and removed from all its groups,
  licence groups included; there's no mailbox worth keeping. Set `Requests.NoShow.DeleteAccount =
  $true` to also delete it (restorable for 30 days).
- **Contact only:** the contact is removed.
- **Not set up yet:** the request is simply *Cancelled*.
- **Safeguard:** if the account has signed in successfully, it isn't touched. It goes to review,
  since they may have started after all.
- **Notifications:** Payroll (employeestatus@) and HR get an "AMA No-show" email so the hire can be
  reversed in Paycom.

A no-show can be marked any time, before or after the start date. The request ends as
**Reversed - did not start**.

## 8. Test it end to end

1. Submit a **New hire, Full user** for a made-up person, starting today. Approve it in Teams.
   Within 15 minutes:
   - the request shows *Completed* with the account name;
   - Payroll gets "GEG New Hire";
   - after 10:00 the start-day email arrives.
2. Mark it **No-show**. The next run turns the account off, and the request shows *Reversed - did
   not start*.
3. Submit a **Termination** for another test account with last day = today. It goes *Scheduled*
   (Payroll gets "GEG Term"), then *Completed* after 6pm site time. Check that sign-in is blocked
   and the mailbox is shared.
4. Submit a **New hire, Contact only** with a personal email. It appears in the address book.
5. As the requester, change the Employee on an approved termination. The next run sends it to
   *Needs IT review*, naming who changed what.
6. Have someone outside the team submit a request. It goes to *Needs IT review*.
