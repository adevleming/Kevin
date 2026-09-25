# Employee lifecycle requests: form, Teams and approval flow

How HR and hiring managers request a new hire, a termination or a change, and how it reaches
Shelbae (payroll/HR) for approval and then IT's automation. Nothing here needs premium Power
Platform licences. It uses a SharePoint list, the standard Power Automate connectors, and Teams.

```
 Hiring manager (Teams tab)                 Shelbae / HR (Teams Approvals)          IT automation (every 15 min)
 ───────────────────────────                ───────────────────────────────         ────────────────────────────
 + New request  ──► SharePoint list item ──► Approve / Reject card ──► Status = ──► Invoke-LifecycleRequests.ps1
   New hire          Status: Pending          (all details, so HR can   Ready for IT   • create account or contact
   Termination       approval                  set them up in Paycom)                  • offboard at 6pm on last day
   Change                                                                               • update title / manager
                     Desk365 ticket ◄──────── (sent once approved)                     • writes Status/log back
                                                                                         • emails requester + manager
 Immediate terminations skip the approval wait: they go straight to Ready for IT and HR gets an FYI.
```

## 1. Groups

| Group | Members | Used for |
|---|---|---|
| **Hiring Managers** (the Team's Microsoft 365 group, see step 5) | Everyone allowed to hire, terminate or change a role | Who can submit. Put its object ID in `Requests.AuthorizedGroupIds`. |
| **HR Approvers** (security or M365 group) | Shelbae plus a backup | Approves requests; owns the list |
| **IT** | IT team | Owns the list, sees everything |

The script checks the requester is in an authorised group before it acts. Even if list permissions
are set up wrong, someone outside those groups can't trigger account changes.

## 2. Create the list

1. Pick the site: the SharePoint site behind the Team you create in step 5 works well.
2. Get its site ID with Graph Explorer: `GET https://graph.microsoft.com/v1.0/sites/iacaero.sharepoint.com:/sites/<site-name>`
   (use your real site path). Put the `id` into `Requests.SiteId` and the list URL into `Requests.ListUrl`.
3. Run the provisioning script once. It needs a temporary `manage` grant on the site (see the README):
   ```powershell
   ./New-LifecycleRequestList.ps1 -ConfigPath ./config.psd1 -WhatIf   # review the columns
   ./New-LifecycleRequestList.ps1 -ConfigPath ./config.psd1
   ```
   Copy the printed list ID into `Requests.ListId`. Then drop the app's site grant back to `write`.
4. Check the Site, Department and Computer access choices match `config.psd1`. They were generated
   from it, so if you edit them later, change both.

## 3. List settings

In the list: **Settings (gear) > List settings**.

- **Versioning settings:** keep version history on, so every change to a request is recorded.
- **Advanced settings > Item-level permissions:**
  - Read access: **Read items that were created by the user**
  - Create and Edit access: **Create items and edit items that were created by the user**

  Managers see only their own requests. HR and IT, as list owners, see all of them.
- **Permissions for this list:** stop inheriting, then grant:
  - Hiring Managers: **Contribute**
  - HR Approvers: **Edit** (or Full Control)
  - IT: **Full Control**
- **Indexed columns:** confirm **Status** is listed. The script queries on it.
- **Validation settings.** SharePoint list validation can't check person columns, but it can
  enforce the basics:
  ```
  =IF([Request type]="New hire",AND(NOT(ISBLANK([Legal first name])),NOT(ISBLANK([Last name])),NOT(ISBLANK([Start date])),NOT(ISBLANK([Computer access]))),IF([Request type]="Termination",OR([Disable access immediately],NOT(ISBLANK([Last day worked]))),TRUE))
  ```
  User message: *New hires need a name, start date and computer access. Terminations need a last
  day, or tick "Disable access immediately".*

## 4. The form

Open the list, click **+ New**, then **Edit form (pencil) > Edit columns**.

**Hide** these (untick them). HR and the automation fill them in, not the requester:
Title, Status, Paycom employee code, Approved by, Account created, IT automation log, Processed at.

**Order** the rest: Request type, Computer access, Legal first name, Preferred first name, Last
name, Employee, Site, Department, Job title, Manager, Start date, Equipment needed, Personal email,
Mobile phone, Last day worked, Termination type, Disable access immediately, Give mailbox and files
to, Change effective date, Notes for HR / IT.

**Show only what's relevant.** For each column, choose **... > Edit conditional formula**:

| Column(s) | Formula |
|---|---|
| Legal first name, Preferred first name, Last name, Personal email, Mobile phone | `=if([$RequestType] == 'New hire' \|\| [$AccessType] == 'Contact only', 'true', 'false')` |
| Start date, Equipment needed | `=if([$RequestType] == 'New hire', 'true', 'false')` |
| Employee | `=if([$RequestType] != 'New hire' && [$AccessType] != 'Contact only', 'true', 'false')` |
| Site, Department, Job title, Manager | `=if([$RequestType] == 'Termination', 'false', 'true')` |
| Last day worked, Termination type, Disable access immediately, Give mailbox and files to | `=if([$RequestType] == 'Termination', 'true', 'false')` |
| Change effective date | `=if([$RequestType] == 'Change', 'true', 'false')` |

(In the formulas, `[$Name]` uses the column's internal name, which is the name without spaces set
by the provisioning script, such as `RequestType` and `AccessType`.)

What each request type means for the requester:

- **New hire, Full or Basic user:** name, site, department, title, manager, start date, equipment.
  IT creates the account before day one.
- **New hire, Contact only (painters and others without a computer):** name plus personal email and
  phone. They're added to the address book and the site distribution list. No account, no licence.
- **Termination:** pick the **Employee** (the people picker selects the exact account), the last
  day, and who gets their mailbox and files. Access ends at 6pm site time on the last day, or
  straight away if *Disable access immediately* is ticked or the termination is involuntary. For a
  contact-only person, type their name and personal email instead of picking an employee.
- **Change:** pick the Employee, then enter the new title, department, site or manager and the date
  it takes effect.

## 5. Put it in Teams

1. In Teams, create a team, e.g. **Hiring & Staffing Requests** (private). Add everyone with hiring
   ability as members, and HR and IT as owners. The team's group is your **Hiring Managers** group.
   Adding someone to the team gives them the ability to submit.
2. In the General channel: **+ (Add a tab) > Lists > Add an existing list** and pick *Employee
   Lifecycle Requests*. Name the tab **Requests**.
3. Post a pinned message: *To hire, let someone go, or change someone's role, open the Requests tab
   and click + New.*
4. Optional: in the Teams admin center, pin the team for the Hiring Managers group through a setup
   policy so it's always in their sidebar.

Shelbae needs nothing extra. Approvals appear in her Teams activity feed and in the **Approvals** app.

## 6. The approval flow

Create it in Power Automate (make.powerautomate.com) under an account that will stay around, ideally
a dedicated `flows@iac.aero` service account. Add that account to `Requests.TrustedEditors`, along
with Shelbae and the HR backup.

**Flow: "Lifecycle request - submitted"**

1. **Trigger:** SharePoint, *When an item is created*. Site = the Team's site, List = Employee
   Lifecycle Requests.
2. **Lock the request.**
   - SharePoint, *Stop sharing an item or a file* on the new item's ID.
   - Then *Grant access to an item or a folder*: recipient = *Created By Email*, role = **Can view**.

   The requester can still see it but can't change it (for example, switch which employee a
   termination points at). HR and IT keep their access as list owners.
3. **Compose a summary** for the approval card and ticket: request type, name (or *Employee
   DisplayName*), site, department, title, manager, start date or last day, computer access,
   equipment, notes, and a link to the item (*Link to item*).
4. **Update item:** Title = `<Request type> - <name>`, Status = **Pending approval**.
   Update item needs every required column, so pass Request type from the trigger.
5. **Condition:** Request type = Termination **and** (Disable access immediately = true **or**
   Termination type = Involuntary).
   - **If yes (immediate):**
     - *Update item*: Status = **Ready for IT**, Approved by = `Immediate termination - HR notified`.
     - *Send an email (V2)* to HR Approvers: FYI with the summary.
     - *Send an email (V2)* to the Desk365 address: subject `[Offboarding] <name> - IMMEDIATE`, body = summary.
   - **If no:**
     - *Start and wait for an approval*: type **Approve/Reject - First to respond**, assigned to
       the HR Approvers (Shelbae; backup). Title `Approve: <Request type> - <name>`, Details = summary,
       Item link = *Link to item*.
     - **Condition:** Outcome = Approve.
       - **Approve:**
         - *Update item*: Status = **Ready for IT**, Approved by = `<Responses Approver name> <Response date>`.
         - *Send an email (V2)* to the Desk365 address: subject `[Onboarding|Offboarding|Change] <name>`, body = summary.
       - **Reject:**
         - *Update item*: Status = **Rejected**.
         - Teams *Post message in a chat or channel* as Flow bot to *Created By Email*: rejected, with the
           approver's comments.

Shelbae sets new hires up in Paycom from the approval card. When she has the Paycom employee code,
she can type it into the item's **Paycom employee code** column (grid view). The automation stamps it
on the account, and the weekly Paycom audit backfills any she didn't enter.

**Optional flow: "Lifecycle request - completed"**

- **Trigger:** *When an item is created or modified*, with a trigger condition
  `@or(equals(triggerOutputs()?['body/Status/Value'],'Completed'),equals(triggerOutputs()?['body/Status/Value'],'Needs IT review'))`.
- **Action:** Teams *Post message in a chat or channel* to the requester, with the IT automation log.
  The IT script already emails the requester, manager and IT, so this is only for people who'd
  rather get a Teams message.

## 7. Run the IT automation

```powershell
./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1           # dry run: prints what it would do
./Invoke-LifecycleRequests.ps1 -ConfigPath ./config.psd1 -Apply    # does it
```

Schedule the `-Apply` run **every 15 minutes** with Task Scheduler on a utility server (certificate
auth). Azure Automation schedules can't run more often than hourly. Each run:

- picks up Ready for IT and Scheduled requests;
- marks each one **In progress** before acting. If a run is ever interrupted part-way through, the
  next run sends that request to **Needs IT review** rather than repeating it, so no duplicate
  accounts;
- holds anything unapproved, edited by someone untrusted, or from a requester outside the
  authorised groups (Status becomes **Needs IT review**);
- creates, offboards or updates, then writes Status, the account name and a log line to the item;
- leaves future terminations as **Scheduled** until 6pm site time on the last day;
- never offboards more than 5 people in one run, or anyone on the protected list;
- carries on with the other requests if one fails, and records the error on the failed one.

**Statuses:** Submitted > Pending approval > Ready for IT > (Scheduled) > In progress > Completed.
The other outcomes are *Needs IT review* (a person has to look; the IT automation log on the item
says why), *Rejected* and *Cancelled*. To re-run a request after fixing the problem, set it back to
Ready for IT.

## 8. Test it end to end

1. Submit a **New hire, Full user** for a made-up person. Approve it in Teams. Within 15 minutes the
   item shows **Completed** with the account name, and the requester gets the email.
2. Submit a **Termination** for that test account with last day = today. It goes **Scheduled**,
   then **Completed** after 6pm site time. Check that sign-in is blocked and the mailbox is shared.
3. Submit a **New hire, Contact only** with a personal email. It appears in the address book.
4. Edit an approved item as a non-HR user. You shouldn't be able to. If you can, the lock step in
   the flow isn't working, and the script sends the item to **Needs IT review**.
5. Have someone outside the team submit a request. It goes to **Needs IT review**.
