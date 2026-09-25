# Test configuration used by the tests in this folder (derived from config.example.psd1).
@{
    CompanyName   = 'IAC'

    # 'Cloud'  - accounts are mastered in Entra ID.
    # 'Hybrid' - accounts sync from on-prem AD: leavers are disabled in AD (needs the
    #            ActiveDirectory module on the host) and joiners get a ticket only.
    DirectoryMode = 'Cloud'

    # Snapshots, reports and state live here. Relative paths are relative to this file.
    StatePath     = './.test-state'
    KeepSnapshots = 26

    Graph         = @{
        TenantId              = '00000000-0000-0000-0000-000000000000'
        ClientId              = '00000000-0000-0000-0000-000000000000'
        CertificateThumbprint = ''          # app-only auth from a server / Task Scheduler
        UseManagedIdentity    = $false      # set $true when running as an Azure Automation runbook
        Environment           = 'Global'    # 'USGov' for GCC High
        IncludeSignInActivity = $true       # needs AuditLog.Read.All + Entra ID P1
    }

    # ---- One-time tenant setup (Setup-LifecycleTenant.ps1) ---------------------------
    Setup         = @{
        TeamName          = 'Hiring & Staffing Requests'
        TeamDescription   = 'Submit new hire, termination (separation) and role change requests. Owned by HR and IT.'
        MailNickname      = 'HiringStaffingRequests'
        # Owners see every request, manage who's on the team, and own the list. Members submit
        # requests and see only their own. Membership of this team = "hiring ability".
        Owners            = @(
            'adam.devleming@iac.aero'        # IT
            'aaron.lueker@iac.aero'          # IT
            'shelbea.bean@iac.aero'          # Payroll / HR (approver)
            'annamarie.gutierrez@iac.aero'   # HR
        )
        Members           = @(
            'Brian.Stamer@IAC.Aero'
            # HR and management: add their addresses here, one per line
        )
        AppName           = 'IT Lifecycle Automation'
        SenderDisplayName = 'IT Automation'
    }

    # ---- Form-driven requests (primary process) ------------------------------------
    Requests      = @{
        SiteId            = 'site-1'
        ListId            = 'list-1'                  # printed by New-LifecycleRequestList.ps1
        ListUrl           = 'https://leascorp.sharepoint.com/sites/HiringStaffingRequests/Lists/Employee%20Lifecycle%20Requests'
        SiteTimeZone      = 'Pacific Standard Time'   # the SharePoint site's regional setting
        DefaultTimeZone   = 'Pacific Standard Time'   # for sites not listed under Sites
        TerminationCutoff = '18:00'             # access ends at this time (site local) on the last day
        # Only members of these groups may submit (checked by the script; if empty, nothing runs).
        # Setup-LifecycleTenant.ps1 prints the Hiring & Staffing Requests team's group ID for this.
        AuthorizedGroupIds = @(
            'g-hiring'
        )
        RequireApproval   = $true               # the flow records ApprovedBy; immediate terminations skip approval
        # Accounts allowed to edit a request after it's submitted: HR approvers and the account
        # the flow's SharePoint connection runs as. Anything else edited last goes to IT review.
        TrustedEditors    = @(
            'shelbea.bean@iac.aero', 'annamarie.gutierrez@iac.aero'   # HR
            'adam.devleming@iac.aero', 'aaron.lueker@iac.aero'        # IT (to re-run a request after fixing it)
            'flows@iac.aero'
        )
        MaxOffboardPerRun = 5                   # more than this in one run are held for a person to check
        MaxOffboardPerDay = 15                  # and more than this in 24 hours
        # Check the list's version history: the change to Ready for IT must come from a trusted
        # editor, and no one else may change a field the automation acts on after that.
        VerifyApprovalHistory = $false
        # New hires who don't show up. On the start date (after CheckTime, site time) the hiring
        # manager and the site's OrientationContacts get one email per site asking who didn't start.
        NoShow            = @{
            CheckTime     = '10:00'
            DeleteAccount = $false              # $true also deletes the unused account (restorable for 30 days)
        }
        # HR's checklists, as tick-boxes on the request (from the New Employee and Separation Checklists).
        HrNewHireChecklist = @(
            'Integrity First results (California only)', 'Drug test results', 'Consumer background check results'
            'FAA Drug Abatement Division acknowledgment', 'Self-identification forms (AAP, disability, veteran)'
            'Copy of employment application (page 5 drug question answered NO)', 'Signed offer letter'
            'Substance Abuse Program pre-employment acknowledgement', 'Drug & Alcohol Policy acknowledgement', 'Resume'
            'IAC employment application', 'Consumer background & applicant disclosure', 'Relocation agreement (if applicable)'
            'I-9, IDs & E-Verify results', 'Job description signed', 'THCNP (Texas only)', 'Physical', 'Sign-on bonus'
            'Time edit sheet', 'Badge policy', 'Benefits enrollment in Paycom'
        )
        HrExitChecklist    = @('Exit interview/form', 'COBRA premiums', 'Life insurance portability/conversion paperwork'
            'Collect resignation letter', 'Unemployment flier', 'Remove from DOT pool', 'Transfer of personnel and benefits folders'
            'Separation documents emailed to employee', 'Badge log updated')
        PayrollExitChecklist = @('Notification of termination for garnishments', 'Remove from Pamir is (confirm name with Payroll)')
        BenefitsChoices    = @('Medical', 'Dental', 'Vision', 'Life', '401(k)')
        ReturnItems        = @('Laptop', 'Phone', 'Tablet', 'Keys and locks', 'Respirator', 'Harness', 'Boots', 'Tools (e.g. sander)', 'AMEX card')
        EquipmentChoices  = @('Laptop', 'Desktop', 'Monitor(s)', 'Docking station', 'Phone', 'Tablet', 'Badge / keys')
    }

    # Locations from HR's Separation Checklist (GEG FTW VCV AMA PDX PAE Irvine, Corporate Office).
    # Name is what the form shows and what goes into Entra officeLocation (matching the existing
    # 'Amarillo (AMA)' style). Code starts the Employee Status email subject ("AMA Term").
    #   GroupIds            site groups a new account joins (security or Microsoft 365 groups)
    #   ContactGroups       distribution lists a contact-only person (e.g. painter) joins
    #   BadgeOfficeEmails   told immediately about terminations so the badge is returned (e.g. the airport badge office)
    #   OrientationContacts site admins asked "did everyone start?" on the start date, with the hiring manager
    Sites         = @(
        @{ Name = 'Spokane (GEG)'; Code = 'GEG'; TimeZone = 'Pacific Standard Time'; OfficeLocation = 'Spokane (GEG)'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @() }
        @{ Name = 'Fort Worth (FTW)'; Code = 'FTW'; TimeZone = 'Central Standard Time'; OfficeLocation = 'Fort Worth (FTW)'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @('alyssa.jacobs@iac.aero') }
        @{ Name = 'Victorville (VCV)'; Code = 'VCV'; TimeZone = 'Pacific Standard Time'; OfficeLocation = 'Victorville (VCV)'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @() }
        @{ Name = 'Amarillo (AMA)'; Code = 'AMA'; TimeZone = 'Central Standard Time'; OfficeLocation = 'Amarillo (AMA)'
            GroupIds = @('g-ama')
            ContactGroups = @('ama-floor@iac.aero')
            BadgeOfficeEmails = @('badges@ama-airport.example')
            OrientationContacts = @('diane.mendez@iac.aero') }
        @{ Name = 'Portland (PDX)'; Code = 'PDX'; TimeZone = 'Pacific Standard Time'; OfficeLocation = 'Portland (PDX)'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @() }
        @{ Name = 'Everett (PAE)'; Code = 'PAE'; TimeZone = 'Pacific Standard Time'; OfficeLocation = 'Everett (PAE)'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @() }
        @{ Name = 'Irvine'; Code = 'Irvine'; TimeZone = 'Pacific Standard Time'; OfficeLocation = 'Irvine'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @() }
        @{ Name = 'Corporate Office'; Code = 'Corp'; TimeZone = 'Pacific Standard Time'; OfficeLocation = 'Corporate Office'
            GroupIds = @(); ContactGroups = @(); BadgeOfficeEmails = @(); OrientationContacts = @() }
    )

    Departments   = @('Operations', 'Quality', 'QC', 'Supply Chain', 'EHS/Facilities', 'Records', 'HR', 'Finance',
        'Sales', 'Engineering', 'IT', 'Administration')

    # "Computer access" on the form. Contact = $true means no account, just an address-book contact.
    AccessTypes   = @(
        @{ Name = 'Full user'; Description = 'Laptop/desktop user: email, Teams, Office apps'
            GroupIds = @('g-full') }                     # licence group, e.g. 'License - M365 Business Standard'
        @{ Name = 'Basic user'; Description = 'Email and Teams on web/mobile only'
            GroupIds = @() }                     # e.g. 'License - M365 Business Basic'
        @{ Name = 'Contact only'; Description = 'No account (e.g. painters): address-book contact only'
            Contact = $true }
    )

    Exchange      = @{
        # Needed for contacts and mailbox conversion. App needs Exchange.ManageAsApp + an Exchange role.
        AppId                 = '00000000-0000-0000-0000-000000000000'
        CertificateThumbprint = ''
        Organization          = 'leascorp.onmicrosoft.com'   # must be the tenant's .onmicrosoft.com domain
        UseManagedIdentity    = $false
        # Adding contacts to a distribution list you don't own needs this, plus the
        # 'Security Group Creation and Membership' role (included in Exchange Administrator).
        BypassGroupOwnerCheck = $false
        # Contacts are only removed if this automation created them (it tags them). $true also
        # removes contacts that were created by hand.
        RemoveUntaggedContacts = $false
    }

    # ---- Weekly Paycom audit (backstop) ---------------------------------------------
    # Paycom's push report email only links to the Report Center, so someone downloads the
    # CSV and saves it to the drop folder. The audit flags hires/terms nobody filed a form for.
    Input         = @{
        ReportName      = 'IT Current Employees'
        # 'Folder'     - newest *.csv in Path (local path or UNC share)
        # 'SharePoint' - newest *.csv in Folder of the document library DriveId
        # Keep the roster somewhere only IT (and whoever downloads it) can write: it holds everyone's
        # details, and a replaced file could fake terminations. Not the Hiring & Staffing team site.
        Source          = 'Folder'
        Path            = './.test-drop'
        DriveId         = 'b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
        Folder          = 'Paycom Roster'
        RemindWhenStale = $true
    }

    Roster        = @{
        # Map our field names to the column headers in the Paycom export.
        # Verify these against a real export: open the CSV and copy the headers exactly.
        Columns             = @{
            EmployeeId    = 'Employee_Code'           # required
            FirstName     = 'Legal_Firstname'         # required
            LastName      = 'Legal_Lastname'          # required
            PreferredName = 'Nickname'
            Email         = 'Work_Email'
            Department    = 'Department_Desc'
            JobTitle      = 'Position'
            Manager       = 'Supervisor_Primary'
            ManagerEmail  = 'Supervisor_Primary_Email'
            Location      = 'Location_Desc'
            Status        = 'Employee_Status'
            HireDate      = 'Hire_Date'
            TermDate      = 'Termination_Date'
        }
        # Status values that mean "currently employed". Anything else (T, Terminated,
        # Inactive...) counts as a leaver. Leave-of-absence counts as active by default.
        ActiveStatusPattern = '^(a|active|l|leave|loa|on leave)'
    }

    # Which Entra accounts Paycom is the source of truth for.
    Scope         = @{
        Domains         = @('iac.aero')           # Eirtech (etas.ie) etc. are out of scope
        ExcludeUpns     = @(
            # 'scanner@iac.aero'
        )
        # Regex against UPN and display name: service, shared, admin, room accounts.
        ExcludePatterns = @('^(admin|svc|service|scan|scanner|noreply|no-reply|conf|room|shared|test|sync|breakglass)([._-]|\d|@)', '\b(admin|service account|conference|mailbox)\b')
        RequireLicense  = $false
    }

    Safety        = @{
        MinRosterRows    = 5     # fewer active employees than this = bad export
        MaxShrinkPercent = 25     # active headcount drop larger than this = bad export
        MaxTerminations  = 3     # more leavers than this in one run = hold for review
    }

    Offboarding   = @{
        Enabled                = $false   # turn on after a few weeks of clean reports
        RemoveGroupMemberships = $false   # licence, dynamic and synced groups are always kept
        KeepGroupIds           = @()
        AddToGroupId           = ''       # e.g. an "Offboarded Users" group targeted by a CA block policy
        DisabledUsersOU        = ''       # Hybrid only, e.g. 'OU=Disabled Users,DC=iac,DC=local'
        ConvertMailboxToShared = $true    # form terminations: convert mailbox, give access to the named delegate
        ProtectedUpns          = @(       # never offboarded automatically
            'ceo@iac.aero'
        )
    }

    Onboarding    = @{
        # Enabled applies to the Paycom audit only. Leave it off: the form creates accounts.
        Enabled         = $false
        Domain          = 'iac.aero'
        UsageLocation   = 'US'
        CompanyName     = 'International Aerospace Coatings'
        DefaultGroupIds = @(
            # All-staff / licence group object IDs
        )
        # Department regex -> group IDs (group-based licensing, Teams, shared mailboxes)
        DepartmentGroups = @{
            # '^Sales'      = @('00000000-0000-0000-0000-000000000000')
            # '^Production' = @('00000000-0000-0000-0000-000000000000')
        }
        # Not every employee needs an M365 account. Regexes on the Paycom fields.
        Eligibility     = @{
            IncludeDepartments = ''
            ExcludeDepartments = '^Production'
            IncludeJobTitles   = ''
            ExcludeJobTitles   = ''
        }
    }

    BackfillEmployeeId = $true   # with -Apply, stamp Paycom employee code on matched accounts

    # Replaces the emails HR sent by hand: the Employee Status Notification (new hire, term, change,
    # no-show) in HR's format, plus the site badge office for terminations and HR for no-shows.
    Notifications = @{
        EmployeeStatusTo = @('employeestatus@iac.aero')
        HrTo             = @('shelbea.bean@iac.aero')
    }

    Mail          = @{
        From     = 'it-automation@iac.aero'   # mailbox the app sends as (scope Mail.Send to it)
        ReportTo = 'it@iac.aero'
    }

    Tickets       = @{
        SendTo               = 'helpdesk@iac.aero'   # PLACEHOLDER: the address Desk365 turns into tickets
        # When form requests are available, the Paycom audit only raises tickets for hires and
        # terminations that have no matching request (the form flow already raised the others).
        OnlyForGaps          = $true
        CreateChangeTickets  = $true
        ChangeTicketFields   = @('Department', 'JobTitle', 'Manager')
        OnboardingChecklist  = @(
            'Confirm account, display name and job title in Entra ID'
            'Licences / groups for the department (check group-based licensing applied)'
            'Laptop / monitors / peripherals ordered and imaged'
            'Enroll device in Intune'
            'Day one: issue a Temporary Access Pass in person; user registers MFA and sets a password'
            'Department apps, shared mailboxes and file shares'
            'Enroll in KnowBe4 security awareness training'
            'Confirm with manager that access is complete'
        )
        OffboardingChecklist = @(
            'Confirm sign-in is blocked and sessions revoked'
            'Convert mailbox to shared; grant manager access and set auto-reply if requested'
            'Remove licences after the mailbox is converted'
            'Remove MFA methods and registered devices'
            'Retire / wipe Intune devices; collect laptop, phone, badge and keys'
            'Remove from third-party apps (ERP, VPN, line-of-business apps)'
            'Transfer OneDrive files to manager'
            'Rotate any shared passwords the user knew'
        )
    }
}
