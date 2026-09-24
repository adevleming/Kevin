# Test configuration used by Test-PaycomLifecycle.ps1.
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

    Input         = @{
        ReportName      = 'IT Current Employees'
        # 'Folder'     - newest *.csv in Path (local path, UNC share, or a synced library)
        # 'SharePoint' - newest *.csv in Folder of the document library DriveId
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
    }

    Onboarding    = @{
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

    Mail          = @{
        From     = 'it-automation@iac.aero'   # mailbox the app sends as (scope Mail.Send to it)
        ReportTo = 'it@iac.aero'
    }

    Tickets       = @{
        SendTo               = 'helpdesk@iac.aero'   # PLACEHOLDER: the address Desk365 turns into tickets
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
