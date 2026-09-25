# Painter crew scheduling: proposed design

Status: **proposal, not built yet.** The open questions at the end need answers first.

## Today

Each site GM builds the painter schedule by hand and shares it their own way:
- email to a distribution group of painters' personal addresses, which IT has to add as contacts;
- a printed schedule.

Most painters have no company mailbox or Teams, so anything behind a Microsoft 365 sign-in is
invisible to them.

## Goals

1. One place where GMs keep the schedule: who is on which crew, in which hangar, on which aircraft,
   on which dates or shifts.
2. Painters see it without a company account:
   - on the **UniFi displays** in the hangars;
   - on their own phone or email.
3. The crew roster maintains itself from the lifecycle requests. New painters appear when they
   start, and leavers drop off when their termination runs. Nobody re-types names or keeps
   distribution lists up to date.

## Proposal

```
Lifecycle request (painter, Started) ──► Crew Roster list ◄── GM edits crew / level
                                             │
Aircraft & Jobs list (tail, hangar, dates) ──┼──► Crew Assignments list (who, which job, when)
                                             ▼
               Publish-CrewSchedule.ps1 (every 15 min, same server as the lifecycle runner)
                 ├─► hangar display page per site ──► UniFi displays (hangar network only)
                 └─► weekly "your schedule" email to each painter's personal address
```

### Data: three SharePoint lists on a new "Crew Scheduling" site

All three are new objects; nothing existing changes.

| List | Columns | Kept up to date by |
|---|---|---|
| **Crew Roster** | Name, site, painter level, crew, shift, personal email, mobile (optional), active | **Automation.** A row is added when a painter request is marked *Started*, and set inactive when their termination runs. GMs set crew, level and shift. |
| **Aircraft & Jobs** | Tail number, customer / work order, site, hangar, planned in, planned out, status | GMs or planners |
| **Crew Assignments** | Roster person, job, start date, end date, shift, notes | GMs |

- **Where GMs work:** a Teams tab per site shows those three lists, filtered to that site. This
  replaces the spreadsheet or whiteboard.
- **Editing rights:** each GM can edit only their own site's rows.

### Painters without company accounts

Two channels, both driven by the lists:

- **Hangar displays (UniFi):**
  - The script builds a plain, auto-refreshing page for each hangar: aircraft in the bay, the
    crew on each shift, and the next 7 days.
  - The page shows first names and last initials only.
  - It is served from an internal web server (e.g. IIS on the automation server) that only the
    hangar network can reach. The schedule is never on the public internet, and the displays
    don't need a Microsoft 365 login.
  - The displays are then pointed at that URL in UniFi Connect.
- **Personal copy:**
  - Each painter gets a short weekly email with their own assignments, and another if their
    assignments change.
  - It goes to the personal address already on their contact.
  - This replaces the GMs' distribution groups. Text messages are possible later through a paid
    SMS service, if email isn't enough.

### Alternative: Microsoft Shifts

Shifts in Teams is built for this kind of frontline scheduling:
- a phone app;
- shift swaps and time-off requests.

But every painter would need their own account and a Frontline licence (F1 or F3, per user per
month), which is exactly what the current "Contact only" setup avoids. It's worth revisiting if
painters ever get accounts. The Crew Roster list would carry over.

## What the lifecycle automation already provides

- Painter new hires arrive as **Contact only** requests with site, title, start date and personal
  email.
- The start-day check tells us who actually *Started*.
- Terminations and no-shows remove the contact. The same run can set the roster row inactive, so
  a leaver disappears from the display the same day.

## Open questions

1. **Displays:**
   - Which UniFi displays are these: UniFi Connect Display (13"/27"), Display Cast, or something
     else?
   - Do they show a **website URL** in your UniFi Connect signage, or only images and video? If
     only images and video, the script can render the page to an image instead.
2. **Scheduling granularity:** by day, by shift (e.g. days/nights), or by aircraft for its whole
   visit?
3. **The GMs' current schedules:** a spreadsheet, a whiteboard, or something else? A copy of one,
   with names removed, would set the columns.
4. **Names on the display:** is first name plus last initial right? Or crew and badge number
   only?
5. **Hangars:** how many at each site, and what are they called?
6. **Who counts as a painter:** every *Contact only* hire, or particular job titles (e.g.
   "Painter", "Painter Helper", "Prep")?
