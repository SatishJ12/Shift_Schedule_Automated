# Shift Roster Automation

Excel/VBA macro that generates a fair, rule-based shift roster (F1, F2, GN,
NS, WO) for a shift-based team, from a small set of configurable coverage
rules.

## What's here

- `RosterGenerator.bas` - the VBA macro (import into a macro-enabled `.xlsm`).
- `Roster_Automation_Template.xlsx` - a blank starter workbook with all the
  sheets the macro expects (`Config`, `Team_Details`, `TravelPlan`,
  `Calendar`, `Rules`, `History`, `Roster`, `ReadMe`). Contains placeholder
  names only (`TeamLead`, `Member1`...`Member8`) - no real data.
- `Shift_Roster_Automation_Runbook.docx` - full runbook: setup, execution
  steps, prompt reference, current rules, output behavior, troubleshooting.

## Quick start

1. Open `Roster_Automation_Template.xlsx`, fill in `Team_Details` with your
   team's real names/emails/contacts (kept local - see Privacy below).
2. Review/edit `Rules` for your team's coverage minimums per day of week.
3. Import `RosterGenerator.bas` as a VBA module (Alt+F11 > Insert > Module),
   save as `.xlsm`.
4. Optionally fill in `Calendar` (holidays/leave/shift preferences) and
   `TravelPlan` before running.
5. Run the `GenerateRoster` macro (Alt+F8).

Full details are in the runbook `.docx`.

## Core rules implemented

- Per-day-of-week minimum headcounts for F1/F2/GN/NS, with conditional GN
  coverage and an F1-shortfall GN bonus.
- Lead has a fixed Mon-Fri GN / Sat-Sun WO schedule.
- Every member finishes each week at exactly 2 Week Offs (max 5 shifts/week).
- Shift-adjacency / rest-cycle rules: no F1 the day right after NS; F2 after
  NS only with an active TravelPlan entry; NS runs in natural 2-5 day
  streaks; 2 consecutive WO required after 4-5 consecutive working days,
  enforced correctly across week boundaries.
- `Calendar` sheet lets you pre-plan Holiday/Leave/Sick Leave/shift
  preference per member per date - the macro works the rota around it.
- `Roster` sheet is append-only - every period you generate stays visible
  in one continuous sheet rather than being overwritten.

## Privacy

The template and this repo intentionally use placeholder names
(`TeamLead`, `Member1`...`Member8`) with no real team data. `.gitignore`
excludes your working `.xlsm` so your filled-in copy (real names, contacts,
calendar entries) never gets committed or pushed by accident.
