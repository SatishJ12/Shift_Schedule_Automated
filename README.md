# Shift Roster Automation

Excel/VBA macro that generates a fair, rule-based shift roster for a
shift-based team, from a small set of configurable coverage rules.

## What's here

- `RosterGenerator.bas` - the VBA macro (import into a macro-enabled `.xlsm`).
- `Roster_Automation_Template.xlsx` - a blank starter workbook with all the
  sheets the macro expects (team/roles, contact details, travel plans,
  pre-planned calendar entries, coverage rules, fairness history, and the
  generated roster output). Contains placeholder names only
  (`TeamLead`, `Member1`...`Member8`) - no real data.

## Quick start

1. Open `Roster_Automation_Template.xlsx`, fill in your team's details
   (kept local - see Privacy below).
2. Review/edit the coverage-rules sheet for your team's minimum headcount
   per shift, per day of week.
3. Import `RosterGenerator.bas` as a VBA module (Alt+F11 > Insert > Module),
   save as `.xlsm`.
4. Optionally fill in pre-planned entries (holidays, leave, shift
   preferences, travel plans) before running.
5. Run the roster-generation macro (Alt+F8).

## Core rules implemented

- Per-day-of-week minimum headcounts across the team's shift types, with
  conditional extra coverage and a shortfall-compensation rule for one
  shift type.
- A fixed schedule for the team lead role.
- Every member finishes each week at exactly 2 days off (capped shifts per
  week).
- Shift-adjacency / rest-cycle rules: no immediate back-to-back shift types
  that aren't physically feasible; night-shift assignments run in natural
  multi-day streaks rather than rotating daily; mandatory consecutive days
  off are enforced after several consecutive working days, correctly
  across week boundaries.
- A pre-planning calendar lets you lock in holidays, leave, sick leave, or
  a shift preference per member per date - the macro works the rota around
  it.
- The output sheet is append-only - every period you generate stays
  visible in one continuous sheet rather than being overwritten.

## Privacy

The template and this repo intentionally use placeholder names
(`TeamLead`, `Member1`...`Member8`) with no real team data. `.gitignore`
excludes your working `.xlsm` (and the runbook) so your filled-in copy
(real names, contacts, calendar entries) never gets committed or pushed
by accident.
