Option Explicit

' ============================================================
' Shift Roster Generator
' Reads Config + Rules, generates 1-2 weeks of roster into
' the "Roster" sheet, tracking fairness in "History".
' Day order used throughout: 1=Sat 2=Sun 3=Mon 4=Tue 5=Wed 6=Thu 7=Fri
' ============================================================

Const NUM_MEMBERS As Integer = 9

Dim MemberNames(1 To NUM_MEMBERS) As String
Dim MemberRole(1 To NUM_MEMBERS) As String   ' "Lead" or "Member"

Dim RuleDay(1 To 7) As String
Dim RuleF1(1 To 7) As Integer
Dim RuleF2(1 To 7) As Integer
Dim RuleG(1 To 7) As Integer
Dim RuleNS(1 To 7) As Integer
Dim RuleGCond(1 To 7) As Boolean
Dim RuleF1ShortfallGBonus(1 To 7) As Boolean

Dim HistWO(1 To NUM_MEMBERS) As Long
Dim HistF1(1 To NUM_MEMBERS) As Long
Dim HistF2(1 To NUM_MEMBERS) As Long
Dim HistG(1 To NUM_MEMBERS) As Long
Dim HistNS(1 To NUM_MEMBERS) As Long

' --- Shift-adjacency / rest-cycle state, carried day to day across the whole run ---
Const MAX_CONSEC_WORKDAYS As Integer = 5   ' after this many days in a row, force 2 WO
Const MAX_NS_STREAK As Integer = 5         ' longest a member can be kept on NS in a row
Const MIN_NS_STREAK As Integer = 2         ' target: once put on NS, try to keep them there at least this
                                            ' long. NOT hard-enforced (a member CAN rotate off NS after 1
                                            ' day if that's the only way to cover a shift) - FillShift's
                                            ' scoring bonus for LastShift=NS members is a soft nudge toward
                                            ' this, not a guarantee.

Dim LastShift(1 To NUM_MEMBERS) As String  ' shift code assigned the previous day ("" before day 1)
Dim ConsecWork(1 To NUM_MEMBERS) As Integer  ' consecutive working days ending yesterday
Dim NSStreak(1 To NUM_MEMBERS) As Integer    ' consecutive NS days ending yesterday
Dim RestOwed(1 To NUM_MEMBERS) As Integer    ' forced WO days still owed (2, once triggered)

Dim TravelName() As String
Dim TravelStart() As Date
Dim TravelEnd() As Date
Dim TravelCount As Integer

' --- Snapshot of rest/streak state taken at the START of each week, used to
' replay it forward correctly after LevelWeek may have changed some WO->GN. ---
Dim SnapLastShift(1 To NUM_MEMBERS) As String
Dim SnapConsecWork(1 To NUM_MEMBERS) As Integer
Dim SnapNSStreak(1 To NUM_MEMBERS) As Integer
Dim SnapRestOwed(1 To NUM_MEMBERS) As Integer

Sub GenerateRoster()
    Dim startDate As Date
    Dim numWeeks As Integer
    Dim extraSun() As Boolean, extraMon() As Boolean
    Dim w As Integer, resp As String

    LoadConfig
    LoadRules
    LoadHistory
    LoadTravelPlan

    Dim i As Integer
    For i = 1 To NUM_MEMBERS
        LastShift(i) = "": ConsecWork(i) = 0: NSStreak(i) = 0: RestOwed(i) = 0
    Next i

    resp = InputBox("Enter roster start date (must be a Saturday), e.g. 26-Sep-2026:", _
                     "Roster Start Date", Format(NextSaturday(Date), "dd-mmm-yyyy"))
    If resp = "" Then Exit Sub
    If Not IsDate(resp) Then
        MsgBox "That's not a valid date.", vbExclamation
        Exit Sub
    End If
    startDate = CDate(resp)
    If Weekday(startDate, vbSaturday) <> 1 Then
        MsgBox "Start date must fall on a Saturday.", vbExclamation
        Exit Sub
    End If

    resp = InputBox("How many weeks to generate? (1 or 2)", "Weeks", "2")
    If resp = "" Then Exit Sub
    If Not IsNumeric(resp) Then
        MsgBox "Enter a number: 1 or 2.", vbExclamation
        Exit Sub
    End If
    numWeeks = CInt(resp)
    If numWeeks < 1 Or numWeeks > 2 Then
        MsgBox "Enter 1 or 2.", vbExclamation
        Exit Sub
    End If

    ReDim extraSun(1 To numWeeks)
    ReDim extraMon(1 To numWeeks)
    For w = 1 To numWeeks
        extraSun(w) = (UCase(InputBox("Week " & w & ": Extra work on Sunday? (Y/N)", "Extra Work - Sunday", "N")) = "Y")
        extraMon(w) = (UCase(InputBox("Week " & w & ": Extra work on Monday? (Y/N)", "Extra Work - Monday", "N")) = "Y")
    Next w

    Dim totalDays As Integer
    totalDays = numWeeks * 7

    Dim Assign() As String
    ReDim Assign(1 To NUM_MEMBERS, 1 To totalDays)

    Dim WorkCount(1 To NUM_MEMBERS) As Integer

    Dim d As Integer, wk As Integer, dow As Integer

    For d = 1 To totalDays
        dow = ((d - 1) Mod 7) + 1
        wk = Int((d - 1) / 7) + 1

        If dow = 1 Then
            For i = 1 To NUM_MEMBERS
                WorkCount(i) = 0
                ' Snapshot rest/streak state as it stood BEFORE this week's days
                ' were assigned, so it can be correctly replayed after LevelWeek
                ' (which may convert some WO days to GN) rather than left stale.
                SnapLastShift(i) = LastShift(i)
                SnapConsecWork(i) = ConsecWork(i)
                SnapNSStreak(i) = NSStreak(i)
                SnapRestOwed(i) = RestOwed(i)
            Next i
        End If

        AssignDay Assign, WorkCount, d, dow, wk, extraSun, extraMon, startDate

        If dow = 7 Then
            ' End of week (Sat..Fri): top up anyone under 5 workdays so every
            ' member lands on EXACTLY 2 Week Offs, not more. Any week where the
            ' Rules minimums add up to less than 8 members x 5 shifts leaves a
            ' surplus - without this pass that surplus always lands on the same
            ' 1-2 members (ties favour the lowest member index), so they'd
            ' silently rack up 3+ WO/week, every week.
            LevelWeek Assign, WorkCount, d - 6, d

            ' LevelWeek may have turned some WO days into GN. Replay the
            ' rest/streak state machine over the final post-leveling week so
            ' ConsecWork/NSStreak/RestOwed/LastShift reflect what was ACTUALLY
            ' scheduled, not the pre-leveling version AssignDay rolled forward
            ' day by day as it went.
            ReplayWeekState Assign, d - 6, d
        End If
    Next d

    WriteRoster Assign, startDate, totalDays
    SaveHistory

    MsgBox "Roster generated for " & totalDays & " days starting " & Format(startDate, "dd-mmm-yyyy") & ".", vbInformation
End Sub

Sub LevelWeek(ByRef Assign() As String, ByRef WorkCount() As Integer, weekStart As Integer, weekEnd As Integer)
    ' Convert surplus WO days into G (general/backup duty) so every member
    ' finishes the week at exactly 5 workdays = exactly 2 Week Offs -
    ' WITHOUT breaking the "2 consecutive WO after 4-5 straight shifts" rule.
    ' So this only converts an ISOLATED WO day (not sitting next to another
    ' WO day of theirs) - a paired rest day is never touched. If every WO
    ' day that member has is part of a mandatory pair, they're left with the
    ' extra WO rather than violate the rest rule.
    Dim i As Integer, d As Integer, converted As Boolean, isolated As Boolean
    Dim reqPass As Integer
    For i = 1 To NUM_MEMBERS
        ' Skip anyone who still owes mandatory rest days as of the START of
        ' this week - converting one of their WO days here could shorten a
        ' rest pair that started last week and crosses the week boundary,
        ' which this function's local (weekStart..weekEnd) adjacency check
        ' can't see.
        If MemberRole(i) = "Member" And SnapRestOwed(i) = 0 Then
            Do While WorkCount(i) < 5
                converted = False
                For d = weekStart To weekEnd
                    If Assign(i, d) = "WO" Then
                        isolated = True
                        If d > weekStart Then
                            If Assign(i, d - 1) = "WO" Then isolated = False
                        End If
                        If d < weekEnd Then
                            If Assign(i, d + 1) = "WO" Then isolated = False
                        End If
                        If isolated Then
                            Assign(i, d) = "GN"
                            WorkCount(i) = WorkCount(i) + 1
                            HistG(i) = HistG(i) + 1
                            HistWO(i) = HistWO(i) - 1
                            converted = True
                            Exit For
                        End If
                    End If
                Next d
                If Not converted Then Exit Do   ' only paired rest days left - leave as-is
            Loop
        End If
    Next i
End Sub

Sub ReplayWeekState(Assign() As String, weekStart As Integer, weekEnd As Integer)
    ' Re-derives ConsecWork/NSStreak/RestOwed/LastShift for this week from
    ' scratch, starting at the snapshot taken before the week began and
    ' replaying day by day against the FINAL Assign() values (i.e. after
    ' LevelWeek's WO->GN conversions), so the state carried into next week is
    ' accurate instead of stale from AssignDay's original first pass.
    Dim i As Integer, d As Integer
    For i = 1 To NUM_MEMBERS
        If MemberRole(i) = "Member" Then
            LastShift(i) = SnapLastShift(i)
            ConsecWork(i) = SnapConsecWork(i)
            NSStreak(i) = SnapNSStreak(i)
            RestOwed(i) = SnapRestOwed(i)

            For d = weekStart To weekEnd
                If IsRestCode(Assign(i, d)) Then
                    ConsecWork(i) = 0
                    If RestOwed(i) > 0 Then RestOwed(i) = RestOwed(i) - 1
                    NSStreak(i) = 0
                Else
                    ConsecWork(i) = ConsecWork(i) + 1
                    If Assign(i, d) = "NS" Then
                        NSStreak(i) = NSStreak(i) + 1
                    Else
                        NSStreak(i) = 0
                    End If
                    If ConsecWork(i) >= MAX_CONSEC_WORKDAYS Then RestOwed(i) = 2
                End If
                LastShift(i) = Assign(i, d)
            Next d
        End If
    Next i
End Sub

Function NextSaturday(d As Date) As Date
    Dim offset As Integer
    offset = (7 - Weekday(d, vbSaturday) + 1) Mod 7
    NextSaturday = d + offset
End Function

Sub LoadConfig()
    Dim ws As Worksheet, r As Long, i As Integer
    Set ws = ThisWorkbook.Sheets("Config")
    i = 1
    r = 2
    Do While ws.Cells(r, 1).Value <> "" And i <= NUM_MEMBERS
        MemberNames(i) = ws.Cells(r, 1).Value
        MemberRole(i) = ws.Cells(r, 2).Value
        i = i + 1
        r = r + 1
    Loop
End Sub

Sub LoadRules()
    Dim ws As Worksheet, r As Long, i As Integer
    Set ws = ThisWorkbook.Sheets("Rules")
    For i = 1 To 7
        r = i + 1
        RuleDay(i) = ws.Cells(r, 1).Value
        RuleF1(i) = ws.Cells(r, 2).Value
        RuleF2(i) = ws.Cells(r, 3).Value
        RuleG(i) = ws.Cells(r, 4).Value
        RuleNS(i) = ws.Cells(r, 5).Value
        RuleGCond(i) = (UCase(ws.Cells(r, 6).Value) = "Y")
        RuleF1ShortfallGBonus(i) = (UCase(ws.Cells(r, 7).Value) = "Y")
    Next i
End Sub

Sub LoadHistory()
    Dim ws As Worksheet, r As Long, i As Integer
    Set ws = ThisWorkbook.Sheets("History")
    For i = 1 To NUM_MEMBERS
        r = i + 1
        HistWO(i) = ws.Cells(r, 2).Value
        HistF1(i) = ws.Cells(r, 3).Value
        HistF2(i) = ws.Cells(r, 4).Value
        HistG(i) = ws.Cells(r, 5).Value
        HistNS(i) = ws.Cells(r, 6).Value
    Next i
End Sub

Sub SaveHistory()
    Dim ws As Worksheet, r As Long, i As Integer
    Set ws = ThisWorkbook.Sheets("History")
    For i = 1 To NUM_MEMBERS
        r = i + 1
        ws.Cells(r, 1).Value = MemberNames(i)
        ws.Cells(r, 2).Value = HistWO(i)
        ws.Cells(r, 3).Value = HistF1(i)
        ws.Cells(r, 4).Value = HistF2(i)
        ws.Cells(r, 5).Value = HistG(i)
        ws.Cells(r, 6).Value = HistNS(i)
    Next i
End Sub

Function GetCalendarCode(memberName As String, theDate As Date) As String
    ' Reads the "Calendar" sheet: row1 = dates from column D onward, column B = Name.
    ' Returns whatever's in that member's cell for that date ("" if none/sheet missing).
    GetCalendarCode = ""
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Calendar")
    On Error GoTo 0
    If wsC Is Nothing Then Exit Function

    Dim lastCol As Long, lastRow As Long, c As Long, r As Long, dateCol As Long
    lastCol = wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column
    lastRow = wsC.Cells(wsC.Rows.Count, 2).End(xlUp).Row

    dateCol = 0
    For c = 4 To lastCol   ' dates start at column D
        If IsDate(wsC.Cells(1, c).Value) Then
            If CDate(wsC.Cells(1, c).Value) = theDate Then
                dateCol = c
                Exit For
            End If
        End If
    Next c
    If dateCol = 0 Then Exit Function

    For r = 2 To lastRow
        If Trim(wsC.Cells(r, 2).Value) = Trim(memberName) Then
            GetCalendarCode = UCase(Trim(wsC.Cells(r, dateCol).Value))
            Exit Function
        End If
    Next r
End Function

Function IsRestCode(code As String) As Boolean
    IsRestCode = (code = "WO" Or code = "H" Or code = "L" Or code = "SL" Or code = "CO")
End Function

Function IsKnownCode(code As String) As Boolean
    ' Every code the Calendar sheet is allowed to lock in - anything else is
    ' treated as blank (see AssignDay) rather than as a valid override.
    IsKnownCode = (code = "H" Or code = "L" Or code = "SL" Or code = "CO" Or _
                   code = "F1" Or code = "F2" Or code = "GN" Or code = "NS" Or code = "WO")
End Function

Sub LoadTravelPlan()
    ' Reads the "TravelPlan" sheet: Name | StartDate | EndDate.
    ' A member with an active travel plan on a given date is allowed the F2
    ' exception straight after an NS shift (otherwise only WO or continuing
    ' NS is allowed the day after a night shift).
    Dim ws As Worksheet, r As Long, n As Integer
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TravelPlan")
    On Error GoTo 0
    If ws Is Nothing Then
        TravelCount = 0
        Exit Sub
    End If

    n = 0
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        n = n + 1
        r = r + 1
    Loop

    If n = 0 Then
        TravelCount = 0
        Exit Sub
    End If
    ReDim TravelName(1 To n)
    ReDim TravelStart(1 To n)
    ReDim TravelEnd(1 To n)

    ' Skip (don't load) any row with a missing/unparseable date rather than
    ' crashing the whole run on a blank or malformed TravelPlan row.
    Dim i As Integer, kept As Integer
    kept = 0
    For i = 1 To n
        r = i + 1
        If IsDate(ws.Cells(r, 2).Value) And IsDate(ws.Cells(r, 3).Value) Then
            kept = kept + 1
            TravelName(kept) = ws.Cells(r, 1).Value
            TravelStart(kept) = CDate(ws.Cells(r, 2).Value)
            TravelEnd(kept) = CDate(ws.Cells(r, 3).Value)
        End If
    Next i
    TravelCount = kept
End Sub

Function IsOnTravelPlan(i As Integer, theDate As Date) As Boolean
    Dim j As Integer
    IsOnTravelPlan = False
    For j = 1 To TravelCount
        If TravelName(j) = MemberNames(i) Then
            If theDate >= TravelStart(j) And theDate <= TravelEnd(j) Then
                IsOnTravelPlan = True
                Exit Function
            End If
        End If
    Next j
End Function

Sub AssignDay(ByRef Assign() As String, ByRef WorkCount() As Integer, d As Integer, dow As Integer, wk As Integer, extraSun() As Boolean, extraMon() As Boolean, startDate As Date)
    Dim i As Integer
    Dim theDate As Date
    theDate = startDate + d - 1

    ' Calendar overrides (pre-planned Holiday/Leave/Sick Leave/shift preference)
    ' are applied FIRST and take priority over everything else, including the
    ' Lead's normal fixed schedule.
    Dim calCode As String
    For i = 1 To NUM_MEMBERS
        calCode = GetCalendarCode(MemberNames(i), theDate)
        If calCode <> "" And IsKnownCode(calCode) Then
            Assign(i, d) = calCode
            If Not IsRestCode(calCode) Then
                WorkCount(i) = WorkCount(i) + 1
                AddHist i, calCode
            End If
        ElseIf calCode <> "" Then
            ' Unrecognized Calendar entry (typo, stray text, etc.) - ignore it
            ' and let the macro schedule the day normally rather than silently
            ' treating a typo as a valid lock-in.
        End If
    Next i

    ' Lead: fixed schedule, not part of the rota pool (unless overridden above)
    For i = 1 To NUM_MEMBERS
        If MemberRole(i) = "Lead" And Assign(i, d) = "" Then
            If dow >= 3 And dow <= 7 Then   ' Mon..Fri
                Assign(i, d) = "GN"
                HistG(i) = HistG(i) + 1
            Else                             ' Sat/Sun
                Assign(i, d) = "WO"
                HistWO(i) = HistWO(i) + 1
            End If
        End If
    Next i

    Dim needF1 As Integer, needF2 As Integer, needG As Integer, needNS As Integer
    needF1 = RuleF1(dow): needF2 = RuleF2(dow): needNS = RuleNS(dow)
    needG = 0
    If RuleG(dow) > 0 Then
        If Not RuleGCond(dow) Then
            needG = RuleG(dow)
        Else
            If dow = 2 And extraSun(wk) Then needG = RuleG(dow)   ' Sunday
            If dow = 3 And extraMon(wk) Then needG = RuleG(dow)   ' Monday
        End If
    End If

    ' Subtract anyone already locked in by the Calendar sheet, so the
    ' coverage minimums adjust automatically around pre-planned preferences.
    needF1 = needF1 - CountAssignedMembers(Assign, d, "F1")
    needF2 = needF2 - CountAssignedMembers(Assign, d, "F2")
    needNS = needNS - CountAssignedMembers(Assign, d, "NS")
    needG = needG - CountAssignedMembers(Assign, d, "GN")
    If needF1 < 0 Then needF1 = 0
    If needF2 < 0 Then needF2 = 0
    If needNS < 0 Then needNS = 0
    If needG < 0 Then needG = 0

    ' Fill hardest-to-staff first. NS goes first and prefers members already
    ' mid-streak, so a member put on NS tends to stay there a few days
    ' running rather than rotating out after just one.
    FillShift Assign, WorkCount, d, "NS", needNS, theDate
    FillShift Assign, WorkCount, d, "F2", needF2, theDate
    FillShift Assign, WorkCount, d, "F1", needF1, theDate

    ' If F1 ended up short of its target (e.g. only 1 got staffed instead of 2),
    ' put an extra member on GN to help cover, since the Lead is pulled into other
    ' assigned work that day.
    If RuleF1ShortfallGBonus(dow) Then
        Dim actualF1 As Integer
        actualF1 = CountAssignedMembers(Assign, d, "F1")
        If actualF1 < RuleF1(dow) Then
            needG = needG + (RuleF1(dow) - actualF1)
        End If
    End If

    FillShift Assign, WorkCount, d, "GN", needG, theDate

    ' Anyone left unassigned today is Week Off
    For i = 1 To NUM_MEMBERS
        If MemberRole(i) = "Member" And Assign(i, d) = "" Then
            Assign(i, d) = "WO"
            HistWO(i) = HistWO(i) + 1
        End If
    Next i

    ' Roll the per-member rest/streak state forward for tomorrow's eligibility checks.
    ' H/L/SL count as rest days too (streaks reset, doesn't count toward the cap)
    ' but only WO pays down a mandatory-rest debt in the usual sense - a rest code
    ' of any kind pays it down, since the member is off either way.
    For i = 1 To NUM_MEMBERS
        If MemberRole(i) = "Member" Then
            If IsRestCode(Assign(i, d)) Then
                ConsecWork(i) = 0
                If RestOwed(i) > 0 Then RestOwed(i) = RestOwed(i) - 1
                NSStreak(i) = 0
            Else
                ConsecWork(i) = ConsecWork(i) + 1
                If Assign(i, d) = "NS" Then
                    NSStreak(i) = NSStreak(i) + 1
                Else
                    NSStreak(i) = 0
                End If
                If ConsecWork(i) >= MAX_CONSEC_WORKDAYS Then RestOwed(i) = 2
            End If
            LastShift(i) = Assign(i, d)
        End If
    Next i
End Sub

Function IsEligible(i As Integer, shiftCode As String, theDate As Date) As Boolean
    ' Mandatory rest after MAX_CONSEC_WORKDAYS days in a row - only WO allowed
    If RestOwed(i) > 0 Then
        IsEligible = (shiftCode = "WO")
        Exit Function
    End If
    If ConsecWork(i) >= MAX_CONSEC_WORKDAYS Then
        IsEligible = (shiftCode = "WO")
        Exit Function
    End If

    ' Coming off a Night Shift yesterday
    If LastShift(i) = "NS" Then
        Select Case shiftCode
            Case "F1"
                IsEligible = False          ' never possible straight after NS
                Exit Function
            Case "F2"
                IsEligible = IsOnTravelPlan(i, theDate)   ' only with a travel plan on file
                Exit Function
            Case "NS"
                IsEligible = (NSStreak(i) < MAX_NS_STREAK)
                Exit Function
        End Select
    End If

    IsEligible = True
End Function

Sub FillShift(ByRef Assign() As String, ByRef WorkCount() As Integer, d As Integer, shiftCode As String, needCount As Integer, theDate As Date)
    Dim filled As Integer, i As Integer, best As Integer, bestScore As Long, score As Long
    filled = 0
    Do While filled < needCount
        best = 0: bestScore = -1
        For i = 1 To NUM_MEMBERS
            If MemberRole(i) = "Member" And Assign(i, d) = "" And WorkCount(i) < 5 Then
                If IsEligible(i, shiftCode, theDate) Then
                    score = HistScore(i, shiftCode)
                    ' Prefer someone already mid-streak on NS so they run a
                    ' natural 2-5 day block instead of rotating daily.
                    If shiftCode = "NS" And LastShift(i) = "NS" And NSStreak(i) < MAX_NS_STREAK Then
                        score = score - 1000000
                    End If
                    If best = 0 Or score < bestScore Then
                        best = i: bestScore = score
                    End If
                End If
            End If
        Next i
        If best = 0 Then
            ' Not enough eligible members left this day - leave short-staffed
            Exit Do
        End If
        Assign(best, d) = shiftCode
        WorkCount(best) = WorkCount(best) + 1
        AddHist best, shiftCode
        filled = filled + 1
    Loop
End Sub

Function CountAssigned(Assign() As String, d As Integer, code As String) As Integer
    Dim i As Integer, cnt As Integer
    cnt = 0
    For i = 1 To NUM_MEMBERS
        If Assign(i, d) = code Then cnt = cnt + 1
    Next i
    CountAssigned = cnt
End Function

Function CountAssignedMembers(Assign() As String, d As Integer, code As String) As Integer
    ' Same as CountAssigned but excludes the Lead row - used when figuring out
    ' how many MORE members are needed for a shift, since the Lead's own GN
    ' assignment is separate from the member coverage target.
    Dim i As Integer, cnt As Integer
    cnt = 0
    For i = 1 To NUM_MEMBERS
        If MemberRole(i) = "Member" And Assign(i, d) = code Then cnt = cnt + 1
    Next i
    CountAssignedMembers = cnt
End Function

Function HistScore(i As Integer, shiftCode As String) As Long
    Select Case shiftCode
        Case "F1": HistScore = HistF1(i)
        Case "F2": HistScore = HistF2(i)
        Case "GN": HistScore = HistG(i)
        Case "NS": HistScore = HistNS(i)
    End Select
End Function

Sub AddHist(i As Integer, shiftCode As String)
    Select Case shiftCode
        Case "F1": HistF1(i) = HistF1(i) + 1
        Case "F2": HistF2(i) = HistF2(i) + 1
        Case "GN": HistG(i) = HistG(i) + 1
        Case "NS": HistNS(i) = HistNS(i) + 1
    End Select
End Sub

Sub WriteRoster(Assign() As String, startDate As Date, totalDays As Integer)
    ' Appends this run's days as NEW columns to the right of whatever is
    ' already in "Roster" - the sheet is never cleared, so every period ever
    ' generated stays visible in one continuous tab (row per member, growing
    ' rightward by date), instead of being overwritten on each run.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Roster")

    Dim i As Integer, d As Integer
    Dim isFirstRun As Boolean
    isFirstRun = (ws.Cells(1, 2).Value = "")

    Dim startCol As Long
    If isFirstRun Then
        ws.Cells(1, 2).Value = "NAME"
        For i = 1 To NUM_MEMBERS
            ws.Cells(i + 2, 1).Value = i
            ws.Cells(i + 2, 2).Value = MemberNames(i)
        Next i
        startCol = 3
    Else
        Dim lastCol As Long
        lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 3 Then lastCol = 2
        startCol = lastCol + 1

        ' Refuse to silently overwrite a date range already in the sheet -
        ' better to stop and let the person pick a later start date.
        Dim c As Long
        For c = 3 To lastCol
            If IsDate(ws.Cells(2, c).Value) Then
                If CDate(ws.Cells(2, c).Value) >= startDate And CDate(ws.Cells(2, c).Value) <= startDate + totalDays - 1 Then
                    MsgBox "This date range overlaps with what's already in the Roster sheet (around " & _
                           Format(CDate(ws.Cells(2, c).Value), "d-mmm-yyyy") & _
                           "). Pick a start date after the last generated day, or clear those columns by hand first.", vbExclamation
                    Exit Sub
                End If
            End If
        Next c
    End If

    For d = 1 To totalDays
        Dim col As Long
        col = startCol + d - 1
        ws.Cells(1, col).Value = Format(startDate + d - 1, "ddd")
        ws.Cells(2, col).Value = startDate + d - 1
        ws.Cells(2, col).NumberFormat = "d-mmm-yyyy"
        For i = 1 To NUM_MEMBERS
            Dim cel As Range
            Set cel = ws.Cells(i + 2, col)
            cel.Value = Assign(i, d)
            ColorCell cel, Assign(i, d)
        Next i
    Next d

    Dim baseRow As Integer
    baseRow = NUM_MEMBERS + 4
    Dim shiftCodes As Variant
    shiftCodes = Array("F1", "F2", "GN", "NS")
    Dim s As Integer
    For s = 0 To 3
        ws.Cells(baseRow + s, 2).Value = shiftCodes(s)
        For d = 1 To totalDays
            Dim col2 As Long
            col2 = startCol + d - 1
            ws.Cells(baseRow + s, col2).Formula = _
                "=COUNTIF(" & ws.Cells(3, col2).Address & ":" & ws.Cells(NUM_MEMBERS + 2, col2).Address & ",""" & shiftCodes(s) & """)"
        Next d
    Next s

    ws.Columns.AutoFit
End Sub

Sub ColorCell(c As Range, code As String)
    ' Colors matched by eye against the team's existing roster screenshots.
    ' Only WO uses white font (on black fill) - every other code uses black
    ' font regardless of how dark its fill is.
    ' Tweak the RGB() values here if you want an exact hex match - Home tab >
    ' Font Color / Fill Color > "More Colors" > Custom on a real cell will
    ' give you the exact values to paste in.
    Select Case code
        Case "F1": c.Interior.Color = RGB(153, 153, 0): c.Font.Color = vbBlack      ' dark olive/yellow
        Case "F2": c.Interior.Color = RGB(0, 176, 80): c.Font.Color = vbBlack       ' green
        Case "GN": c.Interior.Color = RGB(189, 215, 238): c.Font.Color = vbBlack    ' pale blue
        Case "NS": c.Interior.Color = RGB(204, 102, 255): c.Font.Color = vbBlack    ' magenta/orchid
        Case "WO": c.Interior.Color = RGB(0, 0, 0): c.Font.Color = vbWhite          ' black - the one exception
        Case "H":  c.Interior.Color = RGB(112, 48, 160): c.Font.Color = vbBlack     ' dark purple
        Case "L", "SL", "CO": c.Interior.Color = RGB(204, 204, 255): c.Font.Color = vbBlack   ' light lavender
        Case Else: c.Interior.ColorIndex = xlNone
    End Select
End Sub
