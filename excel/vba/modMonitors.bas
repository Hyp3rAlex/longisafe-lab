Attribute VB_Name = "modMonitors"
Option Explicit
Public Sub AddRule(ByVal id As String, ByVal status As String, ByVal reason As String, _
    Optional ByVal threshold As Variant = "", Optional ByVal measured As Variant = "", _
    Optional ByVal units As String = "", Optional ByVal prereq As String = "", _
    Optional ByVal first As Long = 0, Optional ByVal last As Long = 0, Optional ByVal detection As Long = 0)
    Dim ts As Variant, te As Variant, td As Variant, rs As Variant, re As Variant
    ts = "": te = "": td = "": rs = "": re = ""
    If first > 0 Then ts = D(first, CT): rs = first + 1
    If last > 0 Then te = D(last, CT): re = last + 1
    If detection > 0 Then td = D(detection, CT)
    Rules.Add Array(id, SafeText(ScenarioId), status, ts, te, td, rs, re, measured, threshold, units, prereq, reason, "Illustrative condition, not a vehicle safety conclusion")
End Sub
Public Sub CheckSustained(ByVal id As String, ByRef values() As Double, ByRef eligible() As Boolean, _
    ByVal limit As Double, ByVal reason As String, ByVal units As String, ByVal prereq As String, ByVal valid As Boolean)
    Dim i As Long, firstEligible As Long, lastEligible As Long, runStart As Long, detection As Long
    Dim events As Long, peak As Double, globalPeak As Double, hit As Boolean
    If Not valid Then AddRule id, "NOT_EVALUATED", "Invalid timing/range prevents evaluation", limit, , units, prereq: Exit Sub
    For i = 1 To N
        If eligible(i) Then
            If firstEligible = 0 Then firstEligible = i
            lastEligible = i: globalPeak = MaxD(globalPeak, values(i))
        End If
    Next i
    If firstEligible = 0 Then AddRule id, "NOT_EVALUATED", "Prerequisites absent", limit, , units, prereq: Exit Sub
    If D(lastEligible, CT) - D(firstEligible, CT) < Cfg(4) - 0.000000001 Then AddRule id, "NOT_EVALUATED", "Insufficient usable duration", limit, , units, prereq: Exit Sub
    For i = 1 To N + 1
        hit = False
        If i <= N Then hit = eligible(i) And values(i) > limit
        If hit Then
            If runStart = 0 Then runStart = i: peak = 0
            peak = MaxD(peak, values(i))
            If detection = 0 Then
                If D(i, CT) - D(runStart, CT) >= Cfg(4) - 0.000000001 Then detection = i
            End If
        ElseIf runStart > 0 Then
            If detection > 0 Then
                AddRule id, "FAIL", reason, limit, peak, units, prereq, runStart, i - 1, detection
                events = events + 1
            End If
            runStart = 0: detection = 0
        End If
    Next i
    If events = 0 Then AddRule id, "PASS", "Evaluated; no sustained threshold breach", limit, globalPeak, units, prereq
End Sub
Public Sub AnalyzeChecks()
    Dim i As Long, k As Long, firstGap As Long, lastGap As Long, firstBad As Long, lastBad As Long
    Dim maxGap As Double, maxAccel As Double, valid As Boolean, span As Double, av As Double
    Dim names As Variant
    Dim pref() As Double, residual() As Double, stale() As Double, errors() As Double
    Dim eligible() As Boolean, reqEligible() As Boolean, onset As Double, missing As String, r As Variant
    If Not DataReady Then Err.Raise 5, , "Import a valid CSV first"
LoadConfig:     RunConfigSha256 = ConfigDigest(): Set Rules = New Collection
    ReDim pref(1 To N): ReDim residual(1 To N): ReDim stale(1 To N): ReDim errors(1 To N)
    ReDim eligible(1 To N): ReDim reqEligible(1 To N)
    For i = 1 To N
        If D(i, CV) > Cfg(12) Or Abs(D(i, CA)) > Cfg(13) Then
            If firstBad = 0 Then firstBad = i
            lastBad = i
        End If
        maxAccel = MaxD(maxAccel, Abs(D(i, CA)))
        If i > 1 Then
            span = D(i, CT) - D(i - 1, CT): maxGap = MaxD(maxGap, span)
            pref(i) = pref(i - 1) + D(i - 1, CA) * span
            If span > Cfg(1) * Cfg(2) + 0.000000001 Then
                If firstGap = 0 Then firstGap = i - 1
                lastGap = i
            End If
        End If
    Next i
    If firstGap > 0 Then
        AddRule "MON-001", "WARN", "Large gap; downstream analysis withheld", Cfg(1) * Cfg(2), maxGap, "s", "time_s", firstGap, lastGap
    Else
        AddRule "MON-001", "PASS", "Strict timing accepted; no large gaps", Cfg(1) * Cfg(2), maxGap, "s", "time_s"
    End If
    If firstBad > 0 Then
        AddRule "MON-002", "FAIL", "Configured signal range exceeded", "Config speed/acceleration limits", maxAccel, "SI", "core fields", firstBad, lastBad
    Else
        AddRule "MON-002", "PASS", "Parser and configured ranges evaluated", "Config speed/acceleration limits", maxAccel, "SI", "core fields"
    End If
    valid = firstBad = 0 And firstGap = 0
    k = 1
    For i = 2 To N
        Do While k + 1 < i
            If D(i, CT) - D(k + 1, CT) < Cfg(3) - 0.000000001 Then Exit Do
            k = k + 1
        Loop
        span = D(i, CT) - D(k, CT)
        If span >= Cfg(3) - 0.000000001 Then
            eligible(i) = True: av = (pref(i) - pref(k)) / span
            residual(i) = Abs((D(i, CV) - D(k, CV)) / span - av)
            If Abs(D(i, CV) - D(k, CV)) <= Cfg(6) And Abs(av) > Cfg(7) Then stale(i) = 1
        End If
    Next i
    CheckSustained "MON-003", residual, eligible, Cfg(5), "Speed change contradicts integrated measured acceleration", "m/s^2", "time_s; vehicle_speed_mps; accel_long_mps2", valid
    CheckSustained "MON-004", stale, eligible, 0.5, "Near-stationary speed signal with independent measured acceleration", "boolean", "time_s; vehicle_speed_mps; accel_long_mps2", valid
    If Has(CQ) Then
        onset = -1
        For i = 1 To N
            If D(i, CQ) > Cfg(10) And D(i, CV) > Cfg(11) Then
                If onset < 0 Then onset = D(i, CT)
                reqEligible(i) = D(i, CT) - onset >= Cfg(9) - 0.000000001
                errors(i) = MaxD(0, D(i, CQ) + D(i, CA))
            Else
                onset = -1
            End If
        Next i
        CheckSustained "MON-005", errors, reqEligible, Cfg(8), "Requested net deceleration undershoot; cause not diagnosed", "m/s^2", "requested_decel_mps2; accel_long_mps2; vehicle_speed_mps", valid
    Else
        AddRule "MON-005", "NOT_EVALUATED", "Missing requested net deceleration; pedal is not force", Cfg(8), , "m/s^2", "requested_decel_mps2"
    End If
    AddRule "MON-006", "NOT_EVALUATED", "No calibrated pedal-to-force mapping; zero accelerator is not a fault criterion"
    AddRule "MON-007", "NOT_EVALUATED", "Force residual monitor reserved for M2"
    AddRule "MON-008", "NOT_EVALUATED", "Transition monitor reserved for M2; work fraction computed separately"
    AddRule "MON-009", "NOT_EVALUATED", "Terminal comfort screening reserved for M2; full-record jerk is available"
    If Not Has(CQ) Then missing = missing & "requested_decel_mps2;"
    names = FieldNames()
    For i = CG To CF
        If Not Has(i) Then missing = missing & names(i - 1) & ";"
    Next i
    If Len(missing) > 0 Then
        AddRule "MON-010", "NOT_EVALUATED", "Missing optional context: " & missing
    Else
        AddRule "MON-010", "PASS", "M1 context channels present; M2 monitors still unimplemented"
    End If
    MetricsAllowed = valid
    For Each r In Rules
        If (r(0) = "MON-003" Or r(0) = "MON-004") And r(2) = "FAIL" Then MetricsAllowed = False
    Next r
    WriteTable "Checks", Split("rule_id,scenario_id,status,t_start_s,t_end_s,detection_time_s,start_csv_row,end_csv_row,measured_value,threshold,units,prerequisite_fields,rationale,caveat", ","), Rules
End Sub
