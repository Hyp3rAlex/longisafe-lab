Attribute VB_Name = "modMetrics"
Option Explicit
Public FilteredAccel() As Double, Jerk() As Double
Public Sub AddMetric(ByVal name As String, ByVal value As Variant, ByVal units As String, _
    Optional ByVal status As String = "EVALUATED", Optional ByVal definition As String = "")
    MetricRows.Add Array(name, value, units, status, definition)
End Sub
Public Sub ComputeMetrics()
    Dim i As Long, k As Long, first As Long, finish As Long, peakIndex As Long
    Dim total As Double, distance As Double, h As Double, tau As Double, wr As Double, wf As Double
    Dim sumSq As Double, peakJ As Double, ds As Double, status As String
    Set MetricRows = New Collection
    If Not MetricsAllowed Then
        AddMetric "analysis_withheld", "", "", "NOT_EVALUATED", "Timing/range/coherence concern: inspect input before KPIs"
        WriteTable "Metrics", Array("metric", "value", "units", "status", "definition"), MetricRows
        Exit Sub
    End If
    ReDim FilteredAccel(1 To N): ReDim Jerk(1 To N)
    peakIndex = 1: k = 1
    For i = 1 To N
        If D(i, CA) < D(peakIndex, CA) Then peakIndex = i
        total = total + D(i, CA)
        Do While k < i
            If D(i, CT) - D(k, CT) <= Cfg(3) + 0.000000001 Then Exit Do
            total = total - D(k, CA): k = k + 1
        Loop
        FilteredAccel(i) = total / (i - k + 1)
        If i > 1 Then
            h = D(i, CT) - D(i - 1, CT)
            Jerk(i) = (FilteredAccel(i) - FilteredAccel(i - 1)) / h
            peakJ = MaxD(peakJ, Abs(Jerk(i))): sumSq = sumSq + Jerk(i) ^ 2 * h
        End If
        If first = 0 And D(i, CB) > 0 And D(i, CV) > Cfg(11) Then first = i
        If first > 0 And finish = 0 And i > first And D(i, CV) <= 0.000000001 Then finish = i
    Next i
    AddMetric "record_duration_s", D(N, CT) - D(1, CT), "s"
    AddMetric "peak_decel_mps2", MaxD(0, -D(peakIndex, CA)), "m/s^2", , "Raw observed acceleration, full record"
    AddMetric "peak_decel_time_s", D(peakIndex, CT), "s"
    AddMetric "jerk_peak_mps3", peakJ, "m/s^3", , "Trailing Config window arithmetic mean, full record including stop"
    AddMetric "jerk_rms_mps3", Sqr(sumSq / (D(N, CT) - D(1, CT))), "m/s^3", , "Time-weighted RMS, full record"
    If first = 0 Or finish = 0 Then
        AddMetric "stop_time_s", "", "s", "NOT_EVALUATED", "No complete brake-to-zero event"
        AddMetric "stop_distance_m", "", "m", "NOT_EVALUATED", "No complete brake-to-zero event"
    Else
        For i = first To finish - 2
            distance = distance + 0.5 * (D(i, CV) + D(i + 1, CV)) * (D(i + 1, CT) - D(i, CT))
        Next i
        i = finish - 1: h = D(finish, CT) - D(i, CT): tau = h
        If D(i, CA) < 0 Then tau = MinD(h, D(i, CV) / -D(i, CA))
        distance = distance + 0.5 * D(i, CV) * tau
        AddMetric "stop_time_s", D(i, CT) + tau - D(first, CT), "s", , "First brake pedal >0 to zero; final constant-decel interpolation"
        AddMetric "stop_distance_m", distance, "m", , "Observed speed trapezoids with partial final interval"
    End If
    If Has(CR) And Has(CF) Then
        For i = 1 To N - 1
            h = D(i + 1, CT) - D(i, CT)
            If D(i + 1, CV) <= 0.000000001 And D(i, CV) > 0 And D(i, CA) < 0 Then h = MinD(h, D(i, CV) / -D(i, CA))
            ds = 0.5 * (D(i, CV) + D(i + 1, CV)) * h
            wr = wr + D(i, CR) * ds: wf = wf + D(i, CF) * ds
        Next i
        AddMetric "regen_work_J", wr, "J", , "Mechanical retarding work, not battery energy"
        If wr + wf > 0 Then
            AddMetric "regen_work_fraction", wr / (wr + wf), "1"
        Else
            AddMetric "regen_work_fraction", "", "1", "NOT_EVALUATED"
        End If
    Else
        AddMetric "regen_work_J", "", "J", "NOT_EVALUATED"
        AddMetric "regen_work_fraction", "", "1", "NOT_EVALUATED"
    End If
    WriteTable "Metrics", Array("metric", "value", "units", "status", "definition"), MetricRows
End Sub
