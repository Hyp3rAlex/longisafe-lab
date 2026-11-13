Attribute VB_Name = "modCharts"
Option Explicit
Public Sub DrawCharts()
    Dim ws As Worksheet, arr() As Variant, i As Long, heads As Variant
    ClearSheet "Charts": Set ws = Sheet("Charts")
    heads = Array("time_s", "speed_mps", "requested_decel_mps2", "achieved_decel_mps2", "regen_force_N", "friction_force_N", "filtered_jerk_mps3")
    ReDim arr(1 To N + 1, 1 To 7)
    For i = 0 To 6: arr(1, i + 1) = heads(i): Next i
    For i = 1 To N
        arr(i + 1, 1) = D(i, CT): arr(i + 1, 2) = D(i, CV)
        If Has(CQ) Then arr(i + 1, 3) = D(i, CQ)
        arr(i + 1, 4) = -D(i, CA)
        If Has(CR) Then arr(i + 1, 5) = D(i, CR)
        If Has(CF) Then arr(i + 1, 6) = D(i, CF)
        If MetricsAllowed Then arr(i + 1, 7) = Jerk(i)
    Next i
    ws.Range("A1").Resize(N + 1, 7).Value2 = arr
    ws.Columns("A:G").ColumnWidth = 19
    ws.Range("I1").Value2 = "SYNTHETIC DATA / EDUCATIONAL TOOL - " & SafeText(ScenarioId)
    AddXY ws, "Forward speed", "m/s", 2, 0, 10
    AddXY ws, "Requested and achieved net deceleration", "m/s^2", 4, IIf(Has(CQ), 3, 0), 235
    If Has(CR) And Has(CF) Then AddXY ws, "Synthetic retarding force magnitudes", "N", 5, 6, 460
    If MetricsAllowed Then AddXY ws, "Filtered acceleration derivative (full record)", "m/s^3", 7, 0, 685
    ws.rows(1).Font.Bold = True
End Sub
Private Sub AddXY(ByVal ws As Worksheet, ByVal title As String, ByVal unit As String, ByVal col1 As Long, ByVal col2 As Long, ByVal y As Double)
    Dim chart As ChartObject, trace As Series
    Set chart = ws.ChartObjects.Add(ws.Range("I1").Left, y + 25, 620, 215)
    With chart.chart
        .ChartType = xlXYScatterLinesNoMarkers
        Set trace = .SeriesCollection.NewSeries
        trace.name = CStr(ws.Cells(1, col1).Value2)
        trace.XValues = ws.Range("A2").Resize(N, 1)
        trace.values = ws.Cells(2, col1).Resize(N, 1)
        If col2 > 0 Then
            Set trace = .SeriesCollection.NewSeries
            trace.name = CStr(ws.Cells(1, col2).Value2)
            trace.XValues = ws.Range("A2").Resize(N, 1)
            trace.values = ws.Cells(2, col2).Resize(N, 1)
        End If
        .HasTitle = True: .ChartTitle.text = title
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.text = "Time [s]"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.text = unit
        .HasLegend = col2 > 0
    End With
End Sub
