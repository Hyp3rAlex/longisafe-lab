Attribute VB_Name = "modReporting"
Option Explicit
Public Sub MakeReport()
    Dim ws As Worksheet, dataWs As Worksheet, i As Long, row As Long, r As Variant, chart As ChartObject, trace As Series
    ClearSheet "Report": Set ws = Sheet("Report")
    ws.Range("A1").Value2 = "LongiSafe Lab | SYNTHETIC DATA / EDUCATIONAL TOOL"
    ws.Range("A2").Value2 = "Analysis result, not a vehicle safety conclusion | " & LS_VERSION
    ws.Range("A3").Value2 = "Scenario: " & SafeText(ScenarioId) & " | Local time: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A4").Value2 = "Input: " & SafeText(SourcePath)
    ws.Range("A5").Value2 = "Input SHA-256: " & InputSha256 & " | Config SHA-256 (key=value LF): " & RunConfigSha256
    ws.Range("A6").Value2 = "No BMW calibration, ISO compliance, ASIL allocation, real telemetry or safety guarantee."
    ws.Range("A7").Value2 = "Thresholds are illustrative. Missing channels and unimplemented M2 monitors remain NOT_EVALUATED."
    row = 9
    For Each r In MetricRows
        ws.Cells(row, 1).Value2 = r(0): ws.Cells(row, 2).Value2 = r(1)
        ws.Cells(row, 3).Value2 = r(2): ws.Cells(row, 4).Value2 = r(3): row = row + 1
    Next r
    row = row + 1
    For Each r In Rules
        ws.Cells(row, 1).Value2 = r(0): ws.Cells(row, 2).Value2 = r(2)
        ws.Cells(row, 3).Value2 = r(3): ws.Cells(row, 4).Value2 = r(4)
        ws.Cells(row, 5).Value2 = r(12): row = row + 1
    Next r
    ws.Columns("A").ColumnWidth = 28: ws.Columns("B:D").ColumnWidth = 16: ws.Columns("E").ColumnWidth = 72
    For i = 1 To 7
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 5)).Merge
    Next i
    ws.Range("A1:E7").WrapText = True: ws.rows("1:7").RowHeight = 30
    ws.UsedRange.WrapText = True: ws.UsedRange.rows.AutoFit
    ' Build the report chart directly, without copying through the Windows clipboard.
    ' ChartObject.Copy is unreliable during repeated automated smoke runs.
    If N > 0 Then
        Set dataWs = Sheet("Charts")
        Set chart = ws.ChartObjects.Add(ws.Cells(row + 1, 1).Left, _
                                        ws.Cells(row + 1, 1).Top, 650, 220)
        With chart.chart
            .ChartType = xlXYScatterLinesNoMarkers
            Set trace = .SeriesCollection.NewSeries
            trace.name = "achieved_decel_mps2"
            trace.XValues = dataWs.Range("A2").Resize(N, 1)
            trace.values = dataWs.Range("D2").Resize(N, 1)
            If Has(CQ) Then
                Set trace = .SeriesCollection.NewSeries
                trace.name = "requested_decel_mps2"
                trace.XValues = dataWs.Range("A2").Resize(N, 1)
                trace.values = dataWs.Range("C2").Resize(N, 1)
            End If
            .HasTitle = True
            .ChartTitle.text = "Requested and achieved net deceleration"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.text = "Time [s]"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.text = "m/s^2"
            .HasLegend = Has(CQ)
        End With
        row = row + 18
    End If
    With ws.PageSetup
        .Orientation = xlLandscape: .PaperSize = xlPaperA4
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
        .PrintArea = ws.Range("A1:E" & row).Address
    End With
    ws.Range("A1:E2").Font.Bold = True
    ' A pasted chart remains selected in Excel. Select a worksheet cell before PDF export.
    ws.Activate
    ws.Range("A1").Select
End Sub
Private Function CsvCell(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Err.Raise 5, "CsvCell", "Cannot export Excel error cell"
    ' Excel may read blank cells as Empty; they must stay blank, not become numeric zero.
    If IsEmpty(v) Or IsNull(v) Then
        s = ""
    ElseIf VarType(v) = vbString Then
        s = SafeText(CStr(v))
    ElseIf IsNumeric(v) Then
        s = Trim$(Str$(CDbl(v)))
    Else
        s = SafeText(CStr(v))
    End If
    CsvCell = Chr$(34) & Replace(s, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function
Public Sub SaveCsv(ByVal ws As Worksheet, ByVal path As String)
    Dim vals As Variant, i As Long, j As Long, line As String, f As Integer, opened As Boolean
    Dim en As Long, et As String
    On Error GoTo Failed
    ' A one-cell UsedRange is not the table that this exporter expects.
    If ws.UsedRange.Cells.CountLarge < 2 Then Err.Raise 5, "SaveCsv", "No populated table to export from " & ws.name
    vals = ws.UsedRange.Value2: f = FreeFile: Open path For Output As #f: opened = True
    For i = 1 To UBound(vals, 1)
        line = ""
        For j = 1 To UBound(vals, 2)
            If j > 1 Then line = line & ","
            line = line & CsvCell(vals(i, j))
        Next j
        Print #f, line
    Next i
    Close #f: Exit Sub
Failed:
    en = Err.Number: et = Err.Description
    If opened Then Close #f
    Err.Raise en, "SaveCsv", et
End Sub
Public Sub ExportOutputs()
    Dim picked As Variant, folder As String, slash As String, rows As Collection, i As Long
    If Rules Is Nothing Then Err.Raise 5, , "Run analysis first"
    If Rules.count = 0 Then Err.Raise 5, , "Run analysis first"
    picked = Application.GetSaveAsFilename("LongiSafe-summary.pdf", "PDF files (*.pdf), *.pdf")
    If VarType(picked) = vbBoolean Then Exit Sub
    slash = Application.PathSeparator: folder = Left$(CStr(picked), InStrRev(CStr(picked), slash))
    MakeReport
    Sheet("Report").ExportAsFixedFormat xlTypePDF, CStr(picked)
    ExportTables folder
End Sub
Public Sub ExportTables(ByVal folder As String)
    Dim rows As Collection, i As Long
    SaveCsv Sheet("Checks"), folder & "rule_results.csv"
    SaveCsv Sheet("Metrics"), folder & "metrics.csv"
    Set rows = New Collection
    For i = 1 To 13: rows.Add Array(Sheet("Config").Cells(i + 1, 1).Value2, Cfg(i)): Next i
    WriteTable "RunConfig", Array("parameter", "value"), rows
    SaveCsv Sheet("RunConfig"), folder & "run_config.csv"
    Set rows = New Collection
    rows.Add Array("software_version", LS_VERSION): rows.Add Array("source_path", SafeText(SourcePath))
    rows.Add Array("scenario_id", SafeText(ScenarioId)): rows.Add Array("rows", N)
    rows.Add Array("sha256", InputSha256)
    rows.Add Array("config_sha256_key_value_lf", RunConfigSha256)
    WriteTable "RunMetadata", Array("key", "value"), rows
    SaveCsv Sheet("RunMetadata"), folder & "run_metadata.csv"
End Sub
Public Sub RecordComparison()
    Dim ws As Worksheet, row As Long, r As Variant, v As Variant, i As Long
    If MetricRows Is Nothing Then Err.Raise 5, , "Analyze scenario first"
    Set ws = Sheet("ScenarioComparison")
    If ws.Cells(1, 1).Value2 = "" Then
        ws.Range("A1:E1").Value2 = Array("scenario_id", "local_snapshot_time", "metric", "value", "units")
    End If
    row = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    For Each r In MetricRows
        ws.Cells(row, 1).Value2 = SafeText(ScenarioId): ws.Cells(row, 2).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(row, 3).Value2 = r(0): ws.Cells(row, 4).Value2 = r(1): ws.Cells(row, 5).Value2 = r(2)
        row = row + 1
    Next r
    ws.Columns("A:E").AutoFit
End Sub
