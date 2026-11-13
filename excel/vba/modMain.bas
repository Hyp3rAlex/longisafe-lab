Attribute VB_Name = "modMain"
Option Explicit
Public Sub SetupLongiSafe()
    Dim ws As Worksheet, i As Long, names As Variant, defaults As Variant, units As Variant
    Dim rows As Collection, btn As Button, captions As Variant, actions As Variant, specs As Variant
    names = Split("Home,Config,RawData,Checks,Metrics,Charts,ScenarioComparison,Requirements,Report", ",")
    For i = 0 To UBound(names): Set ws = Sheet(CStr(names(i))): Next i
    ClearSheet "Home": Set ws = Sheet("Home")
    For i = ws.Buttons.count To 1 Step -1: ws.Buttons(i).Delete: Next i
    ws.Range("A1").Value2 = "LongiSafe Lab | " & LS_VERSION
    ws.Range("A2").Value2 = "SYNTHETIC DATA / EDUCATIONAL TOOL"
    ws.Range("A3").Value2 = "No vehicle actuation, BMW calibration, ISO compliance or safety guarantee."
    ws.Range("A4").Value2 = "VBA source is supplied for Windows Excel testing; no desktop Excel run was performed in Work."
    ws.Range("A5").Value2 = "RawData is an imported snapshot. Edit the CSV and re-import; worksheet edits do not change in-memory analysis."
    ws.Range("A6").Value2 = "Compare records each completed scenario as a labeled snapshot. Import and analyze another to append."
    ws.Range("A7").Value2 = "Config controls pedagogical thresholds. Changed thresholds take effect on the next Run Checks."
    captions = Array("Import CSV", "Run Checks", "Analyze Braking", "Compare Scenarios", "Export Report", "Reset")
    actions = Array("ImportOnly", "AnalyzeActive", "AnalyzeActive", "CompareActive", "ExportActive", "ResetActive")
    For i = 0 To 5
        Set btn = ws.Buttons.Add(20, 155 + i * 32, 240, 26)
        btn.Caption = captions(i): btn.OnAction = "'" & ThisWorkbook.name & "'!" & actions(i)
    Next i
    ws.Columns("A").ColumnWidth = 130: ws.rows("1:7").RowHeight = 20: ws.rows("1:7").WrapText = True
    ws.Range("A1:A2").Font.Bold = True
    If Sheet("Config").Range("A1").Value2 = "" Then
        defaults = Array(0.05, 3#, 0.2, 0.3, 0.7, 0.005, 0.5, 0.6, 0.2, 0.2, 0.2, 100#, 20#)
        names = Split("nominal_dt_s,gap_factor,filter_window_s,evidence_s,coherence_mps2,stale_speed_mps,motion_accel_mps2,decel_error_mps2,response_delay_s,min_request_mps2,min_speed_mps,max_speed_mps,max_abs_accel_mps2", ",")
        units = Array("s", "1", "s", "s", "m/s^2", "m/s", "m/s^2", "m/s^2", "s", "m/s^2", "m/s", "m/s", "m/s^2")
        Set rows = New Collection
        For i = 0 To 12
            rows.Add Array(names(i), defaults(i), units(i), "Illustrative test threshold; not prescribed by BMW/ISO")
        Next i
        WriteTable "Config", Array("parameter", "value", "units", "rationale"), rows
        Sheet("Config").Range("B2:B14").Interior.Color = RGB(235, 245, 255)
    End If
    Set rows = New Collection
    rows.Add Array("REQ-DATA-001", "MON-001", "S11", "Strict time import / rejection")
    rows.Add Array("REQ-PLAUS-002", "MON-003/MON-004", "S09/S10", "Checks evidence windows")
    rows.Add Array("REQ-BRAKE-003", "MON-005", "S04/S05/S14", "Requested vs achieved net decel")
    rows.Add Array("REQ-CONTEXT-004", "No zero-pedal fault heuristic", "S02", "Pedal alone not force")
    rows.Add Array("REQ-EVID-005", "All M1 rules", "All", "CSV row/time evidence; input SHA-256 and config snapshot")
    WriteTable "Requirements", Array("requirement", "monitor", "scenarios", "evidence"), rows
    Sheet("Home").Activate
End Sub
Public Sub ImportOnly()
    Dim picked As Variant, oldScreen As Boolean, oldEvents As Boolean, en As Long, et As String
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    picked = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select synthetic SI input")
    If VarType(picked) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False: Application.EnableEvents = False
    LoadCsv CStr(picked)
    Sheet("Home").Range("A9").Value2 = "Imported " & N & " rows: " & SafeText(ScenarioId) & " - run analysis"
Clean:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.StatusBar = False
    If en <> 0 Then MsgBox et, vbExclamation, "LongiSafe input rejected"
    Exit Sub
Failed:
    en = Err.Number: et = Err.Description
    Sheet("Home").Range("A9").Value2 = "INPUT REJECTED: " & SafeText(et)
    Resume Clean
End Sub
Public Sub AnalyzeActive()
    Dim oldScreen As Boolean, oldEvents As Boolean, en As Long, et As String
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    If Not DataReady Then Err.Raise 5, , "Import a valid CSV first"

Debug.Print "START:", DataReady, N

AnalyzeChecks
Debug.Print "AFTER CHECKS:", Rules.count, MetricsAllowed

ComputeMetrics
Debug.Print "AFTER METRICS:", Rules.count, MetricRows.count

DrawCharts
Debug.Print "AFTER CHARTS:", Rules.count, MetricRows.count

MakeReport
Debug.Print "AFTER REPORT:", Rules.count, MetricRows.count

    Sheet("Home").Range("A9").Value2 = "Analysis complete: " & SafeText(ScenarioId) & " | Inspect Checks, Metrics and Charts"
    Sheet("Checks").Activate
Clean:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.StatusBar = False
    If en <> 0 Then MsgBox et, vbExclamation, "LongiSafe analysis error"
    Exit Sub
Failed:
    en = Err.Number: et = Err.Description
    Set Rules = New Collection: Set MetricRows = New Collection: MetricsAllowed = False
    ClearSheet "Checks": ClearSheet "Metrics": ClearSheet "Charts": ClearSheet "Report"
    Sheet("Home").Range("A9").Value2 = "ANALYSIS FAILED: " & SafeText(et)
    Resume Clean
End Sub
Public Sub ExportActive()
    On Error GoTo Failed
    ExportOutputs
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "LongiSafe export error"
End Sub
Public Sub CompareActive()
    On Error GoTo Failed
RecordComparison:     Sheet("ScenarioComparison").Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "LongiSafe comparison error"
End Sub
Public Sub ResetActive()
    Dim names As Variant, i As Long
    DataReady = False: MetricsAllowed = False: N = 0
    Set Rules = New Collection: Set MetricRows = New Collection
    names = Split("RawData,Checks,Metrics,Charts,Report,ScenarioComparison", ",")
    For i = 0 To UBound(names): ClearSheet CStr(names(i)): Next i
    Sheet("Home").Range("A9").Value2 = "Reset - import synthetic CSV"
End Sub
