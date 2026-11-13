Attribute VB_Name = "modCore"
Option Explicit
Public Const LS_VERSION As String = "0.1.0-m1"
Public D As Variant, Headers As Variant
Public N As Long, ColCount As Long, SourcePath As String, ScenarioId As String
Public InputSha256 As String, RunConfigSha256 As String
Public DataReady As Boolean, MetricsAllowed As Boolean
Public Rules As Collection, MetricRows As Collection
Public Cfg(1 To 13) As Double
Public Const CT As Long = 1, CV As Long = 2, CA As Long = 3
Public Const CP As Long = 4, CB As Long = 5, CG As Long = 6
Public Const CD As Long = 7, CR As Long = 8, CF As Long = 9, CQ As Long = 10
' Canonical indices; missing optional channels stay Empty, never zero-filled.
Public Function FieldNames() As Variant
    FieldNames = Split("time_s,vehicle_speed_mps,accel_long_mps2,accelerator_pedal_pct,brake_pedal_pct,road_grade_rad,drive_force_N,regen_force_N,friction_force_N,requested_decel_mps2,regen_limit_N,soc_pct,battery_temp_C,estimated_mu,adas_brake_request_mps2", ",")
End Function
Public Function Sheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = name Then Set Sheet = ws: Exit Function
    Next ws
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = name
    Set Sheet = ws
End Function
Public Sub ClearSheet(ByVal name As String)
    Dim ws As Worksheet, i As Long
    Set ws = Sheet(name)
    For i = ws.ChartObjects.count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    ws.Cells.Clear
End Sub
Public Sub WriteTable(ByVal name As String, ByVal heads As Variant, ByVal rows As Collection)
    Dim ws As Worksheet, arr() As Variant, i As Long, j As Long, r As Variant, width As Long
    Set ws = Sheet(name): ws.Cells.Clear
    width = UBound(heads) - LBound(heads) + 1
    ReDim arr(1 To rows.count + 1, 1 To width)
    For j = 1 To width: arr(1, j) = heads(j - 1 + LBound(heads)): Next j
    For i = 1 To rows.count
        r = rows(i)
        For j = 1 To width: arr(i + 1, j) = r(j - 1): Next j
    Next i
    ws.Range("A1").Resize(rows.count + 1, width).Value2 = arr
    With ws.Range("A1").Resize(1, width)
        .Font.Bold = True: .Interior.Color = RGB(224, 235, 244)
    End With
    ws.UsedRange.Columns.AutoFit
    For j = 1 To width
        If ws.Columns(j).ColumnWidth > 65 Then ws.Columns(j).ColumnWidth = 65
    Next j
    ws.UsedRange.WrapText = True
End Sub
Public Function Num(ByVal s As String) As Double
    ' Strict ASCII decimal/exponent grammar BEFORE Val. No numeric-prefix coercion.
    Dim i As Long, digits As Long, size As Long, ch As String, exponentDigits As Long
    size = Len(s): i = 1
    If size = 0 Then Err.Raise 5, , "Empty numeric cell"
    ch = Mid$(s, i, 1)
    If ch = "+" Or ch = "-" Then i = i + 1
    Do While i <= size
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        digits = digits + 1: i = i + 1
    Loop
    If Mid$(s, i, 1) = "." Then
        i = i + 1
        Do While i <= size
            ch = Mid$(s, i, 1)
            If ch < "0" Or ch > "9" Then Exit Do
            digits = digits + 1: i = i + 1
        Loop
    End If
    If digits = 0 Then Err.Raise 5, , "Invalid decimal: " & s
    ch = Mid$(s, i, 1)
    If ch = "e" Or ch = "E" Then
        i = i + 1: ch = Mid$(s, i, 1)
        If ch = "+" Or ch = "-" Then i = i + 1
        Do While i <= size
            ch = Mid$(s, i, 1)
            If ch < "0" Or ch > "9" Then Exit Do
            exponentDigits = exponentDigits + 1: i = i + 1
        Loop
        If exponentDigits = 0 Then Err.Raise 5, , "Invalid exponent: " & s
    End If
    If i <= size Then Err.Raise 5, , "Invalid numeric suffix: " & s
    Num = Val(s)
    If Abs(Num) > 1E+100 Then Err.Raise 5, , "Numeric magnitude outside supported domain"
End Function
Public Function Has(ByVal col As Long) As Boolean
    If DataReady Then Has = Not IsEmpty(D(1, col))
End Function
Public Sub LoadConfig()
    Dim vals As Variant, i As Long, expected As Variant
    expected = Split("nominal_dt_s,gap_factor,filter_window_s,evidence_s,coherence_mps2,stale_speed_mps,motion_accel_mps2,decel_error_mps2,response_delay_s,min_request_mps2,min_speed_mps,max_speed_mps,max_abs_accel_mps2", ",")
    vals = Sheet("Config").Range("B2:B14").Value2
    For i = 1 To 13
        If Sheet("Config").Cells(i + 1, 1).Value2 <> expected(i - 1) Then Err.Raise 5, , "Config labels/order changed; restore from source"
        If IsEmpty(vals(i, 1)) Or VarType(vals(i, 1)) = vbString Or Not IsNumeric(vals(i, 1)) Then Err.Raise 5, , "Config requires numeric values (row " & i + 1 & ")"
        Cfg(i) = CDbl(vals(i, 1))
        If Cfg(i) <= 0 Or Cfg(i) > 1E+100 Then Err.Raise 5, , "Config must be positive and finite"
    Next i
    If Cfg(2) <= 1 Or Cfg(12) <= Cfg(11) Then Err.Raise 5, , "Invalid Config ordering"
End Sub
Public Function SafeText(ByVal s As String) As String
    Dim ch As String
    ch = Left$(LTrim$(s), 1)
    If ch = "=" Or ch = "+" Or ch = "-" Or ch = "@" Then s = "'" & s
    SafeText = s
End Function
Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function
Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function
