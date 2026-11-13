Attribute VB_Name = "modCsvImport"
Option Explicit
Public Sub LoadCsv(ByVal path As String)
    Dim f As Integer, opened As Boolean, raw As String, b() As Byte, lines As Variant
    Dim names As Variant, parts As Variant, mapping() As Long, i As Long, j As Long, k As Long
    Dim count As Long, x As Double, found As Boolean, errNo As Long, errText As String
    On Error GoTo Failed
    DataReady = False: MetricsAllowed = False: SourcePath = path
    ' Invalidate previous output before attempting new input, including failures.
    ClearSheet "RawData": ClearSheet "Checks": ClearSheet "Metrics": ClearSheet "Charts": ClearSheet "Report"
    Set Rules = New Collection: Set MetricRows = New Collection
    LoadConfig
    f = FreeFile: Open path For Binary Access Read As #f: opened = True
    If LOF(f) = 0 Then Err.Raise 5, , "Empty file"
    ReDim b(0 To LOF(f) - 1): Get #f, , b: Close #f: opened = False
    InputSha256 = Sha256Bytes(b)
    i = 0
    If UBound(b) >= 2 Then
        If b(0) = 239 And b(1) = 187 And b(2) = 191 Then i = 3
    End If
    ' ASCII numeric subset is valid UTF-8; optional UTF-8 BOM consumed above.
    For j = i To UBound(b)
        If b(j) > 127 Or b(j) = 0 Then Err.Raise 5, , "Non-ASCII numeric CSV or NUL unsupported"
    Next j
    If i = 3 Then
        For j = 3 To UBound(b): b(j - 3) = b(j): Next j
        ReDim Preserve b(0 To UBound(b) - 3)
    End If
    raw = StrConv(b, vbUnicode)
    If InStr(raw, Chr$(34)) > 0 Then Err.Raise 5, , "Quoted input fields unsupported in M1 numeric CSV"
    raw = Replace(raw, vbCrLf, vbLf)
    If InStr(raw, vbCr) > 0 Then Err.Raise 5, , "Use LF or CRLF line endings"
    If Right$(raw, 1) = vbLf Then raw = Left$(raw, Len(raw) - 1)
    lines = Split(raw, vbLf)
    N = UBound(lines)
    If N < 2 Or N > Sheet("RawData").rows.count - 1 Then Err.Raise 5, , "Need 2 to 1048575 data rows"
    names = Split(lines(0), ","): Headers = FieldNames(): ColCount = UBound(names) + 1
    ReDim mapping(0 To UBound(names))
    For j = 0 To UBound(names)
        found = False
        For k = 0 To UBound(Headers)
            If names(j) = Headers(k) Then mapping(j) = k + 1: found = True: Exit For
        Next k
        If Not found Then Err.Raise 5, , "Unknown/misnamed header: " & names(j)
        For k = 0 To j - 1
            If names(k) = names(j) Then Err.Raise 5, , "Duplicate header: " & names(j)
        Next k
    Next j
    For k = 1 To 5
        found = False
        For j = 0 To UBound(mapping)
            If mapping(j) = k Then found = True
        Next j
        If Not found Then Err.Raise 5, , "Missing required header: " & Headers(k - 1)
    Next k
    ReDim D(1 To N, 1 To 15)
    For i = 1 To N
        parts = Split(lines(i), ",")
        If UBound(parts) <> UBound(names) Then Err.Raise 5, , "Width mismatch at CSV row " & i + 1
        For j = 0 To UBound(parts)
            D(i, mapping(j)) = Num(CStr(parts(j)))
        Next j
        If D(i, CT) < 0 Then Err.Raise 5, , "MON-001 negative time at CSV row " & i + 1
        If i > 1 Then
            If D(i, CT) <= D(i - 1, CT) Then Err.Raise 5, , "MON-001 nonincreasing time at CSV row " & i + 1
        End If
        If D(i, CV) < 0 Or D(i, CP) < 0 Or D(i, CP) > 100 Or D(i, CB) < 0 Or D(i, CB) > 100 Then Err.Raise 5, , "MON-002 speed/pedal range at CSV row " & i + 1
        For k = 7 To 11
            If Not IsEmpty(D(i, k)) Then
                If D(i, k) < 0 Then Err.Raise 5, , "MON-002 negative force/request at CSV row " & i + 1
            End If
        Next k
        If Not IsEmpty(D(i, 12)) Then
            If D(i, 12) < 0 Or D(i, 12) > 100 Then Err.Raise 5, , "SOC range"
        End If
        If Not IsEmpty(D(i, 14)) Then
            If D(i, 14) < 0 Or D(i, 14) > 2 Then Err.Raise 5, , "mu range"
        End If
        If Not IsEmpty(D(i, 15)) Then
            If D(i, 15) < 0 Then Err.Raise 5, , "Negative ADAS request"
        End If
        If Not IsEmpty(D(i, CG)) Then
            If Abs(D(i, CG)) >= 1.5707963267949 Then Err.Raise 5, , "Grade outside domain"
        End If
        If i Mod 1000 = 0 Then Application.StatusBar = "LongiSafe Lab: parsed " & i & " / " & N
    Next i
    DataReady = True
    ScenarioId = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    If InStrRev(ScenarioId, ".") > 0 Then ScenarioId = Left$(ScenarioId, InStrRev(ScenarioId, ".") - 1)
    With Sheet("RawData")
        For j = 0 To 14: .Cells(1, j + 1).Value2 = Headers(j): Next j
        .Range("A2").Resize(N, 15).Value2 = D
        .rows(1).Font.Bold = True: .Columns("A:O").ColumnWidth = 19
    End With
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    If opened Then Close #f
    DataReady = False: MetricsAllowed = False
    Err.Raise errNo, "LoadCsv", errText
End Sub
