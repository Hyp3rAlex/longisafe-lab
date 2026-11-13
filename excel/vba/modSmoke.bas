Attribute VB_Name = "modSmoke"
Option Explicit
Public Sub SmokeTestM1()
    Dim picked As Variant, inputFolder As String, outputFolder As String, sid As Variant
    Dim cases As Variant, expected As String, actual As String, summary As Collection, folder As String
    Dim before As Double, elapsed As Double, oldScreen As Boolean, oldEvents As Boolean
    Dim en As Long, et As String, knownBytes() As Byte
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    If ThisWorkbook.path = "" Then Err.Raise 5, , "Save your own workbook as LongiSafeLab.xlsm first"
    picked = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Pick supplied S05.csv to locate synthetic fixtures")
    If VarType(picked) = vbBoolean Then Exit Sub
    inputFolder = Left$(CStr(picked), InStrRev(CStr(picked), Application.PathSeparator))
    outputFolder = ThisWorkbook.path & Application.PathSeparator & "smoke_exports"
    If Dir$(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    Application.ScreenUpdating = False: Application.EnableEvents = False
    knownBytes = StrConv("abc", vbFromUnicode)
    If Sha256Bytes(knownBytes) <> "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad" Then Err.Raise 5, , "SHA-256 known-answer test failed"
    SetupLongiSafe
    Set summary = New Collection
    cases = Array("S01", "S02", "S04", "S05", "S09", "S10", "S11", "S14", "MISSING_CONTEXT", "PERF_10001")
    For Each sid In cases
        expected = "NO FAIL"
        Select Case CStr(sid)
            Case "S09": expected = "MON-003 FAIL"
            Case "S10": expected = "MON-004 FAIL"
            Case "S11": expected = "MON-001 REJECT"
            Case "S14": expected = "MON-005 FAIL"
            Case "MISSING_CONTEXT": expected = "MON-005 NOT_EVALUATED"
        End Select
        before = Timer
        actual = RunSmokeCase(inputFolder & CStr(sid) & ".csv", CStr(sid), expected, outputFolder)
        elapsed = Timer - before: If elapsed < 0 Then elapsed = elapsed + 86400
        summary.Add Array(CStr(sid), expected, actual, elapsed, "Source-level expectation; Python crosscheck still required")
    Next sid
    WriteTable "SmokeResults", Array("scenario", "expected", "observed", "elapsed_s", "caveat"), summary
    SaveCsv Sheet("SmokeResults"), outputFolder & Application.PathSeparator & "smoke_results.csv"
    Sheet("SmokeResults").Activate
Clean:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.StatusBar = False
    If en <> 0 Then MsgBox et, vbExclamation, "LongiSafe smoke failed"
    Exit Sub
Failed:
    en = Err.Number: et = Err.Description
    Resume Clean
End Sub
Private Function RunSmokeCase(ByVal path As String, ByVal sid As String, ByVal expected As String, ByVal root As String) As String
    Dim r As Variant, found As Boolean, failCount As Long, folder As String, desc As String
    On Error GoTo Failed
    LoadCsv path
    If sid = "S11" Then Err.Raise 5, , "Expected bad time was accepted"

AnalyzeChecks
Debug.Print sid & " AFTER CHECKS | Rules=" & Rules.count & _
    " | A2=" & Sheet("Checks").Range("A2").text

ComputeMetrics
DrawCharts
MakeReport

Debug.Print sid & " BEFORE EXPORT | Rules=" & Rules.count & _
    " | A2=" & Sheet("Checks").Range("A2").text & _
    " | UsedRange=" & Sheet("Checks").UsedRange.Address

    If sid = "PERF_10001" And N <> 10001 Then Err.Raise 5, , "Expected 10001 rows"
    For Each r In Rules
        If r(2) = "FAIL" Then failCount = failCount + 1
        If CStr(r(0)) & " " & CStr(r(2)) = expected Then found = True
    Next r
    If expected = "NO FAIL" Then found = failCount = 0
    folder = root & Application.PathSeparator & sid
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    folder = folder & Application.PathSeparator
    ExportTables folder
    Sheet("Report").ExportAsFixedFormat xlTypePDF, folder & "summary.pdf"
    If found Then RunSmokeCase = "MATCH" Else RunSmokeCase = "MISMATCH"
    Exit Function
Failed:
    desc = Err.Description
    If sid = "S11" And InStr(desc, "MON-001 nonincreasing time") > 0 And Not DataReady Then
        RunSmokeCase = "MATCH (rejected: " & desc & ")"
    Else
        RunSmokeCase = "ERROR: " & desc
    End If
End Function
