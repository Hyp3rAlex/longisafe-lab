Attribute VB_Name = "modHash"
Option Explicit
' Windows CNG, native bcrypt.dll. VBA7/Windows desktop only, 32/64-bit handles.
' No add-in/COM package, shell command, network or private-key operations.
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef alg As LongPtr, ByVal algId As LongPtr, ByVal implementation As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal alg As LongPtr, ByRef hash As LongPtr, ByVal hashObject As LongPtr, ByVal objectSize As Long, ByVal secret As LongPtr, ByVal secretSize As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hash As LongPtr, ByRef data As Any, ByVal size As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hash As LongPtr, ByRef data As Any, ByVal size As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal alg As LongPtr, ByVal flags As Long) As Long
Public Function Sha256Bytes(ByRef bytes() As Byte) As String
    Dim provider As LongPtr, hash As LongPtr, result(0 To 31) As Byte
    Dim i As Long, text As String, algName As String, rc As Long, en As Long, et As String
    On Error GoTo Failed
    algName = "SHA256"
    If BCryptOpenAlgorithmProvider(provider, StrPtr(algName), 0, 0) <> 0 Then Err.Raise 5, , "SHA-256 provider unavailable"
    ' NULL object / size 0: CNG allocates and frees hash storage (Windows 7+).
    If BCryptCreateHash(provider, hash, 0, 0, 0, 0, 0) <> 0 Then Err.Raise 5, , "SHA-256 creation failed"
    If BCryptHashData(hash, bytes(LBound(bytes)), UBound(bytes) - LBound(bytes) + 1, 0) <> 0 Then Err.Raise 5, , "SHA-256 input failed"
    If BCryptFinishHash(hash, result(0), 32, 0) <> 0 Then Err.Raise 5, , "SHA-256 result failed"
    For i = 0 To 31: text = text & Right$("0" & Hex$(result(i)), 2): Next i
    Sha256Bytes = LCase$(text)
Clean:
    If hash <> 0 Then rc = BCryptDestroyHash(hash)
    If provider <> 0 Then rc = BCryptCloseAlgorithmProvider(provider, 0)
    If en <> 0 Then Err.Raise en, "Sha256Bytes", et
    Exit Function
Failed:
    en = Err.Number: et = Err.Description
    Resume Clean
End Function
Public Function ConfigDigest() As String
    Dim bytes() As Byte, text As String, i As Long
    For i = 1 To 13
        text = text & CStr(Sheet("Config").Cells(i + 1, 1).Value2) & "=" & Trim$(Str$(Cfg(i))) & vbLf
    Next i
    ' ASCII key=value LF stream; intentionally distinct from Python JSON hash.
    bytes = StrConv(text, vbFromUnicode)
    ConfigDigest = Sha256Bytes(bytes)
End Function
