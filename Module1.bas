Attribute VB_Name = "Module1"
Option Explicit

Public LogCsV As Boolean
Public Log_TempFile As Long
Public AppPath As String
Public Form2s() As Form2
Public Form5s() As Form5

Public Function GMi() As String
    Dim JsRun As Object
    Set JsRun = CreateObject("MSScriptControl.ScriptControl")
    JsRun.AllowUI = True
    JsRun.Language = "JavaScript"
    GMi = JsRun.Eval("Math.round(new Date().getTime())")
    Set JsRun = Nothing
End Function

Public Function sMid(zhong As String, Optional qian As String, Optional hou As String, Optional QnH As Integer = 0, Optional QHJ As Integer = 0, Optional QK As Integer = 0) As String
On Error Resume Next
DoEvents
Dim P1 As Double, P2 As Double
If zhong = "" Then sMid = "0": Exit Function
If qian <> "" And QHJ = 0 Then P1 = InStr(zhong, qian)
If qian <> "" And QHJ = 1 Then P1 = InStrRev(zhong, qian)
If qian = "" Then P1 = 1
If P1 = 0 And qian <> "" Then sMid = "1": Exit Function
If QnH = 0 And QK = 0 And hou <> "" Then P2 = InStr(zhong, hou)
If QnH = 0 And QK = 1 And hou <> "" Then P2 = InStr(P1 + Len(qian), zhong, hou)
If QnH = 1 And hou <> "" Then P2 = InStrRev(zhong, hou)
If P2 = 0 And hou <> "" Then sMid = "2": Exit Function
If P2 < P1 + Len(qian) And hou <> "" Then sMid = "0": Exit Function
If hou <> "" Then sMid = Mid(zhong, P1 + Len(qian), P2 - (P1 + Len(qian)))
If hou = "" Then sMid = Mid(zhong, P1 + Len(qian))
End Function

Public Function STByte(s As String) As Byte()
    On Error Resume Next
    Dim zc() As Byte
    Dim ZC2() As String
    ZC2 = Split(s, " ")
    Dim i As Long
    ReDim zc(UBound(ZC2)) As Byte
    For i = 0 To UBound(zc)
        zc(i) = H2D(ZC2(i))
    Next i
    STByte = zc
End Function

Public Function H2D(ByVal Hex As String) As Long
     Dim i As Long
     Dim B As Long
    
    Hex = UCase(Hex)
     For i = 1 To Len(Hex)
         Select Case Mid(Hex, Len(Hex) - i + 1, 1)
             Case "0": B = B + 16 ^ (i - 1) * 0
             Case "1": B = B + 16 ^ (i - 1) * 1
             Case "2": B = B + 16 ^ (i - 1) * 2
             Case "3": B = B + 16 ^ (i - 1) * 3
             Case "4": B = B + 16 ^ (i - 1) * 4
             Case "5": B = B + 16 ^ (i - 1) * 5
             Case "6": B = B + 16 ^ (i - 1) * 6
             Case "7": B = B + 16 ^ (i - 1) * 7
             Case "8": B = B + 16 ^ (i - 1) * 8
             Case "9": B = B + 16 ^ (i - 1) * 9
             Case "A": B = B + 16 ^ (i - 1) * 10
             Case "B": B = B + 16 ^ (i - 1) * 11
             Case "C": B = B + 16 ^ (i - 1) * 12
             Case "D": B = B + 16 ^ (i - 1) * 13
             Case "E": B = B + 16 ^ (i - 1) * 14
             Case "F": B = B + 16 ^ (i - 1) * 15
         End Select
     Next i
     H2D = B
End Function

Public Sub AddLog(s As String)
    Print #Log_TempFile, s
    If LogCsV = True Then
        If Form3.List1.ListCount >= 20000 Then Form3.List1.Clear
        Form3.List1.AddItem s
        Form3.List1.ListIndex = Form3.List1.ListCount - 1
    End If
End Sub

Public Function UTF8_URLEncoding(szInput)
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    If szInput = "" Then
        UTF8_URLEncoding = szInput
        Exit Function
    End If
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
       
        If nAsc < 0 Then nAsc = nAsc + 65536
       
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UTF8_URLEncoding = szRet
End Function
