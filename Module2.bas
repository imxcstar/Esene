Attribute VB_Name = "Module2"
Option Explicit

Public Function BytesToBstr(bytes)
    On Error GoTo CuoWu
    Dim SFCW As Boolean
    Dim Unicode As String
    If IsUTF8(bytes) Then                                   '如果不是UTF-8编码则按照GB2312来处理
        Unicode = "UTF-8"
    Else
        Unicode = "GB2312"
    End If
TG:
    Dim objstream As Object
    Set objstream = CreateObject("ADODB.Stream")
    With objstream
        .Type = 1
        .Mode = 3
        .Open
        If SFCW = False Then .Write bytes
        .position = 0
        .Type = 2
        .Charset = Unicode
        BytesToBstr = .ReadText
        .Close
    End With
    Exit Function
CuoWu:
    Unicode = "GB2312"
    SFCW = True
    GoTo TG
End Function

'判断网页编码函数
Private Function IsUTF8(bytes) As Boolean
    On Error GoTo CuoWu
    Dim i As Long, AscN As Long, Length As Long
    Length = UBound(bytes) + 1
    
    If Length < 3 Then
        IsUTF8 = False
        Exit Function
    ElseIf bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
        IsUTF8 = True
        Exit Function
    End If
    
    Do While i <= Length - 1
        If bytes(i) < 128 Then
            i = i + 1
            AscN = AscN + 1
        ElseIf (bytes(i) And &HE0) = &HC0 And (bytes(i + 1) And &HC0) = &H80 Then
            i = i + 2
            
        ElseIf i + 2 < Length Then
            If (bytes(i) And &HF0) = &HE0 And (bytes(i + 1) And &HC0) = &H80 And (bytes(i + 2) And &HC0) = &H80 Then
                i = i + 3
            Else
                IsUTF8 = False
                Exit Function
            End If
        Else
            IsUTF8 = False
            Exit Function
        End If
    Loop
    
    If AscN = Length Then
        IsUTF8 = False
    Else
        IsUTF8 = True
    End If
    Exit Function
CuoWu:
    IsUTF8 = False
End Function


