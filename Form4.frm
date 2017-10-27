VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "脚本编辑-任务-"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12150
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   12150
   Begin VB.Menu C_J_WJ 
      Caption         =   "文件"
      Begin VB.Menu C_J_Save 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu C_J_Dmts 
         Caption         =   "代码提示"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu C_XSQ 
      Caption         =   "显示区"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RWName As String, RWPath As String
Dim CodeEditK As clsSS
Public XSQ As Form1
Public SFYCJ As Boolean

Private Sub C_J_Dmts_Click()
    CodeEditK.CodeTs XSQ
End Sub

Private Sub C_J_Save_Click()
    FileCopy RWPath & RWName & "\r.js", RWPath & RWName & "\backup\r.js." & GMi
    CodeEditK.SaveTFile RWPath & RWName & "\r.js"
End Sub

Private Sub C_XSQ_Click()
    XSQ.Show
End Sub

Private Sub Form_Load()
    SFYCJ = True
    Me.Caption = "脚本编辑-任务-" & RWName
    Set CodeEditK = New clsSS
    CodeEditK.AttachedWindows Me.hwnd
    CodeEditK.Js高亮
    If Dir(RWPath & RWName & "\r.js") = "" Then
        CodeEditK.SetText "var init=function(){" & vbCrLf & vbCrLf & "}" & vbCrLf & vbCrLf & "var run=function(data){" & vbCrLf & vbCrLf & "    return data;" & vbCrLf & "}"
        CodeEditK.SaveTFile RWPath & RWName & "\r.js"
    Else
        CodeEditK.OpenTFile RWPath & RWName & "\r.js"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    FileCopy RWPath & RWName & "\r.js", RWPath & RWName & "\backup\r.js." & GMi
    CodeEditK.SaveTFile RWPath & RWName & "\r.js"
    Me.Hide
End Sub

Private Sub Form_Resize()
    CodeEditK.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub SaveCode()
    If SFYCJ = True Then
        FileCopy RWPath & RWName & "\r.js", RWPath & RWName & "\backup\r.js." & GMi
        CodeEditK.SaveTFile RWPath & RWName & "\r.js"
    End If
End Sub

Public Function ToText() As String
    If SFYCJ = True Then
        ToText = CodeEditK.GetText
    Else
        Dim TempFile As Long
        Dim LoadBytes() As Byte
        TempFile = FreeFile
        Open RWPath & RWName & "\r.js" For Binary As #TempFile
        ReDim LoadBytes(0 To LOF(TempFile) - 1) As Byte
        Get #TempFile, , LoadBytes
        Close TempFile
        ToText = StrConv(LoadBytes, vbUnicode)
    End If
End Function
