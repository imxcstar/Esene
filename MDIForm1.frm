VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Esene"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10515
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Menu C_RWCD 
      Caption         =   "任务"
      Begin VB.Menu C_XJRW 
         Caption         =   "新建任务"
      End
      Begin VB.Menu C_DKRW 
         Caption         =   "打开任务"
      End
      Begin VB.Menu C_Close_All_Rw 
         Caption         =   "关闭所有任务"
      End
   End
   Begin VB.Menu CLog 
      Caption         =   "日志"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub C_Close_All_Rw_Click()
    If MsgBox("确认关闭所有任务吗？", vbYesNo) = vbYes Then
        Dim i As Integer
        For i = 1 To UBound(Form2s)
            Unload Form2s(i)
        Next i
        ReDim Form2s(0)
        For i = 1 To UBound(Form5s)
            Unload Form5s(i)
        Next i
        ReDim Form5s(0)
    End If
End Sub

Public Sub C_DKRW_Click()
    Dim zc As String
    zc = ShowFolderSelection(Me.hwnd, "请选择任务目录")
    If zc <> "" Then
        If Right(zc, 1) <> "\" Then zc = zc & "\"
        If Dir(zc & "c.s") = "" Then
            MsgBox "配置文件不存在！"
        Else
            Dim ZC2 As String
            ZC2 = sMid(Left(zc, Len(zc) - 1), "\", , , 1)
            zc = sMid(Left(zc, Len(zc) - 1), , "\", 1) & "\"
            Dim TempFile As Long
            TempFile = FreeFile
            Dim rw As RWJG
            Open zc & ZC2 & "\c.s" For Binary As #TempFile
            Get #TempFile, , rw
            Close #TempFile
            If App.Major & "." & App.Minor & "." & App.Revision <> rw.RWBB Then
                MsgBox "任务配置文件版本不一致，可能会导致兼容性问题！" & vbCrLf & "当前任务配置文件版本为:" & rw.RWBB & vbCrLf & "当前软件任务配置文件版本为:" & App.Major & "." & App.Minor & "." & App.Revision
            End If
            If rw.RWTYPE = 0 Then
                Dim ZC4 As New Form5
                ReDim Preserve Form5s(UBound(Form5s) + 1)
                Set Form5s(UBound(Form5s)) = ZC4
                Form5s(UBound(Form5s)).RWPath = zc
                Form5s(UBound(Form5s)).RWName = ZC2
                Form5s(UBound(Form5s)).RWTYPE = rw.RWTYPE
                Form5s(UBound(Form5s)).SFXJ = False
                Form5s(UBound(Form5s)).Show
            End If
            If rw.RWTYPE = 1 Then
                Dim ZC3 As New Form2
                ReDim Preserve Form2s(UBound(Form2s) + 1)
                Set Form2s(UBound(Form2s)) = ZC3
                Form2s(UBound(Form2s)).RWPath = zc
                Form2s(UBound(Form2s)).RWName = ZC2
                Form2s(UBound(Form2s)).RWTYPE = rw.RWTYPE
                Form2s(UBound(Form2s)).SFXJ = False
                Form2s(UBound(Form2s)).Show
            End If
        End If
    End If
End Sub

Public Sub C_XJRW_Click()
    Dim XRW As New Form6
    XRW.Show
    XRW.ZOrder
End Sub

Private Sub CLog_Click()
    LogCsV = True
    Form3.Show
    Form3.ZOrder
End Sub

Private Sub MDIForm_Activate()
    LogCsV = True
    Form3.Show
End Sub

Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & "-" & App.Major & "." & App.Minor & "." & App.Revision
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    ReDim Form2s(0)
    ReDim Form5s(0)
    Log_TempFile = FreeFile
    Open AppPath & "Log.txt" For Append As #Log_TempFile
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    C_Close_All_Rw_Click
    Close #Log_TempFile
    End
End Sub
