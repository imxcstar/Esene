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
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Menu C_RWCD 
      Caption         =   "����"
      Begin VB.Menu C_XJRW 
         Caption         =   "�½�����"
      End
      Begin VB.Menu C_DKRW 
         Caption         =   "������"
      End
      Begin VB.Menu C_Close_All_Rw 
         Caption         =   "�ر���������"
      End
   End
   Begin VB.Menu CLog 
      Caption         =   "��־"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub C_Close_All_Rw_Click()
    If MsgBox("ȷ�Ϲر�����������", vbYesNo) = vbYes Then
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
    zc = ShowFolderSelection(Me.hwnd, "��ѡ������Ŀ¼")
    If zc <> "" Then
        If Right(zc, 1) <> "\" Then zc = zc & "\"
        If Dir(zc & "c.s") = "" Then
            MsgBox "�����ļ������ڣ�"
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
                MsgBox "���������ļ��汾��һ�£����ܻᵼ�¼��������⣡" & vbCrLf & "��ǰ���������ļ��汾Ϊ:" & rw.RWBB & vbCrLf & "��ǰ������������ļ��汾Ϊ:" & App.Major & "." & App.Minor & "." & App.Revision
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
