VERSION 5.00
Begin VB.Form Form6 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "新建任务"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5880
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新建"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   270
      Left            =   5040
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   4575
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "串口接收"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "网络接收"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "任务名（任务名不能包括以下特殊字符：/\:*?""""<>|）："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保存路径："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "任务类型："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ZCML As String

Private Sub Command1_Click()
    ZCML = ShowFolderSelection(Me.hwnd, "请选择新建任务的保存位置")
    If ZCML <> "" Then
        If Right(ZCML, 1) <> "\" Then ZCML = ZCML & "\"
        Text1.Text = ZCML
    End If
End Sub

Private Sub Command2_Click()
    Dim zc As String
    zc = Text2.Text
    If zc <> "" Then
        If Dir(ZCML & zc, vbDirectory) <> "" Then
            MsgBox "保存位置已存在和任务名同样的文件夹或文件！请修改任务名并重试！"
        Else
            MkDir ZCML & zc
            MkDir ZCML & zc & "\backup"
            If Option1.Value = True Then
                Dim ZC4 As New Form5
                ReDim Preserve Form5s(UBound(Form5s) + 1)
                Set Form5s(UBound(Form5s)) = ZC4
                Form5s(UBound(Form5s)).RWPath = ZCML
                Form5s(UBound(Form5s)).RWName = zc
                Form5s(UBound(Form5s)).RWTYPE = 0
                Form5s(UBound(Form5s)).SFXJ = True
                Form5s(UBound(Form5s)).Show
            End If
            If Option2.Value = True Then
                Dim ZC3 As New Form2
                ReDim Preserve Form2s(UBound(Form2s) + 1)
                Set Form2s(UBound(Form2s)) = ZC3
                Form2s(UBound(Form2s)).RWPath = ZCML
                Form2s(UBound(Form2s)).RWName = zc
                Form2s(UBound(Form2s)).RWTYPE = 1
                Form2s(UBound(Form2s)).SFXJ = True
                Form2s(UBound(Form2s)).Show
            End If
            Unload Me
        End If
    Else
        MsgBox "请输入任务名"
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
