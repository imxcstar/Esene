VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "网络设置-任务-"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   5655
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   5655
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1455
      Left            =   360
      TabIndex        =   12
      Top             =   4440
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2566
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form5.frx":000C
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1920
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清空"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HEX模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "文本模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   17
      Top             =   5880
      Width           =   2535
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1920
         Top             =   0
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1560
         TabIndex        =   18
         Text            =   "300"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "定时发送(毫秒)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "文本模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HEX模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2175
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3836
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form5.frx":00A9
   End
   Begin VB.CommandButton Command1 
      Caption         =   "监听"
      Height          =   420
      Left            =   4200
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5865
      ScaleWidth      =   5385
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "接收区"
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5175
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "发送区"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   5175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "发送"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1320
         TabIndex        =   9
         Top             =   5640
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1320
         TabIndex        =   8
         Top             =   5400
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已发送(byte)："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已接收(byte)："
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   5400
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   10
         Text            =   "12566"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "无状态"
         Height          =   180
         Left            =   960
         TabIndex        =   25
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接状态："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   150
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "监听端口："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   510
         Width           =   900
      End
   End
   Begin VB.Menu C_RW 
      Caption         =   "任务"
      Begin VB.Menu C_XJRW 
         Caption         =   "新建任务"
      End
      Begin VB.Menu C_DKRW 
         Caption         =   "打开任务"
      End
      Begin VB.Menu C_GBRW 
         Caption         =   "关闭本任务"
      End
      Begin VB.Menu C_GBSYRW 
         Caption         =   "关闭所有任务"
      End
   End
   Begin VB.Menu C_JBBJ 
      Caption         =   "脚本编辑"
   End
   Begin VB.Menu C_XSQ 
      Caption         =   "显示区"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jssjbl As Long, fssjbl As Long
Dim jsxsms As Integer, fsxsms As Integer
Public RWName As String, RWPath As String
Public SFXJ As Boolean
Public RWTYPE As Integer
Dim CodeEditC As Form4
Dim XSQ As Form1
Dim SFKSXZ As Boolean
Dim SaveTempFile As Long
Dim JsRun As Object
Dim DSFS As Boolean
Dim DSFSSJ() As Byte
Dim DSFSSJ2 As String
Dim API As New APIClass
Dim StateZ() As String

Private Sub C_DKRW_Click()
    MDIForm1.C_DKRW_Click
End Sub

Private Sub C_GBRW_Click()
    Unload Me
End Sub

Private Sub C_GBSYRW_Click()
    MDIForm1.C_Close_All_Rw_Click
End Sub

Private Sub C_JBBJ_Click()
    CodeEditC.Show
End Sub

Private Sub C_XJRW_Click()
    MDIForm1.C_XJRW_Click
End Sub

Private Sub C_XSQ_Click()
    XSQ.Show
End Sub

Private Sub Check1_Click()
    Text2.Visible = Check1.Value
    DSFS = Check1.Value
    BCPZ
End Sub

Private Sub Command1_Click()
    On Error GoTo ERROR
    If Command1.Caption = "监听" Then
        Timer2.Enabled = True
        Set JsRun = Nothing
        Set JsRun = CreateObject("MSScriptControl.ScriptControl")
        JsRun.AllowUI = True
        JsRun.Language = "JavaScript"
        JsRun.AddObject "Me", XSQ
        JsRun.AddObject "API", API
        JsRun.AddCode CodeEditC.ToText
        JsRun.Eval "init();"
        Picture1.BackColor = &HC0C0C0
        Picture2.BackColor = &H80000005
        Picture1.Enabled = False
        Picture2.Enabled = True
        Check1.Enabled = False
        Text2.Enabled = False
        Command1.Caption = "断开"
        Winsock1.LocalPort = Val(Text1.Text)
        Winsock1.Listen
        BCPZ
        If DSFS = True Then
            If fsxsms = 1 Then
                DSFSSJ = STByte(RichTextBox2.Text)
            Else
                DSFSSJ2 = RichTextBox2.Text
            End If
            Timer1.Interval = Val(Text2.Text)
            Timer1.Enabled = True
        End If
    Else
ERROR:
        If Err.Number <> 0 Then
            AddLog Now & "——出现一个错误:" & Err.Description
        End If
        Timer2.Enabled = False
        Timer1.Enabled = False
        Winsock1.Close
        Command1.Caption = "监听"
        Label4.Caption = "无状态"
        Picture1.BackColor = &H80000005
        Picture2.BackColor = &HC0C0C0
        Picture1.Enabled = True
        Picture2.Enabled = False
        Check1.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
    If Winsock1.State = 7 Then
        If fsxsms = 1 Then
            Dim zc() As Byte
            zc = STByte(RichTextBox2.Text)
            fssjbl = fssjbl + UBound(zc)
            If DSFS = True Then DSFSSJ = zc
            Winsock1.SendData zc
        Else
            Dim ZC2 As String
            ZC2 = RichTextBox2.Text
            fssjbl = fssjbl + LenB(ZC2)
            If DSFS = True Then DSFSSJ2 = ZC2
            Winsock1.SendData ZC2
        End If
        Label3.Caption = fssjbl
    End If
End Sub

Private Sub Command4_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub Form_Load()
    StateZ = Split("关闭|打开|侦听|挂起|解析域名|已识别主机|正在连接|已连接|同级人员正在关闭连接|错误", "|")
    jsxsms = 1
    fsxsms = 1
    Caption = "网络设置-任务-" & RWName
    C_RW.Caption = "任务-(" & RWName & ")"
    Set CodeEditC = New Form4
    CodeEditC.RWName = RWName
    CodeEditC.RWPath = RWPath
    Set XSQ = New Form1
    XSQ.RWName = RWName
    XSQ.RWPath = RWPath
    Set CodeEditC.XSQ = XSQ
    Dim TempFile As Long
    If SFXJ = True Then
        BCPZ
    Else
        TempFile = FreeFile
        Dim rw As RWJG
        Dim zc As SJG
        Open RWPath & RWName & "\c.s" For Binary As #TempFile
        Get #TempFile, , rw
        Close #TempFile
        zc = rw.RWSJ
        Text1.Text = zc.jtdk
        Text2.Text = zc.dsfstime
        DSFS = zc.sfdsfs
        If DSFS = True Then
            RichTextBox2.Text = zc.FSSJ
            Text2.Visible = True
            Check1.Value = 1
        End If
        jsxsms = zc.JSMS
        fsxsms = zc.FSMS
        If jsxsms = 0 Then Option3.Value = True
        If fsxsms = 0 Then Option1.Value = True
    End If
    SaveTempFile = FreeFile
    Open RWPath & RWName & "\data.txt" For Append As #SaveTempFile
    If Dir(RWPath & RWName & "\r.js") = "" Then
        TempFile = FreeFile
        Open RWPath & RWName & "\r.js" For Output As #TempFile
        Print #TempFile, "var init=function(){" & vbCrLf & vbCrLf & "}" & vbCrLf & vbCrLf & "var run=function(data){" & vbCrLf & vbCrLf & "    return data;" & vbCrLf & "}"
        Close #TempFile
    End If
End Sub

Private Sub BCPZ()
    Dim TempFile As Long
    TempFile = FreeFile
    Dim zc As SJG
    zc.jtdk = Val(Text1.Text)
    zc.sfdsfs = DSFS
    zc.dsfstime = Text2.Text
    zc.JSMS = jsxsms
    zc.FSMS = fsxsms
    If DSFS = True Then
        zc.FSSJ = RichTextBox2.Text
    End If
    Dim rw As RWJG
    rw.RWSJ = zc
    rw.RWBB = App.Major & "." & App.Minor & "." & App.Revision
    rw.RWTYPE = RWTYPE
    Open RWPath & RWName & "\c.s" For Binary As #TempFile
    Put #TempFile, , rw
    Close #TempFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("确认关闭任务[" & RWName & "]吗？", vbYesNo) = vbYes Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        If Winsock1.State = 7 Then Winsock1.Close
        BCPZ
        CodeEditC.SaveCode
        Close #SaveTempFile
        Unload CodeEditC
        Unload XSQ
        Unload Me
    Else
        Cancel = 1
    End If
End Sub

Private Sub Option1_Click()
    fsxsms = 0
    RichTextBox2_Change
    BCPZ
End Sub

Private Sub Option2_Click()
    fsxsms = 1
    RichTextBox2_Change
    BCPZ
End Sub

Private Sub Option3_Click()
    jsxsms = 0
    BCPZ
End Sub

Private Sub Option4_Click()
    jsxsms = 1
    BCPZ
End Sub

Private Sub RichTextBox2_Change()
    If DSFS = True Then
        If fsxsms = 1 Then
            DSFSSJ = STByte(RichTextBox2.Text)
        Else
            DSFSSJ2 = RichTextBox2.Text
        End If
    End If
    BCPZ
End Sub

Private Sub Text1_Change()
    BCPZ
End Sub

Private Sub Text2_Change()
    BCPZ
End Sub

Private Sub Timer1_Timer()
    If Winsock1.State = 7 Then
        If fsxsms = 1 Then
            fssjbl = fssjbl + UBound(DSFSSJ)
            Winsock1.SendData DSFSSJ
        Else
            fssjbl = fssjbl + LenB(DSFSSJ2)
            Winsock1.SendData DSFSSJ2
        End If
        Label3.Caption = fssjbl
    End If
End Sub

Private Sub Timer2_Timer()
    Label4.Caption = StateZ(Winsock1.State)
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ERROR
    Dim tempstr() As Byte
    Dim i As Long
    Dim zc As String, ZC2 As String
    

    Winsock1.GetData tempstr
    jssjbl = jssjbl + UBound(tempstr)
    Label2.Caption = jssjbl
    If jsxsms = 1 Then
        For i = LBound(tempstr) To UBound(tempstr)
            zc = Hex(tempstr(i))
            If Len(zc) = 1 Then zc = "0" & zc
            ZC2 = ZC2 & zc & " "
        Next i
        If Len(RichTextBox1.Text) >= 20000 Then RichTextBox1.Text = ""
        RichTextBox1.SelStart = Len(RichTextBox1.Text)
        zc = JsRun.Eval("run(""" & ZC2 & """);")
        If zc <> "" Then
            Print #SaveTempFile, zc
        End If
        RichTextBox1.SelText = ZC2
    Else
        If Len(RichTextBox1.Text) >= 20000 Then RichTextBox1.Text = ""
        RichTextBox1.SelStart = Len(RichTextBox1.Text)
        zc = StrConv(tempstr, vbUnicode)
        ZC2 = JsRun.Eval("run(""" & Replace(Replace(Replace(zc, Chr(0), ""), Chr(13), ""), Chr(10), "") & """);")
        If ZC2 <> "" Then
            Print #SaveTempFile, ZC2
        End If
        RichTextBox1.SelText = zc
    End If
    Exit Sub
ERROR:
    AddLog Now & "——出现一个错误:" & Err.Description
End Sub

