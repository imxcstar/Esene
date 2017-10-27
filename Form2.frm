VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "串口设置-任务-"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   5715
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   5715
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   30
      Top             =   6480
      Width           =   2175
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HEX模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "文本模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
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
      TabIndex        =   27
      Top             =   4200
      Width           =   2175
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "文本模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HEX模式"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1455
      Left            =   360
      TabIndex        =   22
      Top             =   4920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":000C
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
      TabIndex        =   24
      Top             =   6360
      Width           =   2535
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1920
         Top             =   0
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1560
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   2520
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command4 
      Caption         =   "清空"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   4200
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2175
      Left            =   360
      TabIndex        =   23
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":00A9
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开串口"
      Height          =   300
      Left            =   4200
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   5385
      TabIndex        =   8
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox Combo5 
         Height          =   300
         ItemData        =   "Form2.frx":0146
         Left            =   4560
         List            =   "Form2.frx":0153
         TabIndex        =   14
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         ItemData        =   "Form2.frx":0162
         Left            =   2880
         List            =   "Form2.frx":0175
         TabIndex        =   13
         Text            =   "8"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "Form2.frx":0188
         Left            =   840
         List            =   "Form2.frx":01B3
         TabIndex        =   12
         Text            =   "9600"
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "刷新"
         Height          =   300
         Left            =   4080
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "Form2.frx":020F
         Left            =   840
         List            =   "Form2.frx":0222
         TabIndex        =   9
         Text            =   "无效验"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "串口："
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   165
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "波特率："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "效验位："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据位："
         Height          =   180
         Index           =   3
         Left            =   2160
         TabIndex        =   16
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "停止位："
         Height          =   180
         Index           =   4
         Left            =   3840
         TabIndex        =   15
         Top             =   1080
         Width           =   720
      End
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
      TabIndex        =   0
      Top             =   1560
      Width           =   5415
      Begin VB.CommandButton Command3 
         Caption         =   "发送"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "发送区"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   5175
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "接收区"
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已接收(byte)："
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   5400
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已发送(byte)："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   5640
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1320
         TabIndex        =   5
         Top             =   5400
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   1320
         TabIndex        =   4
         Top             =   5640
         Width           =   90
      End
   End
   Begin VB.Menu C_RW_CD 
      Caption         =   "任务"
      Begin VB.Menu C_XJRW 
         Caption         =   "新建任务"
      End
      Begin VB.Menu C_DKRW 
         Caption         =   "打开任务"
      End
      Begin VB.Menu C_Close_Rw 
         Caption         =   "关闭本任务"
      End
      Begin VB.Menu C_Close_All_Rw 
         Caption         =   "关闭所有任务"
      End
   End
   Begin VB.Menu C_Code_Edit 
      Caption         =   "脚本编辑"
   End
   Begin VB.Menu C_XSQ 
      Caption         =   "显示区"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XYZ(4) As String
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

Private Sub C_Close_All_Rw_Click()
    MDIForm1.C_Close_All_Rw_Click
End Sub

Private Sub C_Close_Rw_Click()
    'Form_Unload 0
    Unload Me
End Sub

Private Sub C_Code_Edit_Click()
    CodeEditC.Show
End Sub

Private Sub C_DKRW_Click()
    MDIForm1.C_DKRW_Click
End Sub

Private Sub C_XJRW_Click()
    MDIForm1.C_XJRW_Click
End Sub

Private Sub C_XSQ_Click()
    XSQ.Show
End Sub

Private Sub Check1_Click()
    Text1.Visible = Check1.Value
    DSFS = Check1.Value
    BCPZ
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    BCPZ
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo3_Click()
    If SFKSXZ = True Then
        BCPZ
    End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo4_Click()
    BCPZ
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo5_Click()
    BCPZ
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    SXComK
End Sub

Private Sub Command2_Click()
    On Error GoTo ERROR
    If Command2.Caption = "打开串口" Then
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
        Text1.Enabled = False
        Command2.Caption = "关闭串口"
        Comm1.CommPort = Val(sMid(LCase(Combo1.Text), "com")) '设置端口
        Comm1.Settings = Combo2.Text & "," & XYZ(Val(Combo3.ListIndex)) & "," & Combo4.Text & "," & Combo5.Text '设置波特率 ,校验位,数据位,停止位
        Comm1.InBufferSize = 1024  '接受缓冲区大小
        Comm1.OutBufferSize = 1024 '发送缓冲区大小
        Comm1.InBufferCount = 0   '清空接受缓冲区
        Comm1.OutBufferCount = 0  '清空发送缓冲区
        Comm1.InputMode = 1 '设置接收数据模式为二进制形式
        Comm1.InputLen = 1 '一次读取1个字节数据
        Comm1.SThreshold = 0 '一次发送所有数据 ,发送数据时不产生OnComm 事件
        Comm1.RThreshold = 1 '每接收1个字节就产生一个OnComm 事件
        Comm1.DTREnable = False
        Comm1.PortOpen = True
        BCPZ
        If DSFS = True Then
            If fsxsms = 1 Then
                DSFSSJ = STByte(RichTextBox2.Text)
            Else
                DSFSSJ2 = RichTextBox2.Text
            End If
            Timer1.Interval = Val(Text1.Text)
            Timer1.Enabled = True
        End If
    Else
ERROR:
        If Err.Number <> 0 Then
            AddLog Now & "——出现一个错误:" & Err.Description
        End If
        Timer1.Enabled = False
        If Comm1.PortOpen = True Then Comm1.PortOpen = False
        Command2.Caption = "打开串口"
        Picture1.BackColor = &H80000005
        Picture2.BackColor = &HC0C0C0
        Picture1.Enabled = True
        Picture2.Enabled = False
        Check1.Enabled = True
        Text1.Enabled = True
    End If
End Sub

Private Sub Command3_Click()
    If fsxsms = 1 Then
        Dim zc() As Byte
        zc = STByte(RichTextBox2.Text)
        fssjbl = fssjbl + UBound(zc)
        If DSFS = True Then DSFSSJ = zc
        Comm1.Output = zc
    Else
        Dim ZC2 As String
        ZC2 = RichTextBox2.Text
        fssjbl = fssjbl + LenB(ZC2)
        If DSFS = True Then DSFSSJ2 = ZC2
        Comm1.Output = ZC2
    End If
    Label3.Caption = fssjbl
End Sub

Private Sub Command4_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub Form_Load()
    XYZ(0) = "N"
    XYZ(1) = "O"
    XYZ(2) = "E"
    XYZ(3) = "M"
    XYZ(4) = "S"
    jsxsms = 1
    fsxsms = 1
    Combo3.ListIndex = 0
    Caption = "串口设置-任务-" & RWName
    C_RW_CD.Caption = "任务-(" & RWName & ")"
    Set CodeEditC = New Form4
    CodeEditC.RWName = RWName
    CodeEditC.RWPath = RWPath
    Set XSQ = New Form1
    XSQ.RWName = RWName
    XSQ.RWPath = RWPath
    Set CodeEditC.XSQ = XSQ
    SXComK
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
        Combo2.Text = zc.Btl
        Combo3.ListIndex = zc.XYW
        Combo4.Text = zc.SJW
        Combo5.Text = zc.TZW
        Text1.Text = zc.dsfstime
        DSFS = zc.sfdsfs
        If DSFS = True Then
            RichTextBox2.Text = zc.FSSJ
            Text1.Visible = True
            Check1.Value = 1
        End If
        Timer1.Interval = zc.dsfstime
        jsxsms = zc.JSMS
        fsxsms = zc.FSMS
        If jsxsms = 0 Then Option3.Value = True
        If fsxsms = 0 Then Option1.Value = True
    End If
    SaveTempFile = FreeFile
    Open RWPath & RWName & "\data.txt" For Append As #SaveTempFile
    SFKSXZ = True
    If Dir(RWPath & RWName & "\r.js") = "" Then
        TempFile = FreeFile
        Open RWPath & RWName & "\r.js" For Output As #TempFile
        Print #TempFile, "var init=function(){" & vbCrLf & vbCrLf & "}" & vbCrLf & vbCrLf & "var run=function(data){" & vbCrLf & vbCrLf & "    return data;" & vbCrLf & "}"
        Close #TempFile
    End If
End Sub

Public Sub Form_Unload(Cancel As Integer)
    If MsgBox("确认关闭任务[" & RWName & "]吗？", vbYesNo) = vbYes Then
        Timer1.Enabled = False
        If Comm1.PortOpen = True Then Comm1.PortOpen = False
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

Private Sub Comm1_OnComm()
    On Error GoTo ERROR
    Dim tempstr() As Byte
    Dim i As Long
    Dim zc As String, ZC2 As String
    
    If Comm1.CommEvent = 2 Then
        If Comm1.InBufferCount <> 0 Then
            tempstr = Comm1.Input
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
        End If
    End If
    Exit Sub
ERROR:
    AddLog Now & "——出现一个错误:" & Err.Description
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

Private Sub SXComK()
    On Error Resume Next
    Dim s As Object
    Combo1.Clear
    For Each s In GetObject("Winmgmts:").InstancesOf("Win32_SerialPortConfiguration")
        If s.IsBusy = False Then
            Combo1.AddItem s.Name
        End If
    Next s
    Combo1.Text = Combo1.List(0)
End Sub

Private Sub BCPZ()
    Dim TempFile As Long
    TempFile = FreeFile
    Dim zc As SJG
    zc.Btl = Combo2.Text
    zc.XYW = Combo3.ListIndex
    zc.SJW = Combo4.Text
    zc.TZW = Combo5.Text
    zc.sfdsfs = DSFS
    zc.dsfstime = Text1.Text
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

Private Sub Timer1_Timer()
    If Comm1.PortOpen = True Then
        If fsxsms = 1 Then
            fssjbl = fssjbl + UBound(DSFSSJ)
            Comm1.Output = DSFSSJ
        Else
            fssjbl = fssjbl + LenB(DSFSSJ2)
            Comm1.Output = DSFSSJ2
        End If
        Label3.Caption = fssjbl
    End If
End Sub
