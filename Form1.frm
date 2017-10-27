VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "显示区-任务-"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   300
   ClientWidth     =   11550
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   770
   Begin Esene.CurveGraph CurveGraph 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9135
      _extentx        =   16113
      _extenty        =   9551
      legendfont      =   "Form1.frx":000C
      axesfont        =   "Form1.frx":0030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDc As Long, ByVal nXStart As Long, ByVal nYStart As Long, ByVal lpString As String, ByVal cchString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public RWName As String, RWPath As String

Private Sub Form_Load()
    Me.Caption = "显示区-任务-" & RWName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub

Public Function printf(s As String, Optional x As Long = 0, Optional y As Long = 0)
    On Error GoTo Err
    Dim l As Long
    l = lstrlen(s)
    printf = TextOut(Me.hDc, x, y, s, l)
    Exit Function
Err:
    AddLog Now & "——出现一个错误:" & Err.Description
End Function

Public Sub line1(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional color As Long = vbBlack)
    On Error GoTo Err
    Me.Line (x1, y1)-(x2, y2), color
    Exit Sub
Err:
    AddLog Now & "——出现一个错误:" & Err.Description
End Sub

Public Sub line2(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional color As Long = vbBlack)
    On Error GoTo Err
    Me.Line (x1, y1)-(x2, y2), color, B
    Exit Sub
Err:
    AddLog Now & "——出现一个错误:" & Err.Description
End Sub

Public Sub line3(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional color As Long = vbBlack)
    On Error GoTo Err
    Me.Line (x1, y1)-(x2, y2), color, BF
    Exit Sub
Err:
    AddLog Now & "——出现一个错误:" & Err.Description
End Sub

Public Sub dpoint(x As Long, y As Long, Optional color As Long = vbBlack)
    On Error GoTo Err
    Me.PSet (x, y), color
    Exit Sub
Err:
    AddLog Now & "——出现一个错误:" & Err.Description
End Sub

Public Sub layout(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    On Error GoTo Err
    Me.Scale (x1, y1)-(x2, y2)
    Exit Sub
Err:
    AddLog Now & "——出现一个错误:" & Err.Description
End Sub

