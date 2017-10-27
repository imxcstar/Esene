Attribute VB_Name = "Module5"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PrivateExtractIcons Lib "user32" Alias "PrivateExtractIconsA" (ByVal sFile As String, ByVal nIconIndex As Long, ByVal cxIcon As Long, ByVal cyIcon As Long, phIcon As Long, pIconID As Long, ByVal nIcons As Long, ByVal lFlags As Long) As Long
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Public hIcon2 As Long

Private Function AP() As String
    AP = IIf(Len(App.Path) <= 3, App.Path, App.Path & "\")
End Function

Public Function SetFormRGBAIcon(F As Form, ByVal IconSize As Long) As Long
    Dim hIcon As Long
    '加载应用程序在资源管理器中显示的图标，可以修改为其他任意包含图标的 PE 文件路径，或者图标路径
    Call PrivateExtractIcons(AP() & App.EXEName & ".exe", 0, IconSize, IconSize, hIcon, ByVal 0&, 1, 0)
    hIcon2 = hIcon
    SetFormRGBAIcon = SendMessage(F.hWnd, WM_SETICON, 0 Or 1, ByVal hIcon)
End Function
