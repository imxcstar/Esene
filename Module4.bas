Attribute VB_Name = "Module4"
Option Explicit
 
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
 
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
 
' 打开 Windows 的选择目录对话框
' hwnd 为窗口句柄(通常设为 Me.hwnd), Prompt 为指示字符串
Public Function ShowFolderSelection(ByVal hwnd As Long, ByVal Prompt As String) As String
 
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
 
    With udtBI
        .hWndOwner = hwnd
        .lpszTitle = lstrcat(Prompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
 
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
    End If
 
    ShowFolderSelection = sPath
 
End Function
