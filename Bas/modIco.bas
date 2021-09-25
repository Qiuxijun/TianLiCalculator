Attribute VB_Name = "modIco"
'********************************************************
'**
'**模 块 名：modIco
'**
'**说    明：真彩图标显示模块
'**
'********************************************************
Option Explicit

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PrivateExtractIcons Lib "User32" Alias "PrivateExtractIconsA" (ByVal sFile As String, ByVal nIconIndex As Long, ByVal cxIcon As Long, ByVal cyIcon As Long, phIcon As Long, pIconID As Long, ByVal nIcons As Long, ByVal lFlags As Long) As Long
Private Declare Function DrawIcon Lib "User32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Function AP() As String
    AP = IIf(Len(App.Path) <= 3, App.Path, App.Path & "\")
End Function
                             
Public Function SetFormRGBAIcon(f As Form, ByVal IconSize As Long) As Long
    Dim hIcon As Long
    '加载应用程序在资源管理器中显示的图标，可以修改为其他任意包含图标的 PE 文件路径，或者图标路径
    Call PrivateExtractIcons(AP() & App.EXEName & ".exe", 0, IconSize, IconSize, hIcon, ByVal 0&, 1, 0)
    SetFormRGBAIcon = SendMessage(f.hWnd, WM_SETICON, 0 Or 1, ByVal hIcon)
End Function

Public Sub SetWindowIcon(hWnd As Long, Optional FileName As String, Optional IconIndex As Integer)
    Dim m_Icon As Long
    Dim hModule As Long
    If Len(FileName) = 0 Or Len(Dir(FileName, vbHidden)) = 0 Then
        Dim MyPath As String
        MyPath = App.Path
        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
        FileName = MyPath & App.EXEName & ".exe"
    End If
    hModule = GetModuleHandle(FileName)
    m_Icon = ExtractIcon(hModule, FileName, IconIndex)
    SendMessage hWnd, WM_SETICON, 0, ByVal m_Icon
End Sub
