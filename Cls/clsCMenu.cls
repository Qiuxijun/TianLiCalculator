VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DT_LEFT           As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_SINGLELINE     As Long = &H20

Private Const WM_SIZE           As Long = &H5
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_MBUTTONDOWN    As Long = &H207
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_NCLBUTTONDOWN  As Long = &HA1
Private Const WM_NCMBUTTONDOWN  As Long = &HA7
Private Const WM_NCRBUTTONDOWN  As Long = &HA4
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_MBUTTONUP      As Long = &H208
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_ERASEBKGND     As Long = &H14
Private Const WM_PAINT          As Long = &HF
Private Const WM_MOUSEACTIVATE  As Long = &H21
Private Const WM_NCACTIVATE     As Long = &H86
Private Const WM_DESTROY        As Long = &H2
Private Const WM_SHOWWINDOW     As Long = &H18
Private Const WM_ALLMSG         As Long = -1
Private Const WM_GETFONT        As Long = &H31
Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_CHAR           As Long = &H102
Private Const WM_KEYUP          As Long = &H101

Private Const MK_LBUTTON        As Long = &H1
Private Const MK_MBUTTON        As Long = &H10
Private Const MK_RBUTTON        As Long = &H2

Private Const TRANSPARENT       As Long = 1

Private Const SW_HIDE           As Long = 0
Private Const SW_SHOWNORMAL     As Long = 1

Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1

'窗口类名
Private Const SZ_CLASSNAME      As String = "#32770"
'菜单项的高度
Private Const ITM_HEIGHT        As Long = 24
'分割条高度
Private Const ITM_SEPHEIGHT     As Long = 3
'边距大小(Item离边距的大小)
Private Const ITM_MARGIN        As Long = 2
'用于显示图标的左侧栏宽度
Private Const ITM_ICONBAR       As Long = 25

'窗口风格
Private Const WS_CHILD          As Long = &H40000000
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const WS_EX_TOPMOST     As Long = &H8&
'==============================================ICO
Private Declare Function CreatePolygonRgn Lib "gdi32 " (lpPoint As POINTAPI, _
                          ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32 " (ByVal hDc As Long, _
                          ByVal hRgn As Long, ByVal hBrush As Long) As Long
Const ALTERNATE = 1
'==========================================
Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Private Type POINTAPI
    x                           As Long
    y                           As Long
End Type

Private Type SIZE
    cX                          As Long
    cY                          As Long
End Type

Private Type PAINTSTRUCT
    hDc                         As Long
    fErase                      As Long
    rcPaint                     As RECT
    fRestore                    As Long
    fIncUpdate                  As Long
    rgbReserved(32)             As Byte
End Type

Private Type MENUITEMDATA
    Text                        As String
    Key                         As String
    MenuType                    As MENUTYPECONSTANT
    lAcckey                     As Long
    Image                       As Long
    Enabled                     As Boolean
    Visible                     As Boolean
    Checked                     As Boolean
    SubMenu                     As clsCMenu '子菜单(不搞了，可以自己搞搞)
    rc                          As RECT
End Type

Public Enum MENUTYPECONSTANT
    MenuString = 0
    MenuSeparate = 1
    CheckBox = 2
End Enum

Private Declare Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hDc As Long) As Long

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "User32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function BeginPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "User32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetFocus Lib "User32" () As Long

Private WithEvents c_Subclass   As clsCSubclass
Attribute c_Subclass.VB_VarHelpID = -1

Private m_Menu()                As MENUITEMDATA '菜单项
Private m_Count                 As Long         '菜单项个数
Private m_VCount                As Long         '可见的个数

Private m_hWnd                  As Long         '窗口句柄
Private m_hOwner                As Long         '菜单父窗口句柄
Private m_hFocus                As Long         '弹出菜单时活动的窗口句柄
Private m_hBGDC                 As Long         '背景MemDC
Private m_hDC                   As Long         '缓冲MemDC

Private m_HoverIdx              As Long         '鼠标移过的项
Private m_cX                    As Long         '菜单窗口宽度
Private m_cY                    As Long         '菜单窗口高度

Public Event MenuClick(ByVal MenuIndex As Long)
Public Event MenuDrawbackground(ByVal hDc As Long, ByVal x As Long, _
              ByVal y As Long, ByVal cX As Long, ByVal cY As Long, SkipDef As Boolean)

Public Event MenuDrawItem(ByVal MenuIndex As Long, ByVal hDc As Long, ByVal x As Long, _
              ByVal y As Long, ByVal cX As Long, ByVal cY As Long, SkipDef As Boolean)

Public Function Create(ByVal hOwner As Long) As Long
    '如果已经创建窗口，则不再创建(如果要再创建请先销毁窗口)
    If m_hWnd Then Exit Function
    'WS_EX_TOOLWINDOW可以不在任务栏上显示
    'WS_EX_TOPMOST总在顶层显示
    'WS_CHILD不抢父窗口焦点
    '创建xy坐标那么大负数，是为了父窗口销毁的时候，菜单窗口不会在屏幕闪一下才销毁
    m_hOwner = hOwner
    m_hWnd = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, _
             SZ_CLASSNAME, _
             vbNullString, _
             WS_CHILD, -1000, -1000, 1, 1, _
             hOwner, 0, 0, 0)
    '设置为屏幕的子窗口
    If m_hWnd Then
        Call SetParent(m_hWnd, 0)
        Call OnSize
        With c_Subclass
            .ssc_Subclass m_hWnd
            .ssc_AddMsg m_hWnd, WM_PAINT, MSG_BEFORE
            .ssc_AddMsg m_hWnd, WM_ERASEBKGND, MSG_BEFORE
            .ssc_AddMsg m_hWnd, WM_SIZE, MSG_AFTER
            .ssc_AddMsg m_hWnd, WM_SHOWWINDOW, MSG_AFTER
            .ssc_AddMsg m_hWnd, WM_MOUSEMOVE, MSG_AFTER
            .ssc_AddMsg m_hWnd, WM_MOUSELEAVE, MSG_AFTER
            .ssc_AddMsg m_hWnd, WM_LBUTTONDOWN, MSG_AFTER
            .ssc_AddMsg m_hWnd, WM_LBUTTONUP, MSG_AFTER
            .ssc_AddMsg m_hWnd, WM_KEYDOWN, MSG_AFTER
        End With
    End If
    Create = m_hWnd
End Function

Public Function Add(Optional ByVal Text As Variant, _
                    Optional ByVal Key As Variant, _
                    Optional ByVal Index As Variant, _
                    Optional ByVal MenuType As MENUTYPECONSTANT = MenuString) As Long
    Dim lIdx    As Long
    If IsMissing(Index) Then
        lIdx = m_Count
    Else
        lIdx = CLng(Index)
        If lIdx < 0 Or lIdx > m_Count Then Exit Function
    End If
    '不允许出现相同的Key(注意Key大小写敏感)
    If Not IsMissing(Key) Then
        If GetMenuIndex(CStr(Key)) <> -1 Then Exit Function
    End If
    ReDim Preserve m_Menu(m_Count)
    If lIdx < m_Count Then
        '如果不是增加到最后，则把前面的项向后移动
        Dim i As Long
        For i = m_Count To lIdx + 1 Step -1
            LSet m_Menu(i) = m_Menu(i - 1)
        Next
    End If
    With m_Menu(lIdx)
        '如果参数没有text则默认为添加分隔条
        If IsMissing(Text) Then
            .MenuType = MenuSeparate
        Else
            .Text = CStr(Text)
            .MenuType = MenuType
        End If
        If Not IsMissing(Key) Then .Key = CStr(Key)
        .Enabled = True
        .Visible = True
        .Checked = True
        .lAcckey = GetAccessKey(.Text)
    End With
    m_Count = m_Count + 1
    Call CalcSize
    Add = 1
End Function

Public Function Remove(ByVal Index As Variant) As Long
    '移除一个菜单项
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Function
    If lIdx < m_Count - 1 Then
        '如果移除的不是最后一个，则把后面的项向前移动
        Dim i As Long
        For i = lIdx To m_Count - 2
            LSet m_Menu(i) = m_Menu(i + 1)
        Next
    End If
    m_Count = m_Count - 1
    ReDim Preserve m_Menu(m_Count - 1)
    Call CalcSize
    Remove = lIdx
End Function

Private Function Clear() As Long
    '清除所有菜单项
    If m_Count = 0 Then Exit Function
    Erase m_Menu()
    m_Count = 0
    Call CalcSize
End Function

Public Function Show(Optional ByVal x As Variant, Optional ByVal y As Variant) As Long
    '显示菜单
    If m_hWnd = 0 Then Exit Function
    If m_VCount = 0 Then Exit Function
    Dim pt      As POINTAPI
    Dim ptShow  As POINTAPI
    Dim sW As Long
    Dim sH  As Long
    Call GetCursorPos(pt)
    sW = Screen.Width / Screen.TwipsPerPixelX
    sH = Screen.Height / Screen.TwipsPerPixelY
    If IsMissing(x) Then
        ptShow.x = pt.x
        If pt.x > sW - m_cX Then ptShow.x = pt.x - m_cX
    Else
        ptShow.x = CLng(Val(x))
        If ptShow.x < sW - m_cX Then ptShow.x = sW - m_cX
        If ptShow.x < 0 Then ptShow.x = 0
    End If
    If IsMissing(y) Then
        ptShow.y = pt.y
        If pt.y > sH - m_cY Then ptShow.y = pt.y - m_cY
    Else
        ptShow.y = CLng(Val(y))
        If ptShow.y > sH - m_cY Then ptShow.y = sH - m_cY
        If ptShow.y < 0 Then ptShow.y = 0
    End If
    Call SetWindowPos(m_hWnd, -1, ptShow.x, ptShow.y, 0, 0, SWP_NOSIZE)
    Show = ShowWindow(m_hWnd, SW_SHOWNORMAL)
End Function

Public Function Hide() As Long
    If m_hWnd = 0 Then Exit Function
    Hide = ShowWindow(m_hWnd, SW_HIDE)
End Function

Public Function HIWORD(ByVal Param As Long) As Long
    HIWORD = ((Param And &H7FFF0000) \ &H10000) Or (((Param And &H80000000) <> 0) And &H8000)
End Function

Public Function LOWORD(ByVal Param As Long) As Long
    LOWORD = (Param And &H7FFF&) Or (((Param And &H8000&) <> 0) And &H8000)
End Function

Private Function RedrawBackground()
    Dim bSkip As Boolean
    RaiseEvent MenuDrawbackground(m_hBGDC, 0, 0, m_cX, m_cY, bSkip)
    If bSkip = False Then
        Call PaintColor(m_hBGDC, 0, 0, m_cX, m_cY, &HB3B3B3) '边框颜色
        Call PaintColor(m_hBGDC, 1, 1, m_cX - 2, m_cY - 2, &HFFFFFF) '背景颜色
    End If
End Function

Private Function RedrawAll(Optional ByVal bRefresh As Boolean)
    Call BitBlt(m_hDC, 0, 0, m_cX, m_cY, m_hBGDC, 0, 0, vbSrcCopy)
    If m_VCount Then
        Dim i As Long
        For i = 0 To m_Count - 1
            Dim lx As Long
            With m_Menu(i)
                lx = ITM_MARGIN + ITM_ICONBAR + 2
                If .MenuType = MenuString Or .MenuType = CheckBox Then
                    Call OnDrawItem(i, ITM_MARGIN, .rc.Top, m_cX - ITM_MARGIN * 2, ITM_HEIGHT)
                Else
                    Call OnDrawItem(i, ITM_MARGIN, .rc.Top, m_cX - ITM_MARGIN * 2, ITM_SEPHEIGHT)
                End If
            End With
        Next
    End If
    If bRefresh Then PaintToWnd
End Function

Private Sub RedrawOne(ByVal Index As Long, Optional ByVal bRefresh As Boolean)
    If Index < 0 Or Index > m_Count - 1 Then Exit Sub
    If m_Menu(Index).Visible = False Then Exit Sub
    Dim lx  As Long
    With m_Menu(Index)
        lx = ITM_MARGIN + ITM_ICONBAR + 2
        If .MenuType = MenuString Or .MenuType = CheckBox Then
            Call OnDrawItem(Index, ITM_MARGIN, .rc.Top, m_cX - ITM_MARGIN * 2, ITM_HEIGHT)
        Else
            Call OnDrawItem(Index, ITM_MARGIN, .rc.Top, m_cX - ITM_MARGIN * 2, ITM_SEPHEIGHT)
        End If
    End With
    If bRefresh Then PaintToWnd
End Sub

Private Sub OnDrawItem(ByVal Index As Long, ByVal x As Long, ByVal y As Long, _
                       ByVal cX As Long, ByVal cY As Long)
    If Index < 0 Or Index > m_Count - 1 Then Exit Sub
    If m_Menu(Index).Visible = False Then Exit Sub
    Dim bSkip As Boolean
    '先拷贝背景的内容过来
    Call BitBlt(m_hDC, x, y, cX, cY, m_hBGDC, x, y, vbSrcCopy)
    RaiseEvent MenuDrawItem(Index, m_hDC, x, y, cX, cY, bSkip)
    If bSkip Then Exit Sub  '如果禁止默认重绘
    If m_Menu(Index).MenuType = MenuString Or m_Menu(Index).MenuType = CheckBox Then
        Call SetTextColor(m_hDC, IIf(m_Menu(Index).Enabled, &H4F4F4F, &HB4ADA1)) '文本颜色 与禁用项颜色

        If Index = m_HoverIdx Then
            Call PaintColor(m_hDC, x, y, cX, cY, &HBDBBB7) '指向边框颜色
            Call PaintColor(m_hDC, x + 1, y + 1, cX - 2, cY - 2, &HEFF1F2) '指向选择颜色
            'Call PaintColor(m_hDC, X + 2, Y + 2, cX - 4, cY - 4, &HEDBD00) '指向颜色
            Call PaintICO(m_hDC, x, y, Index)
        Else
            Call BitBlt(m_hDC, x, y, cX, cY, m_hBGDC, x, y, vbSrcCopy)
            Call PaintICO(m_hDC, x, y, Index)
        End If

        Call PaintText(m_hDC, m_Menu(Index).Text, x + ITM_ICONBAR + 2, y, cX - ITM_ICONBAR - 2, cY, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE)
    Else
        Call PaintColor(m_hDC, x, y + (cY - 1) / 2, cX, 1, &HD2D2D2) '分隔条颜色
        'FBF3E8
        Call PaintColor(m_hDC, x, y + (cY - 1) / 2 + 1, cX, 1, &HFCF5EC) '分隔条颜色
    End If
End Sub

Private Sub PaintICO(ByVal hDc As Long, ByVal x As Long, _
                     ByVal y As Long, ByVal Index As Long)
    If m_Menu(Index).Checked And m_Menu(Index).MenuType = CheckBox Then
        Dim Arr(27) As POINTAPI
        Dim lb     As Long
        Dim lOld     As Long
        Dim hRgn     As Long

        Arr(0).x = 0
        Arr(0).y = 2
        Arr(1).x = 0
        Arr(1).y = 5
        Arr(2).x = 1
        Arr(2).y = 5
        Arr(3).x = 1
        Arr(3).y = 6
        Arr(4).x = 2
        Arr(4).y = 6
        Arr(5).x = 2
        Arr(5).y = 7
        Arr(6).x = 3
        Arr(6).y = 7
        Arr(7).x = 3
        Arr(7).y = 6
        Arr(8).x = 4
        Arr(8).y = 6
        Arr(9).x = 4
        Arr(9).y = 5
        Arr(10).x = 5
        Arr(10).y = 5
        Arr(11).x = 5
        Arr(11).y = 4
        Arr(12).x = 6
        Arr(12).y = 4
        Arr(13).x = 6
        Arr(13).y = 3
        Arr(14).x = 7
        Arr(14).y = 3
        Arr(15).x = 7
        Arr(15).y = 0
        Arr(16).x = 6
        Arr(16).y = 0
        Arr(17).x = 6
        Arr(17).y = 1
        Arr(18).x = 5
        Arr(18).y = 1
        Arr(19).x = 5
        Arr(19).y = 2
        Arr(20).x = 4
        Arr(20).y = 2
        Arr(21).x = 4
        Arr(21).y = 3
        Arr(22).x = 3
        Arr(22).y = 3
        Arr(23).x = 3
        Arr(23).y = 4
        Arr(24).x = 2
        Arr(24).y = 4
        Arr(25).x = 2
        Arr(25).y = 3
        Arr(26).x = 1
        Arr(26).y = 3
        Arr(27).x = 1
        Arr(27).y = 2
        Dim i As Integer
        For i = 0 To UBound(Arr)
            Arr(i).x = Arr(i).x + x + 8
            Arr(i).y = Arr(i).y + y + 8
        Next
        hRgn = CreatePolygonRgn(Arr(0), UBound(Arr) + 1, ALTERNATE)
        lb = CreateSolidBrush(RGB(0, 0, 0))
        Call FillRgn(m_hDC, hRgn, lb)
        Call DeleteObject(lb)
        Call DeleteObject(hRgn)
    End If
End Sub

Private Sub PaintToWnd()
    Dim wDC As Long
    wDC = GetDC(m_hWnd)
    Call BitBlt(wDC, 0, 0, m_cX, m_cY, m_hDC, 0, 0, vbSrcCopy)
    Call ReleaseDC(m_hWnd, wDC)
End Sub

Private Function GetMenuIndex(ByVal Index As Variant) As Long
    If m_Count = 0 Then GetMenuIndex = -1: Exit Function
    Dim lIdx As Long
    lIdx = -1
    If VarType(Index) = vbString Then
        Dim i As Long
        Dim szkey As String
        szkey = CStr(Index)
        If Len(szkey) = 0 Then GetMenuIndex = -1: Exit Function
        For i = 0 To m_Count - 1
            If m_Menu(i).Key = szkey Then
                lIdx = i
                Exit For
            End If
        Next
    Else
        lIdx = CLng(Val(Index))
        If lIdx < 0 Then lIdx = -1
        If lIdx > m_Count - 1 Then lIdx = -1
    End If
    GetMenuIndex = lIdx
End Function

Private Sub CalcSize()
    '计算菜单的高度、宽度、以及项的RECT
    m_VCount = 0
    If m_Count = 0 Then
        m_cX = ITM_ICONBAR + ITM_MARGIN * 2
        m_cY = ITM_HEIGHT + ITM_MARGIN * 2
    Else
        Dim i       As Long
        Dim lMax    As Long
        Dim lcX     As Long
        Dim lcY     As Long
        For i = 0 To m_Count - 1
            If m_Menu(i).Visible Then
                m_VCount = m_VCount + 1
                With m_Menu(i)
                    '得到文字最宽的项目
                    lMax = GetTextWidth(m_hDC, m_Menu(i).Text)
                    If lMax > lcX Then lcX = lMax
                    '得到所有项高度
                    .rc.Top = ITM_MARGIN + lcY
                    .rc.Bottom = ITM_MARGIN + lcY + IIf(.MenuType = MenuString Or .MenuType = 2, ITM_HEIGHT, ITM_SEPHEIGHT)
                    lcY = lcY + IIf(.MenuType = MenuString Or .MenuType = CheckBox, ITM_HEIGHT, ITM_SEPHEIGHT)
                End With
            End If
        Next
        m_cX = lcX + ITM_ICONBAR + ITM_MARGIN * 2 + 32  '这个32是预留的宽度
        m_cY = lcY + ITM_MARGIN * 2
        If m_cX > Screen.Width / Screen.TwipsPerPixelX Then m_cX = Screen.Width / Screen.TwipsPerPixelX
        If m_cY > Screen.Height / Screen.TwipsPerPixelY Then m_cY = Screen.Height / Screen.TwipsPerPixelY
        If m_cX < ITM_ICONBAR + ITM_MARGIN * 2 Then m_cX = ITM_ICONBAR + ITM_MARGIN * 2
        If m_cY < ITM_HEIGHT + ITM_MARGIN * 2 Then m_cY = ITM_HEIGHT + ITM_MARGIN * 2
    End If
    If m_hWnd Then Call SetWindowPos(m_hWnd, 0, 0, 0, m_cX, m_cY, SWP_NOMOVE)
End Sub

Private Function GetAccessKey(ByVal szCaption As String) As Long
    '得到AccessKey
    If Len(szCaption) = 0 Then Exit Function
    Dim lPos  As Long
    lPos = InStrRev(szCaption, "&")
    Do While lPos
        If lPos <> 0 And lPos <= Len(szCaption) - 1 Then
            GetAccessKey = Asc(UCase$(Mid$(szCaption, lPos + 1, 1)))
            Exit Function
        End If
        If lPos <= 1 Then Exit Function
        lPos = InStrRev(szCaption, "&", lPos - 1)
    Loop
End Function

Private Function CreateMemDC(ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpDC   As Long
    Dim rDC     As Long
    Dim rBmp    As Long
    TmpDC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If TmpDC Then
        rDC = CreateCompatibleDC(TmpDC)
        If rDC Then
            rBmp = CreateCompatibleBitmap(TmpDC, Width, Height)
            If rBmp Then
                DeleteObject SelectObject(rDC, rBmp)
                CreateMemDC = rDC
                DeleteObject rBmp
            Else
                DeleteDC rDC
            End If
        End If
        DeleteDC TmpDC
    End If
End Function

Private Function PaintText(ByVal hDc As Long, ByVal Text As String, ByVal x As Long, _
                           ByVal y As Long, ByVal Width As Long, ByVal Height As Long, _
                           ByVal dwFlag As Long) As Long
    Dim rcTemp As RECT
    With rcTemp
        .Left = x
        .Right = x + Width
        .Top = y
        .Bottom = y + Height
    End With
    PaintText = DrawText(hDc, Text, -1, rcTemp, dwFlag)
End Function

Private Function PaintColor(ByVal hDc As Long, ByVal x As Long, ByVal y As Long, _
                            ByVal Width As Long, ByVal Height As Long, _
                            ByVal clrColor As Long) As Long
    Dim rc      As RECT
    Dim hBrush  As Long
    With rc
        .Left = x
        .Top = y
        .Right = x + Width
        .Bottom = y + Height
    End With
    hBrush = CreateSolidBrush(TranslateColor(clrColor))
    PaintColor = FillRect(hDc, rc, hBrush)
    Call DeleteObject(hBrush)
End Function

Private Function TranslateColor(ByVal SysColor As Variant) As Long
    Dim lClr As Long
    Call OleTranslateColor(CLng(SysColor), 0, lClr)
    TranslateColor = lClr
End Function

Private Function HitTest(ByVal x As Long, ByVal y As Long) As Long
    If m_Count = 0 Or m_VCount = 0 Then HitTest = -1: Exit Function
    If x < ITM_MARGIN Or x > m_cX - ITM_MARGIN Then HitTest = -1: Exit Function
    If y < ITM_MARGIN Or y > m_cY - ITM_MARGIN Then HitTest = -1: Exit Function
    Dim i As Long
    For i = 0 To m_Count - 1
        If m_Menu(i).Visible Then
            If y >= m_Menu(i).rc.Top And y <= m_Menu(i).rc.Bottom Then
                HitTest = i
                Exit Function
            End If
        End If
    Next
    HitTest = -1
End Function

Private Sub OnSize()
    If m_hBGDC Then DeleteDC m_hBGDC
    If m_hDC Then DeleteDC m_hDC
    m_hBGDC = CreateMemDC(m_cX, m_cY)
    m_hDC = CreateMemDC(m_cX, m_cY)
    Call SelectObject(m_hDC, SendMessage(m_hOwner, WM_GETFONT, 0, 0))
    Call SetBkMode(m_hDC, TRANSPARENT)
    Call RedrawBackground
    Call RedrawAll(True)
End Sub

Private Sub OnShow()
    Call SetWindowPos(m_hWnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    m_hFocus = GetFocus
    If m_hFocus Then
        c_Subclass.ssc_Subclass m_hFocus
        c_Subclass.ssc_AddMsg m_hFocus, WM_ALLMSG, MSG_BEFORE
    End If
    If m_hOwner <> m_hFocus Then
        If m_hOwner Then
            c_Subclass.ssc_Subclass m_hOwner
            c_Subclass.ssc_AddMsg m_hOwner, WM_ALLMSG, MSG_BEFORE
        End If
    End If
End Sub

Private Sub OnHide()
    If m_hFocus Then c_Subclass.ssc_UnSubclass m_hFocus
    If m_hOwner <> m_hFocus Then
        If m_hOwner Then c_Subclass.ssc_UnSubclass m_hOwner
    End If
End Sub

Private Sub OnKeyDown(ByVal KeyCode As Long)
    If m_Count = 0 Or m_VCount = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To m_Count - 1
        If m_Menu(i).Visible And m_Menu(i).Enabled And m_Menu(i).lAcckey = KeyCode Then
            Call Hide
            RaiseEvent MenuClick(i)
            Exit For
        End If
    Next
End Sub

Private Sub OnMouseDown(ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    '
End Sub

Private Sub OnMouseMove(ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    Dim lIdx As Long
    lIdx = m_HoverIdx
    m_HoverIdx = HitTest(x, y)
    If lIdx <> m_HoverIdx Then
        Call RedrawOne(lIdx, True)
        Call RedrawOne(m_HoverIdx, True)
    End If
End Sub

Private Sub OnMouseUp(ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    Dim lIdx As Long
    lIdx = HitTest(x, y)
    If lIdx <> -1 Then
        If m_Menu(lIdx).Visible And m_Menu(lIdx).Enabled Then
            Call Hide
            RaiseEvent MenuClick(lIdx)
        End If
    End If
End Sub

Private Sub OnMouseLeave()
    Dim lIdx As Long
    lIdx = m_HoverIdx
    m_HoverIdx = -1
    Call RedrawOne(lIdx, True)
End Sub

Private Function GetTextWidth(ByVal hDc As Long, ByVal szChar As String) As Long
    '得到文字在hDC上的宽度
    Dim tSize As SIZE
    If Len(szChar) = 0 Then GetTextWidth = 0: Exit Function
    Call GetTextExtentPoint32(hDc, StrPtr(szChar), Len(szChar), tSize)
    GetTextWidth = tSize.cX
End Function

Private Sub C_Subclass_SubclassProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lhWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
    If lhWnd = m_hWnd Then
        Select Case uMsg
            Case WM_PAINT
            bHandled = True
            lReturn = 0
            Dim PS As PAINTSTRUCT
            Call BeginPaint(m_hWnd, PS)
            Call BitBlt(PS.hDc, 0, 0, m_cX, m_cY, m_hDC, 0, 0, vbSrcCopy)
            Call EndPaint(m_hWnd, PS)

            Case WM_ERASEBKGND
            bHandled = True
            lReturn = 0
            Case WM_SIZE
            Call OnSize

            Case WM_SHOWWINDOW
            If wParam = SW_HIDE Then
                Call OnHide
            Else
                Call OnShow
            End If
            Case WM_MOUSEMOVE
            c_Subclass.TrackMouseTracking m_hWnd
            Call OnMouseMove(wParam, LOWORD(lParam), HIWORD(lParam))
            Case WM_MOUSELEAVE
            Call OnMouseLeave
            Case WM_LBUTTONUP
            Call OnMouseUp(1, LOWORD(lParam), HIWORD(lParam))
        End Select
    Else
        Select Case uMsg
            Case WM_NCLBUTTONDOWN, WM_NCMBUTTONDOWN, WM_NCRBUTTONDOWN
            Call Hide
            Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
            Call Hide
            Case WM_NCACTIVATE
            '这里你可以看出为什么托盘菜单弹出前先 SetForegroundWindow 为父窗口了吧？
            If wParam = SW_HIDE Then Call Hide
            Case WM_MOUSEACTIVATE
            Call Hide
            '这里吃掉部分消息
            Case WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_CHAR, WM_KEYDOWN, WM_KEYUP
            bHandled = True
            lReturn = 0
        End Select
    End If
    If uMsg = WM_KEYDOWN Then
        Call OnKeyDown(wParam)
    End If
End Sub

Private Sub Class_Initialize()
    Set c_Subclass = New clsCSubclass
    m_HoverIdx = -1
End Sub

Private Sub Class_Terminate()
    If m_hBGDC Then DeleteDC m_hBGDC
    If m_hDC Then DeleteDC m_hDC
    If m_hWnd Then DestroyWindow m_hWnd
    Set c_Subclass = Nothing
End Sub

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Text(ByVal Index As Variant) As String
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    Text = m_Menu(lIdx).Text
End Property

Public Property Let Text(ByVal Index As Variant, ByVal newValue As String)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If m_Menu(lIdx).Text = newValue Then Exit Property
    m_Menu(lIdx).Text = newValue
    m_Menu(lIdx).lAcckey = GetAccessKey(newValue)
    Call RedrawOne(lIdx, True)
End Property

Public Property Get Key(ByVal Index As Variant) As String
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    Key = m_Menu(lIdx).Key
End Property

Public Property Let Key(ByVal Index As Variant, ByVal newValue As String)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If GetMenuIndex(newValue) <> -1 Then Exit Property
    m_Menu(lIdx).Key = newValue
End Property

Public Property Get Image(ByVal Index As Variant) As Long
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    Image = m_Menu(lIdx).Image
End Property

Public Property Let Image(ByVal Index As Variant, ByVal newValue As Long)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If m_Menu(lIdx).Image = newValue Then Exit Property
    m_Menu(lIdx).Image = newValue
    Call RedrawOne(lIdx, True)
End Property

Public Property Get Enabled(ByVal Index As Variant) As Boolean
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    Enabled = m_Menu(lIdx).Enabled
End Property

Public Property Let Enabled(ByVal Index As Variant, ByVal newValue As Boolean)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If m_Menu(lIdx).Enabled = newValue Then Exit Property
    m_Menu(lIdx).Enabled = newValue
    Call RedrawOne(lIdx, True)
End Property

Public Property Get Visible(ByVal Index As Variant) As Boolean
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    Visible = m_Menu(lIdx).Visible
End Property

Public Property Let Visible(ByVal Index As Variant, ByVal newValue As Boolean)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If m_Menu(lIdx).Visible = newValue Then Exit Property
    m_Menu(lIdx).Visible = newValue
    Call CalcSize
    Call RedrawOne(lIdx, True)
End Property

Public Property Get Checked(ByVal Index As Variant) As Boolean
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    Checked = m_Menu(lIdx).Checked
End Property

Public Property Let Checked(ByVal Index As Variant, ByVal newValue As Boolean)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If m_Menu(lIdx).Checked = newValue Then Exit Property
    m_Menu(lIdx).Checked = newValue
    Call RedrawOne(lIdx, True)
End Property

Public Property Get MenuType(ByVal Index As Variant) As MENUTYPECONSTANT
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    MenuType = m_Menu(lIdx).MenuType
End Property

Public Property Let MenuType(ByVal Index As Variant, ByVal newValue As MENUTYPECONSTANT)
    Dim lIdx As Long
    lIdx = GetMenuIndex(Index)
    If lIdx = -1 Then Exit Property
    If m_Menu(lIdx).MenuType = newValue Then Exit Property
    m_Menu(lIdx).MenuType = newValue
    Call CalcSize
    Call RedrawOne(lIdx, True)
End Property
