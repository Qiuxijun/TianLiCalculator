VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmLogo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14295
   Icon            =   "FrmLogo.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmLogo.frx":000C
   ScaleHeight     =   5310
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2400
      Top             =   2760
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   -1.00000e5
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Shape Shape1 
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   14295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "加载中……"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   14415
   End
End
Attribute VB_Name = "FrmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Function GetColumns() As Integer
Dim i%
    ' i = 1
    'Do While xlsheet.Cells(i, 1).Value <> ""
    '    i = i + 1
    'Loop
    'GetColumns = i - 1
    GetColumns = 0
End Function

Private Sub Timer1_Timer()
'On Error GoTo outs
Dim sumi%, sumj%, i%, j%
Dim t As String, temp() As String, temp2() As String, t2 As String
Me.Show
DoEvents

    
                Open App.Path + "\Data\怪物.txt" For Binary As #1
                   t = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   temp = Split(t, vbCrLf)
                   sumi = UBound(temp) + 1
                   sumj = 11

    ReDim Enemy(1 To sumi, 1 To sumj)
        For i = 1 To sumi
            temp2 = Split(temp(i - 1), vbTab)
            For j = 1 To sumj
                Enemy(i, j) = temp2(j - 1)
            Next
        Next


                Open App.Path + "\Data\武器.txt" For Binary As #1
                   t = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   temp = Split(t, vbCrLf)
                   sumi = UBound(temp) + 1
                   sumj = 6

    ReDim WeaponList(1 To sumi, 1 To sumj)
        For i = 1 To sumi
            temp2 = Split(temp(i - 1), vbTab)
            For j = 1 To sumj
                WeaponList(i, j) = temp2(j - 1)
            Next
        Next


                Open App.Path + "\Data\角色.txt" For Binary As #1
                   t = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   temp = Split(t, vbCrLf)
                   sumi = UBound(temp) + 1
                   sumj = 5

    ReDim CharList(1 To sumi, 1 To sumj)
        For i = 1 To sumi
            temp2 = Split(temp(i - 1), vbTab)
            For j = 1 To sumj
                CharList(i, j) = temp2(j - 1)
            Next
        Next
        

    
    
    ReDim ArtList(0 To 0, 1 To 11) As String
    FrmMain.UpdateArtList
                    
    t = getHtmlStr("http://spongem.com/ajglz/ys/1.html")
    
    i = InStr(1, t, "version ")
    j = InStr(1, t, "</BODY>")
    If i > 0 Then
    If Dir(App.Path + "\res\version") <> "" Then
        Open App.Path + "\res\version" For Binary As #1
            t2 = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
    Else
        t2 = "[版本号暂缺]"
    End If
        
        t = Mid(t, i + 8, j - i - 8)
        If t <> t2 Then
            MsgBox "目前最新版本是" + t + "，您的版本是" + t2 + "，可打开在线更新软件Update.exe进行更新！", , "更新提醒"
        End If
    End If
    
   'Form1.Show
   FrmMain.Show
    Unload Me

End Sub


Private Function getHtmlStr$(strUrl$)
'On Error Resume Next
WebBrowser1.Navigate strUrl
Do While WebBrowser1.ReadyState <> 4
DoEvents
Loop
getHtmlStr = WebBrowser1.Document.body.outerhtml
End Function
