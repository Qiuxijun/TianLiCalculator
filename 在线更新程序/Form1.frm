VERSION 5.00
Object = "{82C2E93F-4319-4516-962C-8699DDF52122}#1.0#0"; "BSkin.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "天理计算器-在线更新程序"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4995
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4995
   StartUpPosition =   1  '所有者中心
   Begin BSkin.CommandButton CommandButton2 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Text            =   "检查更新"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   735
      Left            =   -10000
      TabIndex        =   2
      Top             =   -10000
      Width           =   975
      ExtentX         =   1720
      ExtentY         =   1296
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
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "下载未开始"
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin BSkin.ProgressBar ProgressBar5 
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      Percentage      =   0   'False
   End
   Begin BSkin.Downloader Downloader1 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StrFormatByteSize Lib "shlwapi" _
Alias "StrFormatByteSizeA" _
(ByVal dw As Long, _
ByVal pszBuf As String, _
ByVal cchBuf As Long) As Long



Dim t As String, DListf() As String, DListt() As String, now As Integer

Private Function VBStrFormatByteSize(ByVal LngSize As Double) As String
On Error Resume Next
    Dim strSize As String * 128
    Dim strData As String
    Dim lPos As Long
    StrFormatByteSize LngSize, strSize, 128
    lPos = InStr(1, strSize, Chr$(0))
    strData = Left$(strSize, lPos - 1)
    VBStrFormatByteSize = strData
End Function


Private Sub ChangeFileType(path As String)
Dim t As String
                Open path For Binary As #1
                   t = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                 
                Open path For Output As #1
                   Print #1, Replace(t, vbLf, vbCrLf);
                 Close #1
                 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CommandButton2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo outs
Dim t2 As String, i%, j%
If CommandButton2.Text = "检查更新" Then
    t = getHtmlStr("http://spongem.com/ajglz/ys/1.html")

    i = InStr(1, t, "version ")
    j = InStr(1, t, "</BODY>")
    If i > 0 Then
        Open App.path + "\res\version" For Binary As #1
            t2 = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        t = Mid(t, i + 8, j - i - 8)
        If t <> t2 Then
            Shell "taskkill /f /im 天理计算器.exe", vbHide
            MsgBox "目前最新版本是" + t + "，您的版本是" + t2 + "，准备开始下载最新资源文件。", , "检查更新"
            CommandButton2.Visible = False
            ProgressBar5.Visible = True
            Text3.Visible = True
            now = 1
            DListf(1) = "1.exe"
            DListf(2) = "1.txt"
            DListf(3) = "2.txt"
            DListf(4) = "3.txt"
            DListf(5) = "4.txt"
            DListt(1) = App.path + "\天理计算器.exe"
            DListt(2) = App.path + "\Data\角色.txt"
            DListt(3) = App.path + "\Data\武器.txt"
            DListt(4) = App.path + "\Data\怪物.txt"
            DListt(5) = App.path + "\Data\圣遗物.txt"
            
            
            
            BeginDownload
        Else
            MsgBox "目前最新版本是" + t + "，您的版本是" + t2 + "，已经是最新的资源文件。", , "检查更新"
            End
        End If
    End If
Else
    End
End If
Exit Sub
outs:
MsgBox "您的系统不支持在线更新，请前往百度网盘手动下载最新版本。", , "检查更新"
End Sub

Private Sub BeginDownload()
On Error Resume Next
Dim path As String
    Kill DListt(now)
    Label1.Caption = "正在下载：" + DListt(now)
    Downloader1.BeginDownload "http://spongem.com/ajglz/ys/tljsq/" + DListf(now), DListt(now)
    now = now + 1
End Sub

Private Sub ErrorSolve()

End Sub

Private Sub Downloader1_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    '下载进度
    ProgressBar5.Value = Format((CurBytes / MaxBytes) * 100, "0.00")
    Text3.Text = VBStrFormatByteSize(CurBytes) & " / " & VBStrFormatByteSize(MaxBytes)
End Sub

Private Sub Downloader1_DownloadError(SaveFile As String)
    '下载失败
    Text3.Text = "下载失败，点击重试"
    ProgressBar5.Style = Style4
End Sub

Private Sub Downloader1_DownloadComplete(MaxBytes As Long, SaveFile As String)
    '下载完毕
    If now = 6 Then
        ChangeFileType App.path + "\data\角色.txt"
        ChangeFileType App.path + "\data\武器.txt"
        ChangeFileType App.path + "\data\怪物.txt"
        ChangeFileType App.path + "\data\圣遗物.txt"
        CheckFile
    End If
    
    If now <= UBound(DListt) Then
        BeginDownload
        
    Else
    
    Text3.Text = "下载完毕"
    ProgressBar5.Style = Style3
    Label1.Caption = ""
        Open App.path + "\res\version" For Output As #1
            Print #1, t;
        Close #1
        CommandButton2.Text = "退出"
        CommandButton2.Visible = True
        ProgressBar5.Visible = False
    End If
End Sub

Private Function getHtmlStr$(strUrl$)
WebBrowser1.Navigate strUrl
Do While WebBrowser1.ReadyState <> 4
DoEvents
Loop
getHtmlStr = WebBrowser1.Document.body.outerhtml
End Function

Private Sub CheckFile()
Dim tall As String, tempc() As String, tempR() As String, tempAll() As String, i%, j%, sumi%, sumj%
                Open App.path + "\Data\角色.txt" For Binary As #1
                   tall = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1

                 
                   tempc = Split(tall, vbCrLf)
                   sumi = UBound(tempc) + 1
                    For i = 1 To sumi
                        If Dir(App.path + "\Data\Data\C" + CStr(i)) = "" Then AddList "C" + CStr(i), App.path + "\Data\Data\" + "C" + CStr(i)
                        If Dir(App.path + "\Data\Data\C" + CStr(i) + "_2") = "" Then AddList "C" + CStr(i) + "_2", App.path + "\Data\Data\" + "C" + CStr(i) + "_2"
                        If Dir(App.path + "\Res\Public\C" + CStr(i) + ".png") = "" Then AddList "C" + CStr(i) + ".png", App.path + "\Res\Public\" + "C" + CStr(i) + ".png"
                    Next
                    For i = 1 To sumi * 2 + 5
                        If Dir(App.path + "\Res\Public\s_" + CStr(i) + ".jpg") = "" Then AddList "s_" + CStr(i) + ".jpg", App.path + "\Res\Public\" + "s_" + CStr(i) + ".jpg"
                    Next
                    
                Open App.path + "\Data\武器.txt" For Binary As #1
                   tall = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   tempc = Split(tall, vbCrLf)
                   sumi = UBound(tempc) + 1
                   sumj = 6
                   ReDim tempAll(1 To sumi, 1 To sumj) As String
                    For i = 1 To sumi
                        tempR = Split(tempc(i - 1), vbTab)
                        If Dir(App.path + "\Res\Public\W_" + CStr(i) + ".png") = "" Then AddList "W_" + CStr(i) + ".png", App.path + "\Res\Public\" + "W_" + CStr(i) + ".png"
                        For j = 1 To sumj
                            If Dir(App.path + "\Data\Data\" + tempR(4) + "_" + tempR(3)) = "" Then
                                Open App.path + "\Data\Data\" + tempR(4) + "_" + tempR(3) For Output As #1
                                Close #1
                                AddList tempR(4) + "_" + tempR(3), App.path + "\Data\Data\" + tempR(4) + "_" + tempR(3)
                            End If
                        Next
                    Next
                    
                Open App.path + "\Data\圣遗物.txt" For Binary As #1
                   tall = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   tempc = Split(tall, vbCrLf)
                   sumi = UBound(tempc) + 1
                    For i = 1 To sumi
                        For j = 1 To 5
                            If Dir(App.path + "\Res\Public\A" + CStr(i) + "_" + CStr(j) + ".png") = "" Then AddList "A" + CStr(i) + "_" + CStr(j) + ".png", App.path + "\Res\Public\" + "A" + CStr(i) + "_" + CStr(j) + ".png"
                        Next
                    Next
                
                    
End Sub

Private Sub AddList(f As String, t As String)
                            ReDim Preserve DListf(1 To UBound(DListf) + 1)
                            ReDim Preserve DListt(1 To UBound(DListt) + 1)
                            
                            DListf(UBound(DListf)) = f
                            DListt(UBound(DListt)) = t
End Sub

Private Sub Form_Load()
ReDim DListt(1 To 5) As String
ReDim DListf(1 To 5) As String
End Sub

Private Sub Text3_Click()
If Text3.Text = "下载失败，点击重试" Then
    now = now - 1
    BeginDownload
End If
End Sub
