VERSION 5.00
Object = "{82C2E93F-4319-4516-962C-8699DDF52122}#1.0#0"; "BSkin.ocx"
Begin VB.Form FrmChar 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "选择角色"
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   Icon            =   "FrmChar.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   6135
   StartUpPosition =   1  '所有者中心
   Begin BSkin.Container ctnMain 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   12938
      BackColor       =   16761024
      Begin BSkin.Container Container1 
         Height          =   6855
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   12091
         Begin BSkin.ScrollBar ScrollBar1 
            Height          =   6735
            Left            =   5880
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   11880
            Max             =   100
            Speed           =   1
         End
         Begin BSkin.Container Container2 
            Height          =   7095
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   12515
            Begin BSkin.AlphaImage ShowImage 
               Height          =   1590
               Index           =   0
               Left            =   120
               Top             =   0
               Visible         =   0   'False
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   2805
               Image           =   "FrmChar.frx":000C
               Props           =   5
            End
         End
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   810
         Index           =   6
         Left            =   240
         Top             =   480
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
         Image           =   "FrmChar.frx":67A0
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   540
         Index           =   5
         Left            =   3960
         Tag             =   "雷"
         Top             =   600
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "FrmChar.frx":74A0
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   540
         Index           =   4
         Left            =   3240
         Tag             =   "岩"
         Top             =   600
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "FrmChar.frx":7944
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   540
         Index           =   3
         Left            =   2520
         Tag             =   "风"
         Top             =   600
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "FrmChar.frx":7E21
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   540
         Index           =   2
         Left            =   1800
         Tag             =   "冰"
         Top             =   600
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "FrmChar.frx":836E
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   540
         Index           =   1
         Left            =   1080
         Tag             =   "水"
         Top             =   600
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "FrmChar.frx":884F
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   540
         Index           =   0
         Left            =   360
         Tag             =   "火"
         Top             =   600
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   953
         Image           =   "FrmChar.frx":8CF4
         Props           =   5
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "×"
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
         Left            =   5760
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选择角色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'窗体阴影———————————————————————————————————————————————
Private FormShadow As clsShadow
Dim Selected As Integer


Private Sub AlphaImage1_Click(Index As Integer, ByVal Button As Integer)
On Error Resume Next
Dim i As Integer, j As Integer, n As Integer

If Index < 6 Then
    FrmMain.zMove1 AlphaImage1(6), AlphaImage1(Index).Left - 100, 480, True
        For i = 1 To 500
            Unload ShowImage(i)
        Next
    j = UBound(CharList)
    n = 0
        For i = 1 To j
            If CharList(i, 4) = AlphaImage1(Index).tag Then
                n = n + 1
                Load ShowImage(n)
                ShowImage(n).Left = 120 + 1900 * IIf(n Mod 3 = 0, 2, n Mod 3 - 1)
                ShowImage(n).Top = 0 + 1900 * Int((n - 1) / 3)
                ShowImage(n).LoadImage_FromFile App.Path + "\res\public\c" + CStr(i) + ".png"
                ShowImage(n).Visible = True
                ShowImage(n).tag = CStr(i)
                If i = Val(FrmMain.AlphaImageChar.tag) Then ShowImage(n).Opacity = 30
            End If
        Next
End If

End Sub

Private Sub Form_Load()

    
    If FormShadow Is Nothing Then Set FormShadow = New clsShadow '窗体阴影
    With FormShadow
        .Depth = 3.5
        .Color = vbBlack
        .Transparency = 100
        .Shadow Me
    End With
    

    
    Call LoadTheme '加载界面风格
    AlphaImage1_Click 0, 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not FormShadow Is Nothing Then Set FormShadow = Nothing
End Sub

Private Sub ctnMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveForm Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveForm Me
End Sub

Private Sub LoadTheme()
    Me.BackColor = &HE9E9E9
    ctnMain.BackColor = vbWhite
End Sub

Private Sub lblClose_Click() '关闭
    Unload Me
End Sub

Private Sub ShowImage_Click(Index As Integer, ByVal Button As Integer)
If Me.Caption = "选择角色" Then
If Val(ShowImage(Index).tag) = Val(FrmMain.AlphaImageChar.tag) Then
Unload Me
Exit Sub
End If
FrmMain.AlphaImageChar.tag = ShowImage(Index).tag
FrmMain.LoadChar (Val(ShowImage(Index).tag))
                Open App.Path + "\Data\User\C0" For Output As #1
                   Print #1, ShowImage(Index).tag;
                 Close #1
ReloadTip = True
Unload Me
Else

If Val(ShowImage(Index).tag) = Val(FrmMain.AlphaImageWeap.tag) Then
Unload Me
Exit Sub
End If
FrmMain.LoadWeapon (Val(ShowImage(Index).tag))

FrmMain.SaveSet0
ReloadTip = True
Unload Me
End If
End Sub

Private Sub ShowImage_MouseEnter(Index As Integer)
If Val(ShowImage(Index).tag) <> Val(FrmMain.AlphaImageChar.tag) And Me.Caption = "选择角色" Then ShowImage(Index).FadeInOut 40, 8
If Val(ShowImage(Index).tag) <> Val(FrmMain.AlphaImageWeap.tag) And Me.Caption = "选择武器" Then ShowImage(Index).FadeInOut 40, 8
End Sub
Private Sub ShowImage_MouseExit(Index As Integer)
If Val(ShowImage(Index).tag) <> Val(FrmMain.AlphaImageChar.tag) And Me.Caption = "选择角色" Then ShowImage(Index).FadeInOut 100, 8
If Val(ShowImage(Index).tag) <> Val(FrmMain.AlphaImageWeap.tag) And Me.Caption = "选择武器" Then ShowImage(Index).FadeInOut 100, 8
End Sub

Sub ShowWeapon(Atype As Integer)
On Error Resume Next
Dim i%, j%, n%
        For i = 1 To 500
            Unload ShowImage(i)
        Next
    j = UBound(WeaponList)
    n = 0
    Container1.Top = 480
    Me.Caption = "选择武器"
    lblTitle.Caption = "选择武器"
        For i = 1 To j
            If Val(WeaponList(i, 3)) = Atype Then
                n = n + 1
                Load ShowImage(n)
                ShowImage(n).Left = 120 + 900 * IIf(n Mod 6 = 0, 5, n Mod 6 - 1)
                ShowImage(n).Top = 0 + 900 * Int((n - 1) / 6)
                ShowImage(n).LoadImage_FromFile App.Path + "\res\public\w_" + CStr(i) + ".png"
                ShowImage(n).Visible = True
                ShowImage(n).Width = ShowImage(n).Width / 2
                ShowImage(n).Height = ShowImage(n).Height / 2
                ShowImage(n).tag = CStr(i)
                If i = Val(FrmMain.AlphaImageWeap.tag) Then ShowImage(n).Opacity = 30
            End If
        Next
End Sub
