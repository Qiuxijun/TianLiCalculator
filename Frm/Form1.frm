VERSION 5.00
Object = "{82C2E93F-4319-4516-962C-8699DDF52122}#1.0#0"; "BSkin.ocx"
Begin VB.Form FrmBuff 
   BorderStyle     =   0  'None
   Caption         =   "ѡ��Buff"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6030
   StartUpPosition =   1  '����������
   Begin BSkin.Container ctnMain 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      BackColor       =   16761024
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϸ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Tag             =   "10"
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˵�ǰ���ԣ��ң���10%"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Tag             =   "10"
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��Buff"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Top             =   45
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmBuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'������Ӱ����������������������������������������������������������������������������������������������
Private FormShadow As clsShadow





Private Sub Form_Load()

    
    If FormShadow Is Nothing Then Set FormShadow = New clsShadow '������Ӱ
    With FormShadow
        .Depth = 3.5
        .Color = vbBlack
        .Transparency = 100
        .Shadow Me
    End With
    

    
    Call LoadTheme '���ؽ�����
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not FormShadow Is Nothing Then Set FormShadow = Nothing
End Sub

Private Sub ctnMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveForm Me
End Sub

Private Sub lblAuthor_Click()

End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveForm Me
End Sub

Private Sub LoadTheme()
    Me.BackColor = &HE9E9E9
    ctnMain.BackColor = vbWhite
End Sub

Private Sub lblClose_Click() '�ر�
    Unload Me
End Sub
