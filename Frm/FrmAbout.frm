VERSION 5.00
Object = "{82C2E93F-4319-4516-962C-8699DDF52122}#1.0#0"; "BSkin.ocx"
Begin VB.Form FrmAbout 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "����ʥ����"
   ClientHeight    =   5940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   StartUpPosition =   1  '����������
   Begin BSkin.CommandButton CommandButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Text            =   "����json"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BSkin.CommandButton btnClose 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackgroundHover =   14737632
      TextColorHover  =   0
      BorderHover     =   12632256
      Text            =   "ȷ��"
      MouseDownBackground=   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BSkin.Container ctnMain 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9128
      BackColor       =   16761024
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   18
         Text            =   "0"
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   17
         Text            =   "0"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   16
         Text            =   "0"
         Top             =   3600
         Width           =   1575
      End
      Begin BSkin.ComboBox ComboBox3 
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   4560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         ShadowColorText =   6908265
         Text            =   "ComboBox4"
      End
      Begin BSkin.ComboBox ComboBox3 
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         ShadowColorText =   6908265
         Text            =   "ComboBox4"
      End
      Begin BSkin.ComboBox ComboBox3 
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         ShadowColorText =   6908265
         Text            =   "ComboBox4"
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   12
         Text            =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin BSkin.ComboBox ComboBox3 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         ShadowColorText =   6908265
         Text            =   "ComboBox3"
      End
      Begin BSkin.ComboBox ComboBox2 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         ShadowColorText =   6908265
      End
      Begin BSkin.ComboBox ComboBox1 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         ShadowColorText =   6908265
      End
      Begin BSkin.AlphaImage AlphaImage2 
         Height          =   1050
         Left            =   4680
         Top             =   3480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1852
         Image           =   "FrmAbout.frx":000C
         Props           =   5
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   810
         Left            =   1440
         Top             =   1080
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
         Image           =   "FrmAbout.frx":26B7
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage 
         Height          =   750
         Index           =   4
         Left            =   4800
         Top             =   1080
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Image           =   "FrmAbout.frx":33B7
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage 
         Height          =   750
         Index           =   3
         Left            =   3960
         Top             =   1080
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Image           =   "FrmAbout.frx":3A23
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage 
         Height          =   750
         Index           =   2
         Left            =   3120
         Top             =   1080
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Image           =   "FrmAbout.frx":3E6B
         Props           =   5
      End
      Begin BSkin.AlphaImage AlphaImage 
         Height          =   750
         Index           =   1
         Left            =   2280
         Top             =   1080
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Image           =   "FrmAbout.frx":455B
         Props           =   5
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "�����ԣ�"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��"
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
         TabIndex        =   6
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ����װ��"
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin BSkin.AlphaImage AlphaImage 
         Height          =   750
         Index           =   0
         Left            =   1440
         Top             =   1080
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         Image           =   "FrmAbout.frx":4A74
         Props           =   5
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʥ����"
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
         TabIndex        =   3
         Top             =   45
         Width           =   1050
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
         TabIndex        =   1
         Top             =   0
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'������Ӱ����������������������������������������������������������������������������������������������
Private FormShadow As clsShadow
Dim Selected As Integer

Private Type Tags
    Atype As String
    avalue As String
End Type

Private Sub AlphaImage_Click(Index As Integer, ByVal Button As Integer)
FrmMain.zMove1 AlphaImage1, AlphaImage(Index).Left + 40, AlphaImage1.Top, True
Selected = Index + 1
AlphaImage2.LoadImage_FromFile (App.Path + "\Res\Public\A" + CStr(ComboBox1.ListIndex) + "_" + CStr(Selected) + ".jpg")
    If Index > 1 Then
        ComboBox2.Visible = True
        Label3.Visible = True
        ComboBox2.Clear
        ComboBox2.AddItem "����ֵ%"
        ComboBox2.AddItem "������%"
        ComboBox2.AddItem "������%"
            If Index = 2 Then
                ComboBox2.AddItem "Ԫ�ؾ�ͨ"
                ComboBox2.AddItem "Ԫ�س���Ч��%"
            End If
            If Index = 3 Then
                ComboBox2.AddItem "Ԫ�ؾ�ͨ"
                ComboBox2.AddItem "��Ԫ���˺�%"
                ComboBox2.AddItem "ˮԪ���˺�%"
                ComboBox2.AddItem "��Ԫ���˺�%"
                ComboBox2.AddItem "��Ԫ���˺�%"
                ComboBox2.AddItem "��Ԫ���˺�%"
                ComboBox2.AddItem "��Ԫ���˺�%"
                ComboBox2.AddItem "�����˺�%"
            End If
            If Index = 4 Then
                ComboBox2.AddItem "Ԫ�ؾ�ͨ"
                ComboBox2.AddItem "������%"
                ComboBox2.AddItem "�����˺�%"
                ComboBox2.AddItem "���Ƽӳ�%"
            End If
        ComboBox2.ListIndex = 2
    Else
        ComboBox2.Visible = False
        Label3.Visible = False
    End If
End Sub

Private Sub ComboBox1_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
AlphaImage2.LoadImage_FromFile (App.Path + "\Res\Public\A" + CStr(ComboBox1.ListIndex) + "_" + CStr(Selected) + ".jpg")
End Sub

Private Sub CommandButton1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MsgBox "��json�ļ���ҷ������հ״����˰�ť�ԣ����ɵ���json�ļ���Ŀǰ����ʹ����Ī��ռ���̼��ݵ�json��ʽ��", , "����json"
End Sub

Private Sub Form_Load()
Dim i%, v2 As Variant
 v2 = Array("ȾѪ����ʿ��", "����������Ů", "�԰�֮��", "���ҵ���֮ħŮ", "����֮Ӱ", "�ɹ��һ������", "���˴�ص�����", "��ɵ�����", "ƽϢ���׵�����", "ǧ���ι�", "������;����ʿ", "����֮��", "��Ե֮��ӡ", "�Ƕ�ʿ����Ļ��", "׷��֮ע��", "���׵�ʢŭ", "�ƹŵ�����", "��������֮��")
    SetFormRGBAIcon Me, 0 '����RGBAͨ��ͼ��
    SetWindowIcon Me.hWnd
    
    If FormShadow Is Nothing Then Set FormShadow = New clsShadow '������Ӱ
    With FormShadow
        .Depth = 3.5
        .Color = vbBlack
        .Transparency = 100
        .Shadow Me
    End With
    
    Me.Caption = "����ʥ����"
    Selected = 1
    
    Call LoadTheme '���ؽ�����
    For i = 0 To UBound(v2)
        ComboBox1.AddItem v2(i), , FrmMain.ImageTemp2(i + 1).Image
    Next
    ComboBox1.ListIndex = 1
    
    For i = 0 To 3
        ComboBox3(i).AddItem "������%"
        ComboBox3(i).AddItem "�����˺�%"
        ComboBox3(i).AddItem "������%"
        ComboBox3(i).AddItem "������"
        ComboBox3(i).AddItem "Ԫ�س���Ч��%"
        ComboBox3(i).AddItem "Ԫ�ؾ�ͨ"
        ComboBox3(i).AddItem "����ֵ%"
        ComboBox3(i).AddItem "����ֵ"
        ComboBox3(i).AddItem "������%"
        ComboBox3(i).AddItem "������"
        ComboBox3(i).ListIndex = i + 1
    Next
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo outs
Dim t As String, tempT As String, tempV As String, tempN As String, flag As Boolean
Dim i As Long, j As Long, k As Long, l As Long, M As Long
Dim n%
Dim tag(1 To 5) As Tags, now As Integer
Dim ans As String, anssum As Integer, anssum2 As Integer
Dim tall As String, tempc() As String, tempR() As String, tempAll() As String, sumi As Integer, sumj As Integer, ii%, jj%
Dim ados As Object

Dim ReturnEncoding As Encoding
t = "     " + Data.Files(1)

If Right(t, 5) <> ".json" Then
    If MsgBox("��ǰ�ļ�����json�ļ����Ƿ�������룿", vbYesNo, "����json") = vbYes Then
    Else
        Exit Sub
    End If
End If

    If MsgBox("�Ƿ񸲸�֮ǰ¼���ʥ�������ǽ��и��ǣ���������׷�ӡ�", vbYesNo, "����json") = vbYes Then
        anssum = 0
    Else
        anssum = 1
        Do While Dir(App.Path + "\data\artifact\" + CStr(anssum)) <> ""
            anssum = anssum + 1
        Loop
        anssum = anssum - 1
    End If
    anssum2 = 0
    

                     
                     

                Open App.Path + "\Data\ʥ����.txt" For Binary As #1
                   tall = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   tempc = Split(tall, vbCrLf)
                   sumi = UBound(tempc) + 1
                   sumj = 7
                   
                   ReDim tempAll(1 To sumi, 1 To sumj) As String
                    For ii = 1 To sumi
                        tempR = Split(tempc(ii - 1), vbTab)
                        For jj = 1 To sumj
                            tempAll(ii, jj) = tempR(jj - 1)
                        Next
                    Next
                    
                    
                    Open Data.Files(1) For Binary As #1
                       t = StrConv(InputB(LOF(1), 1), vbUnicode)
                     Close #1 '�ȳ���ANSI
                     
                    For ii = 1 To sumi
                        For jj = 3 To 7
                           If InStr(1, t, tempAll(ii, jj)) > 0 Then 'ȷʵ��ANSI
                            GoTo begindo
                           End If
                        Next
                    Next


             '�Ҳ���ANSI���ڵ�֤�ݣ��ĳ�utf-8



            Set ados = CreateObject("adodb.stream")
            
            With ados
                .Charset = "utf-8"
                .Type = 2
                .Open
                .LoadFromFile Data.Files(1)
                 t = .ReadText '��ȡUTF-8
                .Close
            End With
            

            
            
begindo:
            
            t = Replace(t, vbLf, vbCrLf) '��UXת��Windows
            t = Replace(t, vbCrLf, "")
            t = Replace(t, " ", "")
            


                    

                 i = 1
                 j = 0
                 
                 l = InStr(i, t, """detailName""")
                 
                 
        Do While l > 0
                 
                M = InStr(l, t, ",")
                tempN = Mid(t, l + 14, M - l - 15)
                If tempN = "���޹��֮��" Then tempN = "���޹��֮��"
                If tempN = "Զ������Ů����" Then tempN = "Զ������Ů֮��"
                    For ii = 1 To sumi
                        For jj = 3 To 7
                            If tempN = tempAll(ii, jj) Then
                                tempN = "a" + CStr(ii) + "_" + CStr(jj - 2)
                                Exit For
                            End If
                        Next
                    Next
                    
                    
                 
                 l = InStr(i, t, """star""")
                 now = 1
                 For n = 1 To 5
                    tag(n).Atype = "������"
                    tag(n).avalue = "0"
                 Next
                 ans = "" '��ʼ��һ��ʥ����
                 i = InStr(M, t, """name""")
                
                 
                 
            Do While (i < l) And (i <> 0)
                     j = InStr(i, t, ",")
                     k = InStr(j, t, "}")
                     tempT = Mid(t, i + 8, j - i - 9)
                     tempV = Mid(t, j + 9, k - j - 9)
                     tempV = Replace(tempV, ":", "")
                     flag = False
    
                     Select Case tempT
                        Case "lifeStatic"
                            tag(now).Atype = "����ֵ"
                        Case "lifePercentage"
                            tag(now).Atype = "����ֵ%"
                            flag = True
                        Case "attackStatic"
                            tag(now).Atype = "������"
                        Case "attackPercentage"
                            tag(now).Atype = "������%"
                            flag = True
                        Case "defendStatic"
                            tag(now).Atype = "������"
                        Case "defendPercentage"
                            tag(now).Atype = "������%"
                            flag = True
                        Case "elementalMastery"
                            tag(now).Atype = "Ԫ�ؾ�ͨ"
                        Case "recharge"
                            tag(now).Atype = "Ԫ�س���Ч��%"
                            flag = True
                        Case "criticalDamage"
                            tag(now).Atype = "�����˺�%"
                            flag = True
                        Case "critical"
                            tag(now).Atype = "������%"
                            flag = True
                        Case "cureEffect"
                            tag(now).Atype = "���Ƽӳ�%"
                            flag = True
                        Case "thunderBonus"
                            tag(now).Atype = "��Ԫ���˺�%"
                            flag = True
                        Case "fireBonus"
                            tag(now).Atype = "��Ԫ���˺�%"
                            flag = True
                        Case "waterBonus"
                            tag(now).Atype = "ˮԪ���˺�%"
                            flag = True
                        Case "iceBonus"
                            tag(now).Atype = "��Ԫ���˺�%"
                            flag = True
                        Case "windBonus"
                            tag(now).Atype = "��Ԫ���˺�%"
                            flag = True
                        Case "rockBonus"
                            tag(now).Atype = "��Ԫ���˺�%"
                            flag = True
                        Case "physicalBonus"
                            tag(now).Atype = "�����˺�%"
                            flag = True
                    End Select
                    tag(now).avalue = CStr(Val(tempV) * IIf(flag, 100, 1))
                    now = now + 1
                    i = InStr(i + 1, t, """name""")
            Loop
            
                
                

                
                
                ans = tempN + vbCrLf
                 For n = 1 To 5
                    ans = ans + tag(n).Atype + vbCrLf
                    ans = ans + tag(n).avalue + vbCrLf
                 Next
                 
                anssum = anssum + 1
                anssum2 = anssum2 + 1
                Open App.Path + "\data\Artifact\" + CStr(anssum) For Output As #1
                    Print #1, ans;
                Close #1
                
                
                l = InStr(l, t, """detailName""") '��һ��ʥ����
                
        Loop
    MsgBox "������ϣ���������" + CStr(anssum2) + "��ʥ���", , "����json"
    With FrmMain
         .UpdateArtList
         If .Container4.Visible = True Then
            .Container4.Visible = False
            .Frame2.Caption = "ʥ���﷽��"
            .Picture2.Visible = False
            .Picture2.tag = ""
            .Container4.tag = ""
        End If
    End With
    Unload Me
Exit Sub
outs:
MsgBox "����δ��ɣ���������" + CStr(anssum2) + "��ʥ���", , "����json"
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

Private Sub btnClose_Click() '�ر�
'On Error GoTo EX
Dim i As Integer, j As Integer, t As Single, n As Integer, temp() As String, t2() As String
n = UBound(ArtList)
ReDim temp(0 To n, 1 To 11)

ReDim t2(1 To 11)
    For i = 0 To n
        For j = 1 To 11
            temp(i, j) = ArtList(i, j)
        Next
    Next
ReDim ArtList(0 To n + 1, 1 To 11) As String
    For i = 1 To n
        For j = 1 To 11
            ArtList(i, j) = temp(i, j)
        Next
    Next
    

    

t2(1) = "a" + CStr(ComboBox1.ListIndex) + "_" + CStr(Selected)


If Selected = 1 Then
    t2(2) = "����ֵ"
    t2(3) = "4780"
End If

If Selected = 2 Then
    t2(2) = "������"
    t2(3) = "311"
End If



If Selected >= 3 Then
    t2(2) = ComboBox2.Text
 If ComboBox2.Text = "����ֵ%" Then t = 46.6
 If ComboBox2.Text = "������%" Then t = 46.6
 If ComboBox2.Text = "������%" Then t = 58.3
 If ComboBox2.Text = "Ԫ�ؾ�ͨ" Then t = 187
 If ComboBox2.Text = "��Ԫ���˺�%" Then t = 46.6
 If ComboBox2.Text = "ˮԪ���˺�%" Then t = 46.6
 If ComboBox2.Text = "��Ԫ���˺�%" Then t = 46.6
 If ComboBox2.Text = "��Ԫ���˺�%" Then t = 46.6
 If ComboBox2.Text = "��Ԫ���˺�%" Then t = 46.6
 If ComboBox2.Text = "��Ԫ���˺�%" Then t = 46.6
 If ComboBox2.Text = "�����˺�%" Then t = 58.3
 If ComboBox2.Text = "Ԫ�س���Ч��%" Then t = 51.8
 If ComboBox2.Text = "������%" Then t = 31.1
 If ComboBox2.Text = "�����˺�%" Then t = 62.2
 If ComboBox2.Text = "���Ƽӳ�%" Then t = 35.9
            t2(3) = CStr(t)
End If


For j = 0 To 3
    t2(4 + j * 2) = ComboBox3(j).Text
    t2(5 + j * 2) = Text1(j).Text
Next

For j = 1 To 11
    ArtList(n + 1, j) = t2(j)
Next

j = 1
    Do While Dir(App.Path + "\Data\Artifact\" + CStr(j)) <> ""
        j = j + 1
    Loop
    
    Open App.Path + "\Data\Artifact\" + CStr(j) For Output As #1
        For i = 1 To 11
            Print #1, t2(i)
        Next
    Close #1


    If FrmMain.Container4.Visible Then
        If Val(FrmMain.Container4.tag) = 0 Then
            FrmMain.ShowArtBoxC
        Else
            FrmMain.ShowArtBoxA Val(FrmMain.Picture2.tag), Val(FrmMain.Container4.tag)
        End If
    End If
    
    Unload Me
    Exit Sub
EX:
    MsgBox "����ʧ�ܣ���ȷ��δ�򿪱����������ʥ����excel�ĵ���", , "����"
End Sub

