VERSION 5.00
Object = "{82C2E93F-4319-4516-962C-8699DDF52122}#1.0#0"; "BSkin.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "���������"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13080
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   872
   StartUpPosition =   1  '����������
   Begin BSkin.Container ctnMain 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   14208
      BackColor       =   12640511
      Begin BSkin.Container Container1 
         Height          =   5895
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   10398
         BackColor       =   16711935
         Begin BSkin.Frame Frame6 
            Height          =   5895
            Left            =   0
            TabIndex        =   177
            Top             =   0
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   10398
            Orientation     =   2
            BackColor       =   16244694
            ColorGradient2  =   16241606
            Caption         =   "����/��ɫ/ʥ����Ч��"
            Icon            =   "FrmMain.frx":000C
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin BSkin.ProgressBar ProgressBar1 
               Height          =   300
               Left            =   6600
               TabIndex        =   256
               Top             =   4210
               Visible         =   0   'False
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   529
               Value           =   50
               Percentage      =   0   'False
            End
            Begin BSkin.Container Container6 
               Height          =   4215
               Left            =   6600
               TabIndex        =   222
               Top             =   0
               Visible         =   0   'False
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   7435
               Begin VB.TextBox Textt 
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   960
                  TabIndex        =   257
                  Text            =   "��"
                  Top             =   3840
                  Width           =   1815
               End
               Begin VB.TextBox Text3 
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
                  Left            =   2280
                  TabIndex        =   253
                  Text            =   "25"
                  Top             =   3480
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   6
                  Left            =   2160
                  TabIndex        =   251
                  Top             =   2880
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Ԫ�ؾ�ͨ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   250
                  Top             =   3120
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "����Ч��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   4
                  Left            =   1080
                  TabIndex        =   249
                  Top             =   2880
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "�����˺�"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   248
                  Top             =   2880
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   247
                  Top             =   2640
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   246
                  Top             =   2640
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����ֵ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox4 
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   245
                  Top             =   2640
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   6
                  Left            =   1800
                  TabIndex        =   244
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Ԫ�ؾ�ͨ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   5
                  Left            =   720
                  TabIndex        =   243
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "���Ƽӳ�"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   4
                  Left            =   1560
                  TabIndex        =   242
                  Top             =   1800
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�����˺�"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   3
                  Left            =   720
                  TabIndex        =   241
                  Top             =   1800
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   240
                  Top             =   1560
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   239
                  Top             =   1560
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����ֵ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox3 
                  Height          =   255
                  Index           =   0
                  Left            =   720
                  TabIndex        =   238
                  Top             =   1560
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox2 
                  Height          =   255
                  Index           =   4
                  Left            =   1560
                  TabIndex        =   237
                  Top             =   1200
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Ԫ�ؾ�ͨ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox2 
                  Height          =   255
                  Index           =   3
                  Left            =   720
                  TabIndex        =   236
                  Top             =   1200
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "����"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox2 
                  Height          =   255
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   235
                  Top             =   960
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox2 
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   234
                  Top             =   960
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����ֵ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox2 
                  Height          =   255
                  Index           =   0
                  Left            =   720
                  TabIndex        =   233
                  Top             =   960
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox1 
                  Height          =   255
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   230
                  Top             =   360
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox1 
                  Height          =   255
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   228
                  Top             =   360
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����ֵ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox1 
                  Height          =   255
                  Index           =   4
                  Left            =   720
                  TabIndex        =   232
                  Top             =   600
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����Ч��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox1 
                  Height          =   255
                  Index           =   3
                  Left            =   1800
                  TabIndex        =   231
                  Top             =   600
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Ԫ�ؾ�ͨ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckBox1 
                  Height          =   255
                  Index           =   1
                  Left            =   720
                  TabIndex        =   229
                  Top             =   360
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "������"
                  BackColor       =   16777215
               End
               Begin VB.Shape Shape5 
                  Height          =   4215
                  Left            =   0
                  Top             =   0
                  Width           =   3375
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
                  Left            =   2880
                  TabIndex        =   255
                  Top             =   0
                  Width           =   375
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��װЧ��"
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   254
                  Top             =   3840
                  Width           =   720
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��������������40���ڣ���"
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   252
                  Top             =   3480
                  Width           =   2190
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��Ҫ��������������"
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   227
                  Top             =   120
                  Width           =   1620
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ñ�ӣ�"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   226
                  Top             =   1560
                  Width           =   540
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���ӣ�"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   225
                  Top             =   960
                  Width           =   540
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ɳ©��"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   224
                  Top             =   360
                  Width           =   540
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "��Ҫ�����ĸ�������"
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   0
                  Left            =   120
                  TabIndex        =   223
                  Top             =   2400
                  Width           =   1935
               End
            End
            Begin BSkin.CommandButton CommandButton7 
               Height          =   615
               Left            =   8400
               TabIndex        =   196
               Top             =   4080
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   1085
               Text            =   "�����ҵʥ����"
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
            Begin BSkin.CommandButton CommandButton6 
               Height          =   615
               Left            =   8400
               TabIndex        =   194
               Top             =   4920
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   1085
               Text            =   "�����˺�"
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
            Begin BSkin.Container SelectBuffBox 
               Height          =   855
               Index           =   0
               Left            =   240
               TabIndex        =   185
               Top             =   480
               Visible         =   0   'False
               Width           =   8295
               _ExtentX        =   14631
               _ExtentY        =   1508
               BackColor       =   16244694
               Begin BSkin.CheckBox BuffCheck2 
                  Height          =   255
                  Index           =   0
                  Left            =   600
                  TabIndex        =   190
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2775
                  _ExtentX        =   4895
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "CheckBox1"
                  BackColor       =   16244694
               End
               Begin BSkin.CheckBox BuffCheck 
                  Height          =   255
                  Index           =   0
                  Left            =   600
                  TabIndex        =   189
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2775
                  _ExtentX        =   4895
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   9
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "CheckBox1"
                  BackColor       =   16244694
               End
               Begin VB.PictureBox SelectBuffBar 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  FillColor       =   &H00C56A31&
                  ForeColor       =   &H00C56A31&
                  Height          =   255
                  Index           =   0
                  Left            =   600
                  LinkTimeout     =   7
                  ScaleHeight     =   17
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   201
                  TabIndex        =   187
                  Top             =   360
                  Width           =   3015
               End
               Begin VB.Label BuffLabel 
                  BackStyle       =   0  'Transparent
                  Caption         =   "0/7 ��Ч��"
                  Height          =   375
                  Index           =   0
                  Left            =   3840
                  TabIndex        =   188
                  Tag             =   "0"
                  Top             =   360
                  Width           =   3975
               End
               Begin VB.Label SelectBuffLabel 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���𰣺���е���ʱ��߹����������7�㣬����ʱ�������"
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
                  Index           =   0
                  Left            =   120
                  TabIndex        =   186
                  Top             =   0
                  Width           =   6375
               End
            End
         End
      End
      Begin BSkin.Container Container1 
         Height          =   5895
         Index           =   2
         Left            =   0
         TabIndex        =   70
         Top             =   2000
         Visible         =   0   'False
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10398
         BackColor       =   16777152
         Begin BSkin.ScrollBar ScrollBar2 
            Height          =   6255
            Left            =   10320
            TabIndex        =   75
            Top             =   0
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   11033
            Max             =   100
            Speed           =   1
         End
         Begin BSkin.Container ContainerBox 
            Height          =   11205
            Left            =   0
            TabIndex        =   71
            Top             =   -2040
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   19764
            Begin BSkin.Container Container5 
               Height          =   255
               Left            =   1200
               TabIndex        =   218
               Top             =   3480
               Visible         =   0   'False
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               Begin VB.Shape Shape4 
                  Height          =   255
                  Left            =   0
                  Top             =   0
                  Width           =   735
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Label3"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   219
                  Top             =   0
                  Width           =   570
               End
            End
            Begin BSkin.Frame Frame3 
               Height          =   2895
               Left            =   0
               TabIndex        =   85
               Top             =   480
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   5106
               Orientation     =   2
               BackColor       =   16244694
               ColorGradient2  =   16241606
               Caption         =   "�������"
               Icon            =   "FrmMain.frx":0028
               ForeColor       =   -2147483630
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   43
                  Left            =   9480
                  TabIndex        =   214
                  ToolTipText     =   "-40%������"
                  Top             =   1800
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   10
                  Left            =   8160
                  TabIndex        =   101
                  ToolTipText     =   "-20%�ҿ��ԣ�+15%���ˣ�����������Ч����ѡĬ�ϴ���"
                  Top             =   1800
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "˫�ҹ���"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   11
                  Left            =   5280
                  TabIndex        =   112
                  ToolTipText     =   "-23%����"
                  Top             =   2280
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����2��"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   1
                  Left            =   3960
                  TabIndex        =   109
                  Text            =   "0"
                  Top             =   2280
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   12
                  Left            =   6600
                  TabIndex        =   108
                  ToolTipText     =   "-15%����"
                  Top             =   2280
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��ɯ�츳2"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   13
                  Left            =   8280
                  TabIndex        =   107
                  ToolTipText     =   "-15%����"
                  Top             =   2280
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����4��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   7
                  Left            =   8640
                  TabIndex        =   92
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   6
                  Left            =   7800
                  TabIndex        =   93
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   5
                  Left            =   6960
                  TabIndex        =   94
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   4
                  Left            =   6120
                  TabIndex        =   95
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   3
                  Left            =   5280
                  TabIndex        =   96
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   2
                  Left            =   4440
                  TabIndex        =   97
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "ˮ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   1
                  Left            =   3600
                  TabIndex        =   98
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   0
                  Left            =   4800
                  TabIndex        =   104
                  Text            =   "0"
                  Top             =   1800
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   8
                  Left            =   5760
                  TabIndex        =   103
                  ToolTipText     =   "-40%�����޻�ˮ���ס���"
                  Top             =   1800
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����4"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   9
                  Left            =   6720
                  TabIndex        =   102
                  ToolTipText     =   "-20%ȫ���Կ���"
                  Top             =   1800
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "���뻤��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   0
                  Left            =   2640
                  TabIndex        =   99
                  Top             =   1260
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   1
                  Caption         =   "��"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   17
                  Left            =   7680
                  TabIndex        =   91
                  Text            =   "1"
                  Top             =   540
                  Width           =   975
               End
               Begin BSkin.ComboBox BuffComboBox2 
                  Height          =   495
                  Left            =   4920
                  TabIndex        =   89
                  Top             =   480
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   873
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
                  Text            =   "ComboBox1"
               End
               Begin BSkin.ComboBox BuffComboBox1 
                  Height          =   495
                  Left            =   1440
                  TabIndex        =   86
                  Top             =   480
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   873
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
                  Text            =   "ComboBox1"
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���˱���������0%"
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
                  Index           =   5
                  Left            =   360
                  TabIndex        =   111
                  Tag             =   "0"
                  Top             =   2280
                  Width           =   2025
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   6
                  Left            =   3000
                  TabIndex        =   110
                  Top             =   2280
                  Width           =   960
               End
               Begin VB.Label Label2 
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
                  Index           =   3
                  Left            =   360
                  TabIndex        =   106
                  Tag             =   "10"
                  Top             =   1800
                  Width           =   2880
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��������"
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
                  Index           =   4
                  Left            =   3480
                  TabIndex        =   105
                  Top             =   1800
                  Width           =   1440
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "��������Ԫ�ظ��ţ�"
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
                  Index           =   0
                  Left            =   360
                  TabIndex        =   100
                  Top             =   1200
                  Width           =   2295
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�ȼ���"
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
                  Index           =   33
                  Left            =   6960
                  TabIndex        =   90
                  Top             =   540
                  Width           =   720
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���ԣ�"
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
                  Index           =   32
                  Left            =   4200
                  TabIndex        =   88
                  Top             =   540
                  Width           =   720
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ѡ��Ŀ�꣺"
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
                  Index           =   31
                  Left            =   240
                  TabIndex        =   87
                  Top             =   540
                  Width           =   1200
               End
            End
            Begin BSkin.Frame Frame5 
               Height          =   4095
               Left            =   0
               TabIndex        =   159
               Top             =   7440
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   7223
               Orientation     =   2
               BackColor       =   16244694
               ColorGradient2  =   16241606
               Caption         =   "����"
               Icon            =   "FrmMain.frx":0044
               ForeColor       =   -2147483630
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   24
                  Left            =   4320
                  TabIndex        =   217
                  Text            =   "5"
                  Top             =   840
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   23
                  Left            =   7080
                  TabIndex        =   212
                  Text            =   "0"
                  Top             =   3720
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   22
                  Left            =   3120
                  TabIndex        =   210
                  Text            =   "0"
                  Top             =   3720
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   21
                  Left            =   6840
                  TabIndex        =   209
                  Text            =   "0"
                  Top             =   3240
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   20
                  Left            =   2880
                  TabIndex        =   206
                  Text            =   "0"
                  Top             =   3240
                  Width           =   375
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   41
                  Left            =   2880
                  TabIndex        =   204
                  ToolTipText     =   "����Ч����ͻʱ��ȡ�����ȼ���ߵ�Ч��"
                  Top             =   2760
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��������֮��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   40
                  Left            =   2880
                  TabIndex        =   203
                  ToolTipText     =   "����Ч����ͻʱ��ȡ�����ȼ���ߵ�Ч��"
                  Top             =   2280
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��������֮��"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   19
                  Left            =   2400
                  TabIndex        =   202
                  Text            =   "1"
                  Top             =   2760
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   18
                  Left            =   2400
                  TabIndex        =   201
                  Text            =   "1"
                  Top             =   2280
                  Width           =   375
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   17
                  Left            =   2880
                  TabIndex        =   197
                  ToolTipText     =   "����Ч����ͻʱ��ȡ�����ȼ���ߵ�Ч��"
                  Top             =   1800
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��������֮��"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   9
                  Left            =   2160
                  TabIndex        =   168
                  Text            =   "5"
                  Top             =   840
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   10
                  Left            =   2400
                  TabIndex        =   167
                  Text            =   "1"
                  Top             =   1800
                  Width           =   375
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   11
                  Left            =   6360
                  TabIndex        =   166
                  Text            =   "1000"
                  Top             =   840
                  Width           =   615
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   12
                  Left            =   1560
                  TabIndex        =   165
                  Text            =   "10"
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   13
                  Left            =   3960
                  TabIndex        =   164
                  Text            =   "10"
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   14
                  Left            =   8880
                  TabIndex        =   163
                  Text            =   "1000"
                  Top             =   840
                  Width           =   615
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   15
                  Left            =   6480
                  TabIndex        =   162
                  Text            =   "100"
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   16
                  Left            =   9000
                  TabIndex        =   161
                  Text            =   "200"
                  Top             =   1320
                  Width           =   615
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   39
                  Left            =   240
                  TabIndex        =   160
                  ToolTipText     =   "������Ӧ��ɵ��˺����15%"
                  Top             =   480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Ī��1����Ч"
                  BackColor       =   16777215
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ɫ����ħ��"
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
                  Index           =   41
                  Left            =   2400
                  TabIndex        =   220
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1440
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�׳�֮��������"
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
                  Index           =   40
                  Left            =   2760
                  TabIndex        =   216
                  Top             =   840
                  Width           =   1680
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ�������ֵ�̶���ֵ�ӳɣ�"
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
                  Index           =   39
                  Left            =   3960
                  TabIndex        =   211
                  Top             =   3720
                  Width           =   3120
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ�������ֵ�ٷֱȼӳɣ�"
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
                  Index           =   38
                  Left            =   240
                  TabIndex        =   208
                  Top             =   3720
                  Width           =   2880
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ�������̶���ֵ�ӳɣ�"
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
                  Index           =   37
                  Left            =   3960
                  TabIndex        =   207
                  Top             =   3240
                  Width           =   2880
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ�������ٷֱȼӳɣ�"
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
                  Index           =   36
                  Left            =   240
                  TabIndex        =   205
                  Top             =   3240
                  Width           =   2640
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ĩ�֮̾ʫ������"
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
                  Index           =   35
                  Left            =   240
                  TabIndex        =   200
                  Top             =   2760
                  Width           =   2160
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��������֮ʱ������"
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
                  Index           =   34
                  Left            =   240
                  TabIndex        =   199
                  Top             =   2280
                  Width           =   2160
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Թ�����֮�ľ�����"
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   22
                  Left            =   -3120
                  TabIndex        =   198
                  Top             =   -600
                  Width           =   2520
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����Ӣ��̷������"
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
                  Index           =   23
                  Left            =   240
                  TabIndex        =   176
                  Top             =   840
                  Width           =   1920
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Թ�����֮�ľ�����"
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
                  Index           =   24
                  Left            =   240
                  TabIndex        =   175
                  Top             =   1800
                  Width           =   2160
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��Ҷ��ͨ��"
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
                  Index           =   25
                  Left            =   5160
                  TabIndex        =   174
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ī��Q�ȼ���"
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
                  Index           =   26
                  Left            =   240
                  TabIndex        =   173
                  Top             =   1320
                  Width           =   1395
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�׵罫��E�ȼ���"
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
                  Index           =   27
                  Left            =   2160
                  TabIndex        =   172
                  Top             =   1320
                  Width           =   1815
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ɰ�Ǿ�ͨ��"
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
                  Index           =   28
                  Left            =   7680
                  TabIndex        =   171
                  Top             =   840
                  Width           =   1200
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ɯ���Ǳ����ʣ�"
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
                  Index           =   29
                  Left            =   4560
                  TabIndex        =   170
                  Top             =   1320
                  Width           =   1920
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��������Ч�ʣ�"
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
                  Index           =   30
                  Left            =   7320
                  TabIndex        =   169
                  Top             =   1320
                  Width           =   1680
               End
            End
            Begin BSkin.Frame Frame4 
               Height          =   4095
               Left            =   0
               TabIndex        =   113
               Top             =   3360
               Width           =   10335
               _ExtentX        =   18230
               _ExtentY        =   7223
               Orientation     =   2
               BackColor       =   16244694
               ColorGradient2  =   16241606
               Caption         =   "Buff�б�"
               Icon            =   "FrmMain.frx":0060
               ForeColor       =   -2147483630
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   44
                  Left            =   3720
                  TabIndex        =   215
                  ToolTipText     =   "+10%~20%Ԫ���˺��������ȼ���[����]һ������"
                  Top             =   1800
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�׳�֮��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   42
                  Left            =   9240
                  TabIndex        =   213
                  ToolTipText     =   "+120��ͨ"
                  Top             =   2160
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�̹�4"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   28
                  Left            =   7920
                  TabIndex        =   136
                  ToolTipText     =   "+125��ͨ"
                  Top             =   2160
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������Q"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   27
                  Left            =   6480
                  TabIndex        =   137
                  ToolTipText     =   "+200��ͨ"
                  Top             =   2160
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�ϰ���6��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   26
                  Left            =   5160
                  TabIndex        =   138
                  ToolTipText     =   "+200��ͨ"
                  Top             =   2160
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��Ҷ2��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   38
                  Left            =   8880
                  TabIndex        =   143
                  ToolTipText     =   "+15%�����ʣ��������ڱ���������Ч"
                  Top             =   2640
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "˫������"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   8
                  Left            =   3960
                  TabIndex        =   158
                  Text            =   "0"
                  Top             =   3600
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   33
                  Left            =   5040
                  TabIndex        =   155
                  ToolTipText     =   "������������[����]һ������"
                  Top             =   3600
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�����츳"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   7
                  Left            =   3720
                  TabIndex        =   152
                  Text            =   "0"
                  Top             =   3120
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   32
                  Left            =   4800
                  TabIndex        =   151
                  ToolTipText     =   "+60%�����˺�������������"
                  Top             =   3120
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����6��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   34
                  Left            =   6240
                  TabIndex        =   150
                  ToolTipText     =   "+20%�����ʣ�+20%�����˺�"
                  Top             =   3120
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�ɵ�����/û��δ����"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   31
                  Left            =   7680
                  TabIndex        =   144
                  ToolTipText     =   "+12%������"
                  Top             =   2640
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����4��"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   30
                  Left            =   5880
                  TabIndex        =   149
                  ToolTipText     =   "��ɯ���Ǳ�������[����]һ������"
                  Top             =   2640
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��ɯ�����츳"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   6
                  Left            =   3480
                  TabIndex        =   146
                  Text            =   "0"
                  Top             =   2640
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   29
                  Left            =   4560
                  TabIndex        =   145
                  ToolTipText     =   "+12%������"
                  Top             =   2640
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�����츳"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   25
                  Left            =   3840
                  TabIndex        =   140
                  ToolTipText     =   "ɰ����徫ͨ��[����]һ������"
                  Top             =   2160
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "ɰ���츳"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   5
                  Left            =   3360
                  TabIndex        =   139
                  Text            =   "0"
                  Top             =   2160
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   36
                  Left            =   5040
                  TabIndex        =   127
                  ToolTipText     =   "+25%��ӦԪ���˺����������"
                  Top             =   1800
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   24
                  Left            =   8400
                  TabIndex        =   128
                  ToolTipText     =   "+20%�������˺�"
                  Top             =   1440
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����Q"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   23
                  Left            =   6960
                  TabIndex        =   129
                  ToolTipText     =   "����ʹ��Ԫ�ر���ʱ���׵罫��E�ȼ���[����]һ������"
                  Top             =   1440
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "�׵罫��E"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   22
                  Left            =   6000
                  TabIndex        =   130
                  ToolTipText     =   "Ī��Q�ȼ���[����]һ������"
                  Top             =   1440
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Ī��Q"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   20
                  Left            =   3720
                  TabIndex        =   131
                  ToolTipText     =   "�����ѡ����Ҷ2������Ĭ�϶��8�����ˣ���Ҷ��徫ͨ��[����]һ������"
                  Top             =   1440
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��Ҷ�츳"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   4
                  Left            =   3240
                  TabIndex        =   133
                  Text            =   "0"
                  Top             =   1440
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   21
                  Left            =   5040
                  TabIndex        =   132
                  ToolTipText     =   "+35%�����޻�ˮ���ס���"
                  Top             =   1440
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����4"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   18
                  Left            =   5160
                  TabIndex        =   124
                  ToolTipText     =   "�����ص���ϸ��Ϣ���л���ɫ�����"
                  Top             =   960
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "������Q"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   3
                  Left            =   3960
                  TabIndex        =   123
                  Text            =   "0"
                  Top             =   960
                  Width           =   975
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   19
                  Left            =   6480
                  TabIndex        =   122
                  ToolTipText     =   "��������ϸ��Ϣ���л���ɫ�����"
                  Top             =   960
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��������"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   35
                  Left            =   7920
                  TabIndex        =   121
                  ToolTipText     =   "+372��������+12%������"
                  Top             =   960
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "��ζ������ǽ"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   37
                  Left            =   7920
                  TabIndex        =   114
                  ToolTipText     =   "+25%������"
                  Top             =   480
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "˫����"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   16
                  Left            =   6960
                  TabIndex        =   115
                  ToolTipText     =   "+24-48%�������������ȼ���[����]һ������"
                  Top             =   480
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   15
                  Left            =   5880
                  TabIndex        =   116
                  ToolTipText     =   "+20%������"
                  Top             =   480
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "ǧ��4"
                  BackColor       =   16777215
               End
               Begin BSkin.CheckBox CheckState 
                  Height          =   255
                  Index           =   14
                  Left            =   4920
                  TabIndex        =   117
                  ToolTipText     =   "+20%������"
                  Top             =   480
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "����4"
                  BackColor       =   16777215
               End
               Begin VB.TextBox Text1 
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "΢���ź�"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   2
                  Left            =   4080
                  TabIndex        =   120
                  Text            =   "0"
                  Top             =   480
                  Width           =   975
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ԫ�س���Ч�ʼӳɣ�0%"
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
                  Index           =   20
                  Left            =   240
                  TabIndex        =   157
                  Top             =   3600
                  Width           =   2505
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   21
                  Left            =   3000
                  TabIndex        =   156
                  Top             =   3600
                  Width           =   960
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�����˺��ӳɣ�0%"
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
                  Index           =   18
                  Left            =   240
                  TabIndex        =   154
                  Top             =   3120
                  Width           =   2025
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   19
                  Left            =   2760
                  TabIndex        =   153
                  Top             =   3120
                  Width           =   960
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�����ʼӳɣ�0%"
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
                  Index           =   16
                  Left            =   240
                  TabIndex        =   148
                  Top             =   2640
                  Width           =   1785
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   17
                  Left            =   2520
                  TabIndex        =   147
                  Top             =   2640
                  Width           =   960
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ԫ�ؾ�ͨ�ӳɣ�0"
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
                  Index           =   14
                  Left            =   240
                  TabIndex        =   142
                  Top             =   2160
                  Width           =   1815
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   15
                  Left            =   2400
                  TabIndex        =   141
                  Top             =   2160
                  Width           =   960
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�˺��ӳɣ�0%"
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
                  Index           =   12
                  Left            =   240
                  TabIndex        =   135
                  Top             =   1440
                  Width           =   1545
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   13
                  Left            =   2280
                  TabIndex        =   134
                  Top             =   1440
                  Width           =   960
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���������ּӳɣ�0"
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
                  Index           =   10
                  Left            =   240
                  TabIndex        =   126
                  Tag             =   "0"
                  Top             =   960
                  Width           =   2055
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   11
                  Left            =   3000
                  TabIndex        =   125
                  Top             =   960
                  Width           =   960
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�������ٷֱȼӳɣ�0%"
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
                  Index           =   8
                  Left            =   240
                  TabIndex        =   119
                  Tag             =   "0"
                  Top             =   480
                  Width           =   2505
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Զ��壺"
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
                  Index           =   9
                  Left            =   3240
                  TabIndex        =   118
                  Top             =   480
                  Width           =   960
               End
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " "
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
               Index           =   7
               Left            =   240
               TabIndex        =   74
               Top             =   2400
               Width           =   75
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "�������˺�"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   2760
               TabIndex        =   73
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ǰ���ܣ�"
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
               Index           =   1
               Left            =   240
               TabIndex        =   72
               Top             =   120
               Width           =   1200
            End
         End
      End
      Begin BSkin.Container Container1 
         Height          =   6015
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10610
         BackColor       =   16761024
         Begin BSkin.Timer zTimCtn1 
            Left            =   240
            Top             =   1440
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin BSkin.ListBox ListBox1 
            Height          =   2775
            Left            =   360
            TabIndex        =   8
            Top             =   3120
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4895
            SelectBackColor =   16744576
            SelectForeColor =   4342338
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
         Begin BSkin.Frame Frame1 
            Height          =   6015
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   10610
            Orientation     =   2
            BackColor       =   16244694
            ColorGradient2  =   16241606
            ShowIcon        =   0   'False
            Caption         =   "��ɫ������������"
            Icon            =   "FrmMain.frx":007C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin BSkin.CheckBox RBox 
               Height          =   495
               Index           =   4
               Left            =   10080
               TabIndex        =   30
               Top             =   1725
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   ""
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox RBox 
               Height          =   495
               Index           =   3
               Left            =   9600
               TabIndex        =   29
               Top             =   1725
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox RBox 
               Height          =   495
               Index           =   2
               Left            =   9120
               TabIndex        =   28
               Top             =   1725
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox RBox 
               Height          =   495
               Index           =   1
               Left            =   8640
               TabIndex        =   27
               Top             =   1725
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox RBox 
               Height          =   495
               Index           =   0
               Left            =   8160
               TabIndex        =   26
               Top             =   1725
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               Enabled         =   0   'False
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.ComboBox WeaponBox 
               Height          =   300
               Left            =   8640
               TabIndex        =   25
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               Alignment       =   2
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
               Text            =   "ComboBox1"
            End
            Begin BSkin.ComboBox LevelBox 
               Height          =   300
               Index           =   2
               Left            =   4440
               TabIndex        =   23
               Top             =   2295
               Width           =   960
               _ExtentX        =   1693
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
               Text            =   ""
            End
            Begin BSkin.ComboBox LevelBox 
               Height          =   300
               Index           =   1
               Left            =   3000
               TabIndex        =   22
               Top             =   2280
               Width           =   960
               _ExtentX        =   1693
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
               Text            =   ""
            End
            Begin BSkin.ComboBox LevelBox 
               Height          =   300
               Index           =   0
               Left            =   1560
               TabIndex        =   20
               Top             =   2280
               Width           =   960
               _ExtentX        =   1693
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
               Text            =   ""
            End
            Begin BSkin.CheckBox CBox 
               Height          =   495
               Index           =   5
               Left            =   5400
               TabIndex        =   18
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   ""
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox CBox 
               Height          =   495
               Index           =   4
               Left            =   4920
               TabIndex        =   17
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox CBox 
               Height          =   495
               Index           =   3
               Left            =   4440
               TabIndex        =   16
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox CBox 
               Height          =   495
               Index           =   2
               Left            =   3960
               TabIndex        =   15
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox CBox 
               Height          =   495
               Index           =   1
               Left            =   3480
               TabIndex        =   14
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin BSkin.CheckBox CBox 
               Height          =   495
               Index           =   0
               Left            =   3000
               TabIndex        =   13
               Top             =   1680
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "��"
               BackColor       =   16244694
               ForeColor       =   16744576
            End
            Begin VB.PictureBox LevelBar 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               FillColor       =   &H00C56A31&
               ForeColor       =   &H00C56A31&
               Height          =   255
               Left            =   2280
               ScaleHeight     =   17
               ScaleMode       =   0  'User
               ScaleWidth      =   192.344
               TabIndex        =   12
               Top             =   1350
               Width           =   3008
            End
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѡ���ܣ�"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   13
               Left            =   480
               TabIndex        =   195
               Tag             =   "90"
               Top             =   2760
               Width           =   1200
            End
            Begin VB.Label lblTab 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00000000&
               Height          =   735
               Index           =   11
               Left            =   6120
               TabIndex        =   34
               Tag             =   "90"
               Top             =   2280
               Width           =   4020
            End
            Begin VB.Label lblTab 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   10
               Left            =   3420
               TabIndex        =   33
               Tag             =   "90"
               Top             =   600
               Width           =   720
            End
            Begin VB.Label lblTab 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ҽ���"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   9
               Left            =   7440
               TabIndex        =   32
               Tag             =   "90"
               Top             =   600
               Width           =   2760
            End
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   8
               Left            =   7440
               TabIndex        =   31
               Tag             =   "90"
               Top             =   1800
               Width           =   720
            End
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����ȼ���"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   7
               Left            =   7440
               TabIndex        =   24
               Tag             =   "90"
               Top             =   1200
               Width           =   1200
            End
            Begin BSkin.AlphaImage AlphaImageWeap 
               Height          =   1125
               Left            =   6120
               Top             =   600
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   1984
               Image           =   "FrmMain.frx":02F6
               Props           =   5
            End
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�츳��A                 E                 Q"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   480
               TabIndex        =   21
               Tag             =   "90"
               Top             =   2280
               Width           =   3765
            End
            Begin VB.Label lblTab 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   2280
               TabIndex        =   19
               Tag             =   "90"
               Top             =   1740
               Width           =   720
            End
            Begin VB.Label lblTab 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ɫ�ȼ���90��"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   11
               Tag             =   "90"
               Top             =   1005
               Width           =   3015
            End
            Begin BSkin.AlphaImage AlphaImageChar 
               Height          =   1590
               Left            =   360
               Top             =   480
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   2805
               Image           =   "FrmMain.frx":24F1
               Props           =   5
            End
         End
      End
      Begin VB.PictureBox ImageTemp2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   310
         Index           =   0
         Left            =   0
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   221
         Top             =   0
         Visible         =   0   'False
         Width           =   310
      End
      Begin BSkin.Container Container1 
         Height          =   5895
         Index           =   4
         Left            =   -840
         TabIndex        =   191
         Top             =   2040
         Visible         =   0   'False
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   10398
         BackColor       =   16761087
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            Height          =   5895
            Index           =   0
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   192
            Top             =   0
            Visible         =   0   'False
            Width           =   3375
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5760
         Tag             =   "0"
         Top             =   360
      End
      Begin BSkin.Container Container1 
         Height          =   6015
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10610
         BackColor       =   12632319
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   5640
            ScaleHeight     =   735
            ScaleWidth      =   4575
            TabIndex        =   181
            Top             =   120
            Visible         =   0   'False
            Width           =   4575
            Begin BSkin.CommandButton CommandButton4 
               Height          =   495
               Left            =   2640
               TabIndex        =   182
               Top             =   120
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   873
               Text            =   "�ر�"
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
            Begin BSkin.CommandButton CommandButton5 
               Height          =   495
               Left            =   600
               TabIndex        =   183
               Top             =   120
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   873
               Text            =   "����ʥ����"
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
         End
         Begin BSkin.ScrollBar ScrollBar1 
            Height          =   4935
            Left            =   10320
            TabIndex        =   39
            Top             =   960
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   8705
            Max             =   100
            Speed           =   1
         End
         Begin BSkin.CommandButton CommandButton2 
            Height          =   495
            Left            =   2280
            TabIndex        =   38
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Text            =   "�鿴ʥ������"
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
         Begin BSkin.Container Container2 
            Height          =   5055
            Left            =   0
            TabIndex        =   36
            Top             =   960
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   8916
            Begin BSkin.Frame Frame2 
               Height          =   5055
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   10695
               _ExtentX        =   18865
               _ExtentY        =   8916
               Orientation     =   2
               BackColor       =   16244694
               ColorGradient2  =   16241606
               Caption         =   "ʥ���﷽��"
               Icon            =   "FrmMain.frx":8877
               ForeColor       =   -2147483630
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "΢���ź�"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin BSkin.Container ShowBox 
                  Height          =   2055
                  Left            =   360
                  TabIndex        =   44
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   2775
                  _ExtentX        =   4895
                  _ExtentY        =   3625
                  Begin VB.Image Image1 
                     Height          =   240
                     Left            =   960
                     Picture         =   "FrmMain.frx":8893
                     Top             =   600
                     Width           =   900
                  End
                  Begin VB.Label ShowBoxLabel 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "��Ԫ���˺��ӳ�+46.6%"
                     BeginProperty Font 
                        Name            =   "΢���ź�"
                        Size            =   12
                        Charset         =   134
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   615
                     Index           =   0
                     Left            =   120
                     TabIndex        =   179
                     Top             =   300
                     Width           =   2655
                  End
                  Begin VB.Label ShowBoxLabel 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "ħŮ-��֮��"
                     BeginProperty Font 
                        Name            =   "΢���ź�"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   5
                     Left            =   120
                     TabIndex        =   178
                     Top             =   0
                     Width           =   2655
                  End
                  Begin VB.Shape Shape2 
                     Height          =   2055
                     Left            =   0
                     Top             =   0
                     Width           =   2775
                  End
                  Begin VB.Shape Shape1 
                     Height          =   2055
                     Left            =   3120
                     Top             =   -3960
                     Width           =   1815
                  End
                  Begin VB.Label ShowBoxLabel 
                     BackStyle       =   0  'Transparent
                     Caption         =   "������+5"
                     BeginProperty Font 
                        Name            =   "΢���ź�"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   4
                     Left            =   240
                     TabIndex        =   48
                     Top             =   1560
                     Width           =   2535
                  End
                  Begin VB.Label ShowBoxLabel 
                     BackStyle       =   0  'Transparent
                     Caption         =   "������+5"
                     BeginProperty Font 
                        Name            =   "΢���ź�"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   3
                     Left            =   240
                     TabIndex        =   47
                     Top             =   1320
                     Width           =   2535
                  End
                  Begin VB.Label ShowBoxLabel 
                     BackStyle       =   0  'Transparent
                     Caption         =   "������+5"
                     BeginProperty Font 
                        Name            =   "΢���ź�"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   2
                     Left            =   240
                     TabIndex        =   46
                     Top             =   1080
                     Width           =   2415
                  End
                  Begin VB.Label ShowBoxLabel 
                     BackStyle       =   0  'Transparent
                     Caption         =   "������+5"
                     BeginProperty Font 
                        Name            =   "΢���ź�"
                        Size            =   10.5
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   1
                     Left            =   240
                     TabIndex        =   45
                     Top             =   840
                     Width           =   2415
                  End
               End
               Begin BSkin.Container Container4 
                  Height          =   4575
                  Left            =   0
                  TabIndex        =   180
                  Tag             =   "0"
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   10335
                  _ExtentX        =   18230
                  _ExtentY        =   8070
                  Begin BSkin.AlphaImage ArtShowImage 
                     Height          =   1050
                     Index           =   0
                     Left            =   360
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   1852
                     Image           =   "FrmMain.frx":8E20
                     Props           =   5
                  End
               End
               Begin BSkin.Container SetBox 
                  Height          =   2055
                  Index           =   0
                  Left            =   480
                  TabIndex        =   40
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   9255
                  _ExtentX        =   16325
                  _ExtentY        =   3625
                  Begin BSkin.CommandButton SetCopyButton 
                     Height          =   375
                     Index           =   0
                     Left            =   4920
                     TabIndex        =   184
                     Top             =   120
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   661
                     Text            =   "����"
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
                  Begin BSkin.Container SetBox2 
                     Height          =   1455
                     Index           =   0
                     Left            =   120
                     TabIndex        =   49
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   8055
                     _ExtentX        =   14208
                     _ExtentY        =   2566
                     Begin VB.TextBox SetText7 
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
                        Left            =   6960
                        TabIndex        =   69
                        Text            =   "0"
                        Top             =   1080
                        Width           =   975
                     End
                     Begin VB.TextBox SetText6 
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
                        Left            =   6960
                        TabIndex        =   68
                        Text            =   "0"
                        Top             =   720
                        Width           =   975
                     End
                     Begin VB.TextBox SetText5 
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
                        Left            =   6960
                        TabIndex        =   67
                        Text            =   "0"
                        Top             =   360
                        Width           =   975
                     End
                     Begin VB.TextBox SetText4 
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
                        Left            =   6960
                        TabIndex        =   66
                        Text            =   "0"
                        Top             =   0
                        Width           =   975
                     End
                     Begin VB.TextBox SetText3 
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
                        Left            =   4320
                        TabIndex        =   65
                        Text            =   "0"
                        Top             =   840
                        Width           =   975
                     End
                     Begin VB.TextBox SetText2 
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
                        Left            =   4320
                        TabIndex        =   64
                        Text            =   "0"
                        Top             =   480
                        Width           =   975
                     End
                     Begin VB.TextBox SetText1 
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
                        Left            =   4320
                        TabIndex        =   57
                        Text            =   "0"
                        Top             =   120
                        Width           =   975
                     End
                     Begin BSkin.ComboBox SetCombo3 
                        Height          =   375
                        Index           =   0
                        Left            =   960
                        TabIndex        =   52
                        Top             =   960
                        Width           =   1815
                        _ExtentX        =   3201
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
                        Text            =   "ComboBox1"
                     End
                     Begin BSkin.ComboBox SetCombo2 
                        Height          =   375
                        Index           =   0
                        Left            =   960
                        TabIndex        =   51
                        Top             =   480
                        Width           =   1815
                        _ExtentX        =   3201
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
                        Text            =   "ComboBox1"
                     End
                     Begin BSkin.ComboBox SetCombo1 
                        Height          =   375
                        Index           =   0
                        Left            =   960
                        TabIndex        =   50
                        Top             =   0
                        Width           =   1815
                        _ExtentX        =   3201
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
                        Text            =   "ComboBox1"
                     End
                     Begin VB.Label SetTipLabel11 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+����Ч��%��"
                        Height          =   375
                        Index           =   0
                        Left            =   5760
                        TabIndex        =   63
                        Top             =   1080
                        Width           =   1335
                     End
                     Begin VB.Label SetTipLabel10 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+�����˺�%��"
                        Height          =   375
                        Index           =   0
                        Left            =   5760
                        TabIndex        =   62
                        Top             =   720
                        Width           =   1335
                     End
                     Begin VB.Label SetTipLabel9 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+������%��"
                        Height          =   375
                        Index           =   0
                        Left            =   5760
                        TabIndex        =   61
                        Top             =   360
                        Width           =   1215
                     End
                     Begin VB.Label SetTipLabel8 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+Ԫ�ؾ�ͨ��"
                        Height          =   375
                        Index           =   0
                        Left            =   5760
                        TabIndex        =   60
                        Top             =   0
                        Width           =   1215
                     End
                     Begin VB.Label SetTipLabel7 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+��������"
                        Height          =   375
                        Index           =   0
                        Left            =   3360
                        TabIndex        =   59
                        Top             =   840
                        Width           =   975
                     End
                     Begin VB.Label SetTipLabel6 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+��������"
                        Height          =   375
                        Index           =   0
                        Left            =   3360
                        TabIndex        =   58
                        Top             =   480
                        Width           =   975
                     End
                     Begin VB.Label SetTipLabel5 
                        BackStyle       =   0  'Transparent
                        Caption         =   "+����ֵ��"
                        Height          =   375
                        Index           =   0
                        Left            =   3360
                        TabIndex        =   56
                        Top             =   120
                        Width           =   975
                     End
                     Begin VB.Label SetTipLabel4 
                        BackStyle       =   0  'Transparent
                        Caption         =   "��֮��"
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        TabIndex        =   55
                        Top             =   1020
                        Width           =   735
                     End
                     Begin VB.Label SetTipLabel3 
                        BackStyle       =   0  'Transparent
                        Caption         =   "��֮��"
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        TabIndex        =   54
                        Top             =   540
                        Width           =   735
                     End
                     Begin VB.Label SetTipLabel2 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ʱ֮ɳ"
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        TabIndex        =   53
                        Top             =   60
                        Width           =   735
                     End
                  End
                  Begin BSkin.Switch SetSwitch 
                     Height          =   375
                     Index           =   0
                     Left            =   3720
                     TabIndex        =   42
                     Top             =   120
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   661
                     skin            =   2
                  End
                  Begin BSkin.AlphaImage SetPic5 
                     Height          =   1050
                     Index           =   0
                     Left            =   5640
                     Top             =   720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   1852
                     Image           =   "FrmMain.frx":B4CB
                     Props           =   5
                  End
                  Begin VB.Label SetTipLabel13 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "ħŮ2"
                     Height          =   735
                     Index           =   0
                     Left            =   8280
                     TabIndex        =   84
                     Top             =   600
                     Width           =   765
                  End
                  Begin VB.Label SetTipLabel12 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "��װЧ��"
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
                     Left            =   8040
                     TabIndex        =   76
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1215
                  End
                  Begin BSkin.AlphaImage SetPic4 
                     Height          =   1050
                     Index           =   0
                     Left            =   4320
                     Top             =   720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   1852
                     Image           =   "FrmMain.frx":D9AA
                     Props           =   5
                  End
                  Begin BSkin.AlphaImage SetPic3 
                     Height          =   1050
                     Index           =   0
                     Left            =   3000
                     Top             =   720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   1852
                     Image           =   "FrmMain.frx":FE89
                     Props           =   5
                  End
                  Begin BSkin.AlphaImage SetPic2 
                     Height          =   1050
                     Index           =   0
                     Left            =   1680
                     Top             =   720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   1852
                     Image           =   "FrmMain.frx":12611
                     Props           =   5
                  End
                  Begin BSkin.AlphaImage SetPic1 
                     Height          =   1050
                     Index           =   0
                     Left            =   360
                     Top             =   720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   1852
                     Image           =   "FrmMain.frx":149BA
                     Props           =   5
                  End
                  Begin VB.Label SetTipLabel 
                     BackStyle       =   0  'Transparent
                     Caption         =   "�����ʥ���ﵥ�����"
                     Height          =   255
                     Index           =   0
                     Left            =   1680
                     TabIndex        =   43
                     Top             =   165
                     Width           =   1815
                  End
                  Begin VB.Label SetLabel 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "����1"
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
                     Index           =   0
                     Left            =   120
                     TabIndex        =   41
                     Top             =   120
                     Width           =   615
                  End
               End
               Begin BSkin.Container Container3 
                  Height          =   1575
                  Left            =   4800
                  TabIndex        =   77
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   2778
                  Begin BSkin.CommandButton CommandButton3 
                     Height          =   375
                     Left            =   2160
                     TabIndex        =   83
                     Top             =   1080
                     Width           =   855
                     _ExtentX        =   1508
                     _ExtentY        =   661
                     Text            =   "ȷ��"
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
                  Begin BSkin.ComboBox SetSelectBox 
                     Height          =   375
                     Index           =   1
                     Left            =   960
                     TabIndex        =   79
                     Top             =   600
                     Width           =   2055
                     _ExtentX        =   3625
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
                     Text            =   ""
                  End
                  Begin BSkin.ComboBox SetSelectBox 
                     Height          =   375
                     Index           =   0
                     Left            =   960
                     TabIndex        =   78
                     Top             =   120
                     Width           =   2055
                     _ExtentX        =   3625
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
                     Text            =   ""
                  End
                  Begin VB.Shape Shape3 
                     Height          =   1575
                     Left            =   0
                     Top             =   0
                     Width           =   3255
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  'Transparent
                     Caption         =   "*�Ѵ����ļ���"
                     ForeColor       =   &H000000FF&
                     Height          =   495
                     Index           =   2
                     Left            =   240
                     TabIndex        =   82
                     Top             =   1080
                     Width           =   1815
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  'Transparent
                     Caption         =   "2����"
                     Height          =   495
                     Index           =   1
                     Left            =   240
                     TabIndex        =   81
                     Top             =   600
                     Width           =   495
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  'Transparent
                     Caption         =   "2����"
                     Height          =   495
                     Index           =   0
                     Left            =   240
                     TabIndex        =   80
                     Top             =   120
                     Width           =   495
                  End
               End
            End
         End
         Begin BSkin.CommandButton CommandButton1 
            Height          =   495
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Text            =   "+ ����һ�׷���"
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
         Begin BSkin.Tray Tray1 
            Left            =   9960
            Top             =   4200
            _ExtentX        =   741
            _ExtentY        =   741
            PictureIcon     =   "FrmMain.frx":17065
         End
      End
      Begin VB.Label lblTab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�鿴���"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   5280
         TabIndex        =   193
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin BSkin.AlphaImage pngTab 
         Height          =   735
         Index           =   4
         Left            =   5280
         Top             =   720
         Visible         =   0   'False
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1296
         Image           =   "FrmMain.frx":18AE7
         Props           =   5
      End
      Begin BSkin.AlphaImage pngMenu 
         Height          =   225
         Left            =   9540
         Top             =   180
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         Image           =   "FrmMain.frx":19027
         Props           =   5
      End
      Begin BSkin.AlphaImage pngPowered 
         Height          =   450
         Left            =   3360
         Top             =   3960
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   794
         Image           =   "FrmMain.frx":19B97
         Props           =   5
      End
      Begin VB.Label lblTab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ч����"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   6
         Top             =   1425
         Width           =   735
      End
      Begin BSkin.AlphaImage pngTab 
         Height          =   735
         Index           =   3
         Left            =   4080
         Top             =   720
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1296
         Image           =   "FrmMain.frx":1B7C5
         Props           =   5
      End
      Begin VB.Label lblTab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buff����"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   1425
         Width           =   735
      End
      Begin BSkin.AlphaImage pngTab 
         Height          =   735
         Index           =   2
         Left            =   2760
         Top             =   660
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1296
         Image           =   "FrmMain.frx":1BD05
         Props           =   5
      End
      Begin VB.Label lblTab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʥ����"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1650
         TabIndex        =   4
         Top             =   1420
         Width           =   555
      End
      Begin BSkin.AlphaImage pngTab 
         Height          =   885
         Index           =   1
         Left            =   1560
         Top             =   630
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1561
         Image           =   "FrmMain.frx":1C245
         Props           =   5
      End
      Begin VB.Label lblTab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ɫ��Ϣ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1420
         Width           =   735
      End
      Begin BSkin.AlphaImage AlphaImage1 
         Height          =   1290
         Left            =   165
         Top             =   480
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   2275
         Image           =   "FrmMain.frx":1C878
         Props           =   5
      End
      Begin BSkin.AlphaImage pngTab 
         Height          =   825
         Index           =   0
         Left            =   360
         Top             =   630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1455
         Image           =   "FrmMain.frx":1D578
         Props           =   5
      End
      Begin BSkin.AlphaImage pngMinimize 
         Height          =   225
         Left            =   10020
         Top             =   180
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         Image           =   "FrmMain.frx":1DEA1
         Props           =   5
      End
      Begin BSkin.AlphaImage pngMinimizeBG 
         Height          =   360
         Left            =   9960
         Top             =   120
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Image           =   "FrmMain.frx":1EA76
         Props           =   5
      End
      Begin BSkin.AlphaImage pngClose 
         Height          =   225
         Left            =   10500
         Top             =   180
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         Image           =   "FrmMain.frx":1EB30
         Props           =   5
      End
      Begin BSkin.AlphaImage pngCloseBG 
         Height          =   360
         Left            =   10440
         Top             =   120
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Image           =   "FrmMain.frx":1F6D1
         Props           =   5
      End
      Begin VB.Label lblLogo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   480
         TabIndex        =   1
         Top             =   105
         Width           =   1200
      End
      Begin BSkin.AlphaImage pngLogo 
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Image           =   "FrmMain.frx":1F789
         Props           =   5
      End
      Begin VB.Image ImageTemp 
         Height          =   1050
         Index           =   0
         Left            =   -360
         Picture         =   "FrmMain.frx":216C2
         Top             =   0
         Visible         =   0   'False
         Width           =   1050
      End
      Begin BSkin.AlphaImage pngMain 
         Height          =   7920
         Left            =   0
         Top             =   0
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   13970
         Image           =   "FrmMain.frx":21C96
         Props           =   5
      End
   End
   Begin VB.Menu mnuApp 
      Caption         =   "�˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "��ʾ����"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "ϵͳ����"
      End
      Begin VB.Menu mnuOther 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�����"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'**
'** @ BL2CK Software Co.Ltd All Rights Reserved
'**
'********************************************************
Option Explicit

'���� ���� - ���� - ��ѡ�� BSkin.ocx

'ϵͳAPI����������������������������������������������������������������������������������������������
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long '�����ޱ߿�
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000

'�������ݡ���������������������������������������������������������������������������������������������
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String '��������

'��̬Ч������������������������������������������������������������������������������������������������
Dim mObj1 As Object '�߳�һ
Dim mToTop1 As Single, mToLeft1 As Single '���һ�ε���ƶ������ؼ�������дһ���߳�

'������Ӱ����������������������������������������������������������������������������������������������
Private FormShadow As clsShadow

'ϵͳ�˵�����������������������������������������������������������������������������������������������
Public Enum CMenuTypeEnum
    MenuString = 0
    MenuSeparate = 1
    CheckBox = 2
End Enum

Private Const WM_TASKMENU As Long = &H313
Private Const WM_SYSCOMMAND As Long = &H112

Private WithEvents C_Menu As clsCMenu
Attribute C_Menu.VB_VarHelpID = -1
Private WithEvents C_Sort As clsCMenu
Attribute C_Sort.VB_VarHelpID = -1
Private WithEvents c_Subclass As clsCSubclass
Attribute c_Subclass.VB_VarHelpID = -1

'ȡɫģ�顪��������������������������������������������������������������������������������������������
Dim Red As Long
Dim Green As Long
Dim Blue As Long
Dim sRed As Long
Dim sGreen As Long
Dim sBlue As Long
Dim Color As Long
Dim Text1ban As Boolean

Private MODEL_TRAY As Boolean '�Ƿ���ʾ�������ݣ���������ģʽ��ÿ�ιرմ��嶼�ᵯ�����ݣ�
Dim LevelText(1 To 96) As String, SetCount As Integer, BoxTemp(0 To 3) As Integer
Dim CurrChar As Integer



Private Sub AlphaImageChar_Click(ByVal Button As Integer)
Unload FrmChar
FrmChar.Show

End Sub

'Public SET_TRAY As String

Private Sub AlphaImageChar_MouseEnter()
AlphaImageChar.FadeInOut 40, 8
End Sub
Private Sub AlphaImageChar_MouseExit()
AlphaImageChar.FadeInOut 100, 8
End Sub

Private Sub AlphaImageWeap_Click(ByVal Button As Integer)
'Unload FrmChar
FrmChar.Show
FrmChar.ShowWeapon Val(CharList(Val(AlphaImageChar.tag), 2))
End Sub

Private Sub AlphaImageWeap_MouseEnter()
AlphaImageWeap.FadeInOut 40, 8
End Sub
Private Sub AlphaImageWeap_MouseExit()
AlphaImageWeap.FadeInOut 100, 8
End Sub


Private Sub ArtShowImage_Click(Index As Integer, ByVal Button As Integer)
Dim i As Integer, j As Integer
i = Val(Picture2.tag)
j = Val(Container4.tag)
If j = 0 Then Exit Sub
    If i = 1 Then
        SetPic1(j).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(ArtShowImage(Index).tag), 1) + ".jpg"
        SetPic1(j).tag = ArtShowImage(Index).tag
    End If
    If i = 2 Then
        SetPic2(j).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(ArtShowImage(Index).tag), 1) + ".jpg"
        SetPic2(j).tag = ArtShowImage(Index).tag
    End If
    If i = 3 Then
        SetPic3(j).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(ArtShowImage(Index).tag), 1) + ".jpg"
        SetPic3(j).tag = ArtShowImage(Index).tag
    End If
    If i = 4 Then
        SetPic4(j).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(ArtShowImage(Index).tag), 1) + ".jpg"
        SetPic4(j).tag = ArtShowImage(Index).tag
    End If
    If i = 5 Then
        SetPic5(j).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(ArtShowImage(Index).tag), 1) + ".jpg"
        SetPic5(j).tag = ArtShowImage(Index).tag
    End If
    Container4.Visible = False
    Container4.tag = 0
    Call SaveSet(j)
    Frame2.Caption = "ʥ���﷽��"
    Picture2.Visible = False
     ReloadTip = True
End Sub

Private Sub SaveSet(Index As Integer)
Dim s As String
            Open App.Path + "\Data\User\C" + AlphaImageChar.tag + "\set" + CStr(Index) For Output As #1
                s = IIf(SetSwitch(Index).Value, "1", "0") + vbCrLf + SetPic1(Index).tag + vbCrLf + SetPic2(Index).tag + vbCrLf + SetPic3(Index).tag + vbCrLf + SetPic4(Index).tag + vbCrLf + SetPic5(Index).tag + vbCrLf + CStr(SetCombo1(Index).ListIndex) + vbCrLf + CStr(SetCombo2(Index).ListIndex) + vbCrLf + CStr(SetCombo3(Index).ListIndex)
                s = s + vbCrLf + SetText1(Index).Text + vbCrLf + SetText2(Index).Text + vbCrLf + SetText3(Index).Text + vbCrLf + SetText4(Index).Text + vbCrLf + SetText5(Index).Text + vbCrLf + SetText6(Index).Text + vbCrLf + SetText7(Index).Text + vbCrLf + SetTipLabel13(Index).Caption
            Print #1, s
            Close #1
End Sub



Private Sub ArtShowImage_MouseEnter(Index As Integer)
ShowBox.Top = ArtShowImage(Index).Top + ArtShowImage(Index).Height + 200
ShowBox.Left = ArtShowImage(Index).Left + 700 + IIf(ArtShowImage(Index).Left = 8280, -2000, 0)
LoadArtShowBox Val(ArtShowImage(Index).tag)
ShowBox.Visible = True
End Sub
Private Sub ArtShowImage_MouseExit(Index As Integer)
ShowBox.Visible = False
End Sub


Private Sub LoadArtShowBox(Index As Integer)
On Error GoTo Outs
Dim s As String, v As Variant, v2 As Variant, i As Integer
If Index <= 0 Then Exit Sub
v = Array("��֮��", "��֮��", "ʱ֮ɳ", "��֮��", "��֮��")
v2 = Array("ȾѪ����ʿ��", "����������Ů", "�԰�֮��", "���ҵ���֮ħŮ", "����֮Ӱ", "�ɹ��һ������", "���˴�ص�����", "��ɵ�����", "ƽϢ���׵�����", "ǧ���ι�", "������;����ʿ", "����֮��", "��Ե֮��ӡ", "�Ƕ�ʿ����Ļ��", "׷��֮ע��", "���׵�ʢŭ", "�ƹŵ�����", "��������֮��")
s = ArtList(Index, 1)
ShowBoxLabel(5).Caption = v2(Val(Mid(s, 2, InStr(1, s, "_") - 2)) - 1) + "-" + v(Val(Mid(s, InStr(1, s, "_") + 1, 1)) - 1)
For i = 0 To 4
    ShowBoxLabel(i).Caption = ""
    s = ArtList(Index, i * 2 + 2)
    If Right(s, 1) = "%" Then
        ShowBoxLabel(i).Caption = "%"
        s = Mid(s, 1, Len(s) - 1)
    End If
    ShowBoxLabel(i).Caption = s + "+" + ArtList(Index, i * 2 + 3) + ShowBoxLabel(i).Caption
Next
Outs:
End Sub



Private Sub BuffCheck_Click(Index As Integer)
Dim s As String
s = SelectBuffLabel(Index).Caption
If InStr(1, s, "�˰�����") Then
    test.��ħ = IIf(BuffCheck(Index).Value = Checked, "��", ""): LoadBuff 10
End If
If InStr(1, s, "��ħ") Then
    test.��ħ = IIf(BuffCheck(Index).Value = Checked, "��", ""): LoadBuff 10
End If
If InStr(1, s, "����ħ") Then
    test.��ħ = IIf(BuffCheck(Index).Value = Checked, "��", ""): LoadBuff 10
End If

End Sub

Private Sub BuffComboBox1_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
Dim t As String, i%
t = Enemy(BuffComboBox1.ListIndex + 1, 11)
BuffComboBox2.Clear
    If t <> "" Then
        For i = 1 To Len(t)
            BuffComboBox2.AddItem Mid(t, i, 1)
        Next
    Else
        BuffComboBox2.AddItem "��"
    End If
    BuffComboBox2.ListIndex = 1
    LoadBuff 1
End Sub

Private Sub BuffComboBox2_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
    LoadBuff 1
End Sub

Private Sub CBox_Click(Index As Integer)
Dim i%
 If CBox(Index).Value = Checked Then
    CBoxFlag = Index
 Else
    CBoxFlag = Index - 1
 End If
 
 For i = 0 To 5
    If i <= CBoxFlag Then
        CBox(i).Value = Checked
    Else
        CBox(i).Value = Unchecked
    End If
 Next
 SaveSet0
 ReloadTip = True
End Sub






Private Sub CheckState_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If CheckState(Index).ToolTipText = "" Then Exit Sub
    Container5.Top = CheckState(Index).Container.Top + CheckState(Index).Top + Container5.Height + 60
    Container5.Left = CheckState(Index).Left
    Label3.Caption = CheckState(Index).ToolTipText
    Container5.Width = Label3.Width + Label3.Left + 100
    Shape4.Width = Label3.Width + Label3.Left + 100
    Container5.Visible = True
End Sub

Private Sub CommandButton1_Click()
On Error GoTo Outs

SetCount = SetCount + 1
Load SetBox(SetCount)
SetBox(SetCount).Left = SetBox(0).Left
SetBox(SetCount).Top = (SetCount - 1) * 2280 + 600
SetBox(SetCount).Visible = True

Load SetLabel(SetCount)
Set SetLabel(SetCount).Container = SetBox(SetCount)
SetLabel(SetCount).Left = SetLabel(0).Left
SetLabel(SetCount).Top = SetLabel(0).Top
SetLabel(SetCount).Visible = True
SetLabel(SetCount).Caption = "����" + CStr(SetCount)

Load SetCopyButton(SetCount)
Set SetCopyButton(SetCount).Container = SetBox(SetCount)
SetCopyButton(SetCount).Left = SetCopyButton(0).Left
SetCopyButton(SetCount).Top = SetCopyButton(0).Top
SetCopyButton(SetCount).Visible = True


Load SetTipLabel(SetCount)
Set SetTipLabel(SetCount).Container = SetBox(SetCount)
SetTipLabel(SetCount).Left = SetTipLabel(0).Left
SetTipLabel(SetCount).Top = SetTipLabel(0).Top
SetTipLabel(SetCount).Visible = True

Load SetSwitch(SetCount)
Set SetSwitch(SetCount).Container = SetBox(SetCount)
SetSwitch(SetCount).Left = SetSwitch(0).Left
SetSwitch(SetCount).Top = SetSwitch(0).Top
SetSwitch(SetCount).Visible = True

Load SetPic1(SetCount)
Set SetPic1(SetCount).Container = SetBox(SetCount)
SetPic1(SetCount).Left = SetPic1(0).Left
SetPic1(SetCount).Top = SetPic1(0).Top
SetPic1(SetCount).Visible = True
SetPic1(SetCount).LoadImage_FromStdPicture ImageTemp(0).Picture
Load SetPic2(SetCount)
Set SetPic2(SetCount).Container = SetBox(SetCount)
SetPic2(SetCount).Left = SetPic2(0).Left
SetPic2(SetCount).Top = SetPic2(0).Top
SetPic2(SetCount).Visible = True
SetPic2(SetCount).LoadImage_FromStdPicture ImageTemp(0).Picture
Load SetPic3(SetCount)
Set SetPic3(SetCount).Container = SetBox(SetCount)
SetPic3(SetCount).Left = SetPic3(0).Left
SetPic3(SetCount).Top = SetPic3(0).Top
SetPic3(SetCount).Visible = True
SetPic3(SetCount).LoadImage_FromStdPicture ImageTemp(0).Picture
Load SetPic4(SetCount)
Set SetPic4(SetCount).Container = SetBox(SetCount)
SetPic4(SetCount).Left = SetPic4(0).Left
SetPic4(SetCount).Top = SetPic4(0).Top
SetPic4(SetCount).Visible = True
SetPic4(SetCount).LoadImage_FromStdPicture ImageTemp(0).Picture
Load SetPic5(SetCount)
Set SetPic5(SetCount).Container = SetBox(SetCount)
SetPic5(SetCount).Left = SetPic5(0).Left
SetPic5(SetCount).Top = SetPic5(0).Top
SetPic5(SetCount).Visible = True
SetPic5(SetCount).LoadImage_FromStdPicture ImageTemp(0).Picture
Load SetBox2(SetCount)
Set SetBox2(SetCount).Container = SetBox(SetCount)
SetBox2(SetCount).Left = SetBox2(0).Left
SetBox2(SetCount).Top = SetBox2(0).Top
SetBox2(SetCount).Visible = False

Load SetTipLabel2(SetCount)
Set SetTipLabel2(SetCount).Container = SetBox2(SetCount)
SetTipLabel2(SetCount).Left = SetTipLabel2(0).Left
SetTipLabel2(SetCount).Top = SetTipLabel2(0).Top
SetTipLabel2(SetCount).Visible = True

Load SetTipLabel3(SetCount)
Set SetTipLabel3(SetCount).Container = SetBox2(SetCount)
SetTipLabel3(SetCount).Left = SetTipLabel3(0).Left
SetTipLabel3(SetCount).Top = SetTipLabel3(0).Top
SetTipLabel3(SetCount).Visible = True

Load SetTipLabel4(SetCount)
Set SetTipLabel4(SetCount).Container = SetBox2(SetCount)
SetTipLabel4(SetCount).Left = SetTipLabel4(0).Left
SetTipLabel4(SetCount).Top = SetTipLabel4(0).Top
SetTipLabel4(SetCount).Visible = True

Load SetTipLabel5(SetCount)
Set SetTipLabel5(SetCount).Container = SetBox2(SetCount)
SetTipLabel5(SetCount).Left = SetTipLabel5(0).Left
SetTipLabel5(SetCount).Top = SetTipLabel5(0).Top
SetTipLabel5(SetCount).Visible = True

Load SetTipLabel6(SetCount)
Set SetTipLabel6(SetCount).Container = SetBox2(SetCount)
SetTipLabel6(SetCount).Left = SetTipLabel6(0).Left
SetTipLabel6(SetCount).Top = SetTipLabel6(0).Top
SetTipLabel6(SetCount).Visible = True

Load SetTipLabel7(SetCount)
Set SetTipLabel7(SetCount).Container = SetBox2(SetCount)
SetTipLabel7(SetCount).Left = SetTipLabel7(0).Left
SetTipLabel7(SetCount).Top = SetTipLabel7(0).Top
SetTipLabel7(SetCount).Visible = True

Load SetTipLabel8(SetCount)
Set SetTipLabel8(SetCount).Container = SetBox2(SetCount)
SetTipLabel8(SetCount).Left = SetTipLabel8(0).Left
SetTipLabel8(SetCount).Top = SetTipLabel8(0).Top
SetTipLabel8(SetCount).Visible = True

Load SetTipLabel9(SetCount)
Set SetTipLabel9(SetCount).Container = SetBox2(SetCount)
SetTipLabel9(SetCount).Left = SetTipLabel9(0).Left
SetTipLabel9(SetCount).Top = SetTipLabel9(0).Top
SetTipLabel9(SetCount).Visible = True

Load SetTipLabel10(SetCount)
Set SetTipLabel10(SetCount).Container = SetBox2(SetCount)
SetTipLabel10(SetCount).Left = SetTipLabel10(0).Left
SetTipLabel10(SetCount).Top = SetTipLabel10(0).Top
SetTipLabel10(SetCount).Visible = True

Load SetTipLabel11(SetCount)
Set SetTipLabel11(SetCount).Container = SetBox2(SetCount)
SetTipLabel11(SetCount).Left = SetTipLabel11(0).Left
SetTipLabel11(SetCount).Top = SetTipLabel11(0).Top
SetTipLabel11(SetCount).Visible = True

Load SetTipLabel12(SetCount)
Set SetTipLabel12(SetCount).Container = SetBox(SetCount)
SetTipLabel12(SetCount).Left = SetTipLabel12(0).Left
SetTipLabel12(SetCount).Top = SetTipLabel12(0).Top
SetTipLabel12(SetCount).Visible = False
Load SetTipLabel13(SetCount)
Set SetTipLabel13(SetCount).Container = SetBox(SetCount)
SetTipLabel13(SetCount).Left = SetTipLabel13(0).Left
SetTipLabel13(SetCount).Top = SetTipLabel13(0).Top
SetTipLabel13(SetCount).Visible = False

Load SetCombo1(SetCount)
Set SetCombo1(SetCount).Container = SetBox2(SetCount)
SetCombo1(SetCount).Left = SetCombo1(0).Left
SetCombo1(SetCount).Top = SetCombo1(0).Top
SetCombo1(SetCount).Visible = True

Load SetCombo2(SetCount)
Set SetCombo2(SetCount).Container = SetBox2(SetCount)
SetCombo2(SetCount).Left = SetCombo2(0).Left
SetCombo2(SetCount).Top = SetCombo2(0).Top
SetCombo2(SetCount).Visible = True

Load SetCombo3(SetCount)
Set SetCombo3(SetCount).Container = SetBox2(SetCount)
SetCombo3(SetCount).Left = SetCombo3(0).Left
SetCombo3(SetCount).Top = SetCombo3(0).Top
SetCombo3(SetCount).Visible = True

Load SetText1(SetCount)
Set SetText1(SetCount).Container = SetBox2(SetCount)
SetText1(SetCount).Left = SetText1(0).Left
SetText1(SetCount).Top = SetText1(0).Top
SetText1(SetCount).Visible = True

Load SetText2(SetCount)
Set SetText2(SetCount).Container = SetBox2(SetCount)
SetText2(SetCount).Left = SetText2(0).Left
SetText2(SetCount).Top = SetText2(0).Top
SetText2(SetCount).Visible = True

Load SetText3(SetCount)
Set SetText3(SetCount).Container = SetBox2(SetCount)
SetText3(SetCount).Left = SetText3(0).Left
SetText3(SetCount).Top = SetText3(0).Top
SetText3(SetCount).Visible = True

Load SetText4(SetCount)
Set SetText4(SetCount).Container = SetBox2(SetCount)
SetText4(SetCount).Left = SetText4(0).Left
SetText4(SetCount).Top = SetText4(0).Top
SetText4(SetCount).Visible = True

Load SetText5(SetCount)
Set SetText5(SetCount).Container = SetBox2(SetCount)
SetText5(SetCount).Left = SetText5(0).Left
SetText5(SetCount).Top = SetText5(0).Top
SetText5(SetCount).Visible = True

Load SetText6(SetCount)
Set SetText6(SetCount).Container = SetBox2(SetCount)
SetText6(SetCount).Left = SetText6(0).Left
SetText6(SetCount).Top = SetText6(0).Top
SetText6(SetCount).Visible = True

Load SetText7(SetCount)
Set SetText7(SetCount).Container = SetBox2(SetCount)
SetText7(SetCount).Left = SetText7(0).Left
SetText7(SetCount).Top = SetText7(0).Top
SetText7(SetCount).Visible = True
    SetCombo1(SetCount).AddItem "������%"
    SetCombo1(SetCount).AddItem "����ֵ%"
    SetCombo1(SetCount).AddItem "������%"
    SetCombo1(SetCount).AddItem "Ԫ�ؾ�ͨ"
    SetCombo1(SetCount).AddItem "����Ч��"
    SetCombo2(SetCount).AddItem "������%"
    SetCombo2(SetCount).AddItem "����ֵ%"
    SetCombo2(SetCount).AddItem "������%"
    SetCombo2(SetCount).AddItem "Ԫ�ؾ�ͨ"
    SetCombo2(SetCount).AddItem "�������˺��ӳ�"
    SetCombo2(SetCount).AddItem "ˮ�����˺��ӳ�"
    SetCombo2(SetCount).AddItem "�������˺��ӳ�"
    SetCombo2(SetCount).AddItem "�������˺��ӳ�"
    SetCombo2(SetCount).AddItem "�������˺��ӳ�"
    SetCombo2(SetCount).AddItem "�������˺��ӳ�"
    SetCombo2(SetCount).AddItem "�������˺��ӳ�"
    SetCombo2(SetCount).AddItem "�����˺��ӳ�"
    SetCombo3(SetCount).AddItem "������%"
    SetCombo3(SetCount).AddItem "����ֵ%"
    SetCombo3(SetCount).AddItem "������%"
    SetCombo3(SetCount).AddItem "Ԫ�ؾ�ͨ"
    SetCombo3(SetCount).AddItem "������"
    SetCombo3(SetCount).AddItem "�����˺�"
    SetCombo3(SetCount).AddItem "���Ƽӳ�"
    SetCombo1(SetCount).ListIndex = 1
    SetCombo2(SetCount).ListIndex = 1
    SetCombo3(SetCount).ListIndex = 1
Outs:
    
    SetBox(SetCount).Visible = True
    SetPic1(SetCount).LoadImage_FromStdPicture ImageTemp(0)
    SetPic2(SetCount).LoadImage_FromStdPicture ImageTemp(0)
    SetPic3(SetCount).LoadImage_FromStdPicture ImageTemp(0)
    SetPic4(SetCount).LoadImage_FromStdPicture ImageTemp(0)
    SetPic5(SetCount).LoadImage_FromStdPicture ImageTemp(0)
    SetPic1(SetCount).tag = 0
    SetPic2(SetCount).tag = 0
    SetPic3(SetCount).tag = 0
    SetPic4(SetCount).tag = 0
    SetPic5(SetCount).tag = 0
    SetBox2(SetCount).Visible = False
    SetCombo1(SetCount).ListIndex = 1
    SetCombo2(SetCount).ListIndex = 1
    SetCombo3(SetCount).ListIndex = 1
    SetText1(SetCount).Text = "0"
    SetText2(SetCount).Text = "0"
    SetText3(SetCount).Text = "0"
    SetText4(SetCount).Text = "0"
    SetText5(SetCount).Text = "0"
    SetText6(SetCount).Text = "0"
    SetText7(SetCount).Text = "0"
    SetTipLabel12(SetCount).Visible = False
    SetTipLabel13(SetCount).Visible = False
    SetTipLabel13(SetCount).Caption = ""
    
    
If SetCount > 2 Then
    ScrollBar1.Visible = True
    Frame2.Height = 5100 / 2 * SetCount
    ScrollBar1.Max = Frame2.Height - Container2.Height
Else
    ScrollBar1.Visible = False
End If
End Sub


Private Sub CommandButton2_Click()
If Container4.Visible = False Then

    Call ShowArtBoxC
End If
End Sub

Sub ShowArtBoxC()
Dim i%, j%, k%, t%
    On Error Resume Next
        For i = 1 To 1000
            Unload ArtShowImage(i)
        Next

        i = UBound(ArtList)
            For j = 1 To i
                Load ArtShowImage(j)
                ArtShowImage(j).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(j, 1) + ".jpg"
                k = j Mod 7
                If k = 0 Then k = 7
                t = Int((j - 1) / 7)
                ArtShowImage(j).Left = 360 + (k - 1) * 1320
                ArtShowImage(j).Top = 240 + t * 1320
                ArtShowImage(j).Visible = True
                ArtShowImage(j).tag = j
            Next

        t = Int((i - 1) / 7) + 1
        If t > 2 Then
            Container4.Height = 240 + (t + 1.5) * 1320
            Frame2.Height = Container4.Height
            ScrollBar1.Visible = True
            ScrollBar1.Max = Container4.Height - Container2.Height
            
        End If

            
    Container4.Visible = True
    Container4.tag = 0
    Frame2.Caption = "ʥ������"
    Picture2.Visible = True
End Sub

Sub UpdateArtList()
Dim i%, j%, k%
Dim s As String, temp() As String, t() As String

        s = Dir(App.Path + "\Data\Artifact\*.*")
        j = 0
        Do While s <> ""
          j = j + 1
          ReDim Preserve t(1 To j)
                Open App.Path + "\Data\Artifact\" + s For Binary As #1
                   t(j) = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
            s = Dir()
        Loop
        If j = 0 Then Exit Sub
        ReDim ArtList(0 To j, 1 To 11)
        
        For i = 1 To j
                 temp = Split(t(i), vbCrLf)
                    For k = 0 To 10
                         ArtList(i, k + 1) = temp(k)
                    Next
        Next
End Sub


 Sub ShowArtBoxA(Atype As Integer, Index As Integer)
Dim i%, j%, v As Variant, n As Integer, k%, t%
v = Array("��֮��", "��֮��", "ʱ֮ɳ", "��֮��", "��֮��")
n = 0
    On Error Resume Next
        For i = 1 To 1000
            Unload ArtShowImage(i)
        Next
    i = UBound(ArtList)
            For j = 1 To i
                If Val(Mid(ArtList(j, 1), InStr(1, ArtList(j, 1), "_") + 1, 1)) = Atype Then
                    n = n + 1
                    Load ArtShowImage(n)
                    ArtShowImage(n).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(j, 1) + ".jpg"
                    k = n Mod 7
                    If k = 0 Then k = 7
                    t = Int((n - 1) / 7)
                    ArtShowImage(n).Left = 360 + (k - 1) * 1320
                    ArtShowImage(n).Top = 240 + t * 1320
                    ArtShowImage(n).Visible = True
                    ArtShowImage(n).tag = j
                    If CStr(j) = SetPic1(Index).tag Or CStr(j) = SetPic2(Index).tag Or CStr(j) = SetPic3(Index).tag Or CStr(j) = SetPic4(Index).tag Or CStr(j) = SetPic5(Index).tag Then ArtShowImage(n).Opacity = 30
                End If
            Next
            
        t = Int((n - 1) / 7) + 1
        If t > 2 Then
            Container4.Height = 240 + (t + 1.5) * 1320
            Frame2.Height = Container4.Height
            ScrollBar1.Visible = True
            ScrollBar1.Max = Container4.Height - Container2.Height
            
        End If
            
            Container4.Visible = True
            Picture2.Visible = True
            Picture2.tag = Atype
            Container4.tag = Index
            Frame2.Caption = "ѡ��ʥ���" + v(Atype - 1)
End Sub




Private Sub CommandButton3_Click()
Dim flag As Boolean, s As String
Container3.Visible = False
If Val(CommandButton3.tag) > 0 Then
SetTipLabel13(Val(CommandButton3.tag)) = ""

If Label1(2).Visible Then
    SetTipLabel13(Val(CommandButton3.tag)) = Left(SetSelectBox(0).Text, 2) + "4"
Else
    If SetSelectBox(0).Text <> "��" Then SetTipLabel13(Val(CommandButton3.tag)) = SetTipLabel13(Val(CommandButton3.tag)) + SetSelectBox(0).Text + vbCrLf
    If SetSelectBox(1).Text <> "��" Then SetTipLabel13(Val(CommandButton3.tag)) = SetTipLabel13(Val(CommandButton3.tag)) + SetSelectBox(1).Text
End If
SaveSet (Val(CommandButton3.tag))
 ReloadTip = True
Else

flag = False
s = ": "
If Label1(2).Visible Then
    s = Left(SetSelectBox(0).Text, 2) + "4"
Else
    If SetSelectBox(0).Text <> "��" Then s = s + SetSelectBox(0).Text: flag = True
    If SetSelectBox(1).Text <> "��" Then s = s + IIf(flag, "+", "") + SetSelectBox(1).Text
End If
If s = ": " Then s = ": ��"
Textt.Text = s
End If
End Sub


Private Sub CommandButton4_Click()
If Container4.Visible = True Then
    Container4.Visible = False
    Frame2.Caption = "ʥ���﷽��"
    Picture2.Visible = False
    Picture2.tag = ""
    Container4.tag = ""
End If
End Sub

Private Sub CommandButton5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FrmAbout.Show
End Sub

Private Sub CommandButton6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim i%
    For i = 1 To 10
        Unload Text2(i)
    Next
    
    For i = 1 To SetCount
            test.cNumber = Val(AlphaImageChar.tag)
            test.cWeapon = Val(AlphaImageWeap.tag)
    
            
            CreatChar test, FrmMain.lblTab(4).tag, WeaponBox.ListIndex '��ʼ����ɫ
            
            AddArt test, i '����ʥ����ӳ�
            AddBuffListBonus test '����buff�б�ļӳ�
            SolveBonus test '����������Ч
            SolveCharBonus test '���Ͻ�ɫ��Ч
            
            Load Text2(i)
            Text2(i).Visible = True
            Text2(i).Left = 0 + (i - 1) * (Text2(i).Width + 100)
            
            
            Text2(i).Text = "����" + CStr(i) + "��" + vbCrLf + Calc(test, Label2(2).Caption, Val(Label2(3).tag), Val(Label2(5).tag), Val(Text1(17).Text), 1) + vbCrLf + vbCrLf
            test.lowHP = False
            
    Next
                Call pngTab_Click(4, 1)
End Sub


Private Sub CommandButton7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer, tip3 As String, j As Integer, k%
Dim si%, sj%, sk%, ii%
Dim s1() As String, s2() As String, s3() As String
Dim ans As String
If CommandButton7.Text = "�����С���" Then Exit Sub
If CommandButton7.Text = "ȷ�ϼ���" Then

ReDim byctfcj(1 To 1) As String

For i = 0 To 4
    If CheckBox1(i).Value = Checked Then
        si = si + 1
        
        ReDim Preserve s1(1 To si) As String
        s1(si) = CheckBox1(i).Caption
    End If
Next
If si = 0 Then si = 1: ReDim s1(1 To 1) As String

For i = 0 To 4
    If CheckBox2(i).Value = Checked Then
        sj = sj + 1
        
        ReDim Preserve s2(1 To sj) As String
        s2(sj) = CheckBox2(i).Caption
    End If
Next
If sj = 0 Then sj = 1: ReDim s2(1 To 1) As String

For i = 0 To 6
    If CheckBox3(i).Value = Checked Then
        sk = sk + 1
        
        ReDim Preserve s3(1 To sk) As String
        s3(sk) = CheckBox3(i).Caption
    End If
Next
If sk = 0 Then sk = 1: ReDim s3(1 To 1) As String

BYCTM = 0
For i = 0 To 6
    If CheckBox4(i).Value = Checked Then
        BYCTM = BYCTM + 1
        ReDim Preserve BYCTfct(1 To BYCTM)
        BYCTfct(BYCTM) = CheckBox4(i).Caption
    End If
Next
If BYCTM < 1 Then GoTo Outs

i = Val(Text3.Text)
BYCTans = SolveTimes(BYCTM - 1, i) * si * sj * sk
BYCTnow = 0

For i = 0 To 10
    BYCTzcta(i) = ""
Next

BYCTa(0) = 0
tip3 = ""


    Select Case Label2(2).Caption
        Case "��"
            If FrmMain.CheckState(2).Value = Checked Then
                tip3 = "�����������ˮ��"
            End If
            If FrmMain.CheckState(3).Value = Checked Or FrmMain.CheckState(7).Value = Checked Then
                tip3 = "���ڻ���������"
            End If

        Case "ˮ"
            If FrmMain.CheckState(1).Value = Checked Then
                tip3 = "��������ˮ���"
            End If
            
        Case "��"
            If FrmMain.CheckState(1).Value = Checked Then
                tip3 = "���ڻ��������"
            End If
    End Select

            BYCTc.cNumber = Val(AlphaImageChar.tag)
            BYCTc.cWeapon = Val(AlphaImageWeap.tag)

            
            CreatChar BYCTc, FrmMain.lblTab(4).tag, WeaponBox.ListIndex '��ʼ����ɫ
            BYCTc.HPFlag = BYCTc.HPFlag + 4780

            BYCTc.ATKFlag = BYCTc.ATKFlag + 311
            AddBuffListBonus BYCTc '����buff�б�ļӳ�
            

            

CommandButton7.Text = "�����С���"
CommandButton7.Visible = False
ProgressBar1.Visible = True


For i = 1 To UBound(s1)
    BYCTzct(1) = s1(i)
        For j = 1 To UBound(s2)
            BYCTzct(2) = s2(j)
                For k = 1 To UBound(s3)
                    BYCTzct(3) = s3(k)
                    BYCTa(0) = 0
                    Call BYCTSolve(1, Val(Text3.Text))
                    If BYCTa(0) > Val(BYCTzcta(0)) Then
                        BYCTzcta(0) = CStr(BYCTa(0))
                        BYCTzcta(1) = s1(i)
                        BYCTzcta(2) = s2(j)
                        BYCTzcta(3) = s3(k)
                            For ii = 1 To BYCTM
                                BYCTzcta(3 + ii) = CStr(BYCTa(ii))
                            Next
                    End If
                Next
        Next
        
Next


ans = Text3.Text + "���������ı�ҵʥ����ڵ�ǰbuff�����£���ǰ����" + tip3 + "����������˺���" + BYCTzcta(0) + "������ѵ�������ѡ���ǣ�ɳ©-" + BYCTzcta(1) + "������-" + BYCTzcta(2) + "��ñ��-" + BYCTzcta(3) + "������Ѵ�������ǣ�"
For i = 1 To BYCTM
    ans = ans + BYCTfct(i) + BYCTzcta(i + 3) + "����"
Next
ans = Mid(ans, 1, Len(ans) - 1) + "��"

MsgBox ans, , "��ҵ��������"

Outs:
CommandButton7.Visible = True
ProgressBar1.Visible = False
CommandButton7.Text = "�����ҵʥ����"
Container6.Visible = False
Exit Sub
End If

If CommandButton7.Text = "�����ҵʥ����" Then
Container6.Visible = True
CommandButton7.Text = "ȷ�ϼ���"
End If
End Sub
Private Sub BYCTSolve(pos As Integer, last As Integer)
Dim i As Single, tempc As Chars, j%
DoEvents
        If pos < BYCTM Then
          For i = 0 To last Step 1
               BYCTt(pos) = i
               Call BYCTSolve(pos + 1, last - i)
          Next
        Else
                BYCTnow = BYCTnow + 1
                i = BYCTnow / BYCTans
                ProgressBar1.Value = i * 100
                
                
                
              BYCTt(pos) = last
              tempc.cNumber = BYCTc.cNumber
              tempc.cWeapon = BYCTc.cWeapon
              tempc.Level = BYCTc.Level
              tempc.ATK = BYCTc.ATK
              tempc.ATKBonus = BYCTc.ATKBonus
              tempc.ATKFlag = BYCTc.ATKFlag
              tempc.DEF = BYCTc.DEF
              tempc.DEFBonus = BYCTc.DEFBonus
              tempc.DEFFlag = BYCTc.DEFFlag
              tempc.MaxHP = BYCTc.MaxHP
              tempc.HPBonus = BYCTc.HPBonus
              tempc.HPFlag = BYCTc.HPFlag
              tempc.CritRate = BYCTc.CritRate
              tempc.CritDmg = BYCTc.CritDmg
              tempc.EM = BYCTc.EM
              tempc.Energy = BYCTc.Energy
              tempc.HealingBonus = BYCTc.HealingBonus
              tempc.CryoDMG = BYCTc.CryoDMG
              tempc.HydroDMG = BYCTc.HydroDMG
              tempc.CryoDMG = BYCTc.CryoDMG
              tempc.ElectroDMG = BYCTc.ElectroDMG
              tempc.AnemoDMG = BYCTc.AnemoDMG
              tempc.GeoDMG = BYCTc.GeoDMG
              tempc.DendroDMG = BYCTc.DendroDMG
              tempc.PhysicalDMG = BYCTc.PhysicalDMG
              tempc.SPower = BYCTc.SPower
              tempc.UsedE = BYCTc.UsedE
              tempc.UsedEA = BYCTc.UsedEA
              tempc.InSheild = BYCTc.InSheild
              tempc.lowHP = BYCTc.lowHP
              
              
              
              
            '����ʥ����ĸ�����
            For j = 1 To BYCTM
                Select Case BYCTfct(j)
                    Case "����ֵ"
                        tempc.HPBonus = tempc.HPBonus + 5 * BYCTt(j)
                    Case "������"
                        tempc.ATKBonus = tempc.ATKBonus + 5 * BYCTt(j)
                    Case "������"
                        tempc.CritRate = tempc.CritRate + 3.3 * BYCTt(j)
                    Case "�����˺�"
                        tempc.CritDmg = tempc.CritDmg + 6.6 * BYCTt(j)
                    Case "Ԫ�ؾ�ͨ"
                        tempc.EM = tempc.EM + 19.75 * BYCTt(j)
                    Case "����Ч��"
                        tempc.Energy = tempc.Energy + 5.5 * BYCTt(j)
                    Case "������"
                        tempc.DEFBonus = tempc.DEFBonus + 6.2 * BYCTt(j)
                End Select
            Next
            
             '����ʥ�����������
            For j = 1 To 3
                Select Case BYCTzct(j)
                    Case "����ֵ"
                        tempc.HPBonus = tempc.HPBonus + 46.6
                    Case "������"
                        tempc.ATKBonus = tempc.ATKBonus + 46.6
                    Case "������"
                        tempc.CritRate = tempc.CritRate + 31.1
                    Case "�����˺�"
                        tempc.CritDmg = tempc.CritDmg + 62.2
                    Case "Ԫ�ؾ�ͨ"
                        tempc.EM = tempc.EM + 187
                    Case "����Ч��"
                        tempc.Energy = tempc.Energy + 51.8
                    Case "������"
                        tempc.DEFBonus = tempc.DEFBonus + 58.3
                    Case "���Ƽӳ�"
                        tempc.HealingBonus = tempc.HealingBonus + 35.9
                End Select
                If BYCTzct(j) = "����" Then
                    tempc.PyroDMG = tempc.PyroDMG + 58.3
                Else
                    If Right(BYCTzct(j), 1) = "��" Then Call Jug2(tempc, 46.6, True)
                End If
            Next
        
            
            
            AddArt tempc, 1, Textt.Text
            SolveBonus tempc '����������Ч���趯̬
            SolveCharBonus tempc '���Ͻ�ɫ��Ч���趯̬
            
            If BYCTt(1) = 1 And BYCTt(2) = 13 And BYCTzcta(1) = "������" And BYCTzcta(2) = "������" And BYCTzcta(3) = "������" Then
                Print
            End If
            
              i = Val(Calc(tempc, Label2(2).Caption, Val(Label2(3).tag), Val(Label2(5).tag), Val(Text1(17).Text), 2))

              If i > BYCTa(0) Then
                BYCTa(0) = i
                BYCTa(1) = BYCTt(1)
                BYCTa(2) = BYCTt(2)
                BYCTa(3) = BYCTt(3)
                BYCTa(4) = BYCTt(4)
                BYCTa(5) = BYCTt(5)
                BYCTa(6) = BYCTt(6)
                BYCTa(7) = BYCTt(7)
              End If
        End If
End Sub

Private Sub BYCTArt()

End Sub




Private Sub Form_Load()
Dim t As String
On Error Resume Next
    '�Զ�ע��ؼ���Win7һ����Զ�ע�ᣬ����Win8.1��Win10�ؼ���ע�ᣬֱ�����оͻ����
    Dim objTemp As Object
    Set objTemp = CreateObject("BSkin.Container") '�жϴ��������Ƿ�ɹ�
    If Err.Number <> 0 Then '�����������
        UniOCX "BSkin.ocx" '��ȡ��ע��
        RegOCX "BSkin.ocx" '������ע��
    End If '�ؼ��Զ�ע����ϣ�ͬ��ԭ�����ע�������ؼ���
    
    '����RGBAͨ��ͼ�꣨������ͼ�겻ʧ�棩
    SetFormRGBAIcon Me, 0
    SetWindowIcon Me.hWnd
    
    
    '������Ӱ����Ӱģʽ��Ҫ�����ޱ߿�Ч��֮ǰ��
    If FormShadow Is Nothing Then Set FormShadow = New clsShadow
    With FormShadow
        .Depth = 3.5
        .Color = vbBlack
        .Transparency = 100
        .Shadow Me
    
    End With
    '�ޱߴ�Ч���������˲˵��༭�������� None ģʽ����������߿�
    Dim lStyle As Long
    lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    lStyle = Not (WS_CAPTION) And lStyle
    SetWindowLong Me.hWnd, GWL_STYLE, lStyle
    
    'MODEL_TRAY = True '����ģʽ������ֻ��ʾһ���������ݣ�
    
    'Call LoadCMenu '���ز˵�

    Call LoadControl '���ؿؼ�����
                Open App.Path + "\Data\User\C0" For Binary As #1
                   t = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
    LoadChar (Val(t)) 'Ӧ�����ϴμ���ʣ�µ�
    ReloadTip = True
    iniSolveTimes
End Sub

Private Sub Form_Unload(Cancel As Integer) '����ر�ʱ��һ�������ͷ���Դ���������
On Error Resume Next
    Set C_Menu = Nothing
    Set C_Sort = Nothing
    Set c_Subclass = Nothing

    End
End Sub

Sub zMove1(zObject As Object, ToLeft As Single, ToTop As Single, Enable As Boolean) '��̬Ч�� �߳�һ���ȼ����ƶ��㷨��
On Error Resume Next
    zTimCtn1.Enabled = False
    Set mObj1 = Nothing
    mToTop1 = 0: mToLeft1 = 0
    If ToLeft = 0 Then ToLeft = 1
    If ToTop = 0 Then ToTop = 1
    
    If Enable = False Then
        Exit Sub
    ElseIf Enable = True Then
        Set mObj1 = zObject
        mToTop1 = ToTop: mToLeft1 = ToLeft
        zTimCtn1.Enabled = True
    End If
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Container5.Visible = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Container5.Visible = False
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Container5.Visible = False
End Sub

Private Sub ImageTemp2_Click(Index As Integer)
MsgBox Index
End Sub

Private Sub Label2_Click(Index As Integer)
'frmBuff.Show

End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Container5.Visible = False
End Sub

Private Sub Label4_Click(Index As Integer)
If Index = 6 Then
Set Container3.Container = Container6
Container3.Top = 2200
Container3.Left = 60
Container3.Visible = True
CommandButton3.tag = "0"
End If
End Sub

Private Sub lblClose_Click()
Container6.Visible = False
CommandButton7.Text = "�����ҵʥ����"
End Sub

Private Sub LevelBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
LevelBar.tag = "1"
End Sub
Private Sub LevelBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If LevelBar.tag = "1" Then
LevelBar.Cls
LevelBar.Line (0, 0)-(x, LevelBar.Height), , BF
If x > 192 Then x = 192
If x < 2 Then x = 2
lblTab(4).Caption = "��ɫ�ȼ���" + LevelText(Int(x / 2))
lblTab(4).tag = Int(x / 2)
End If
End Sub
Private Sub LevelBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
LevelBar.tag = "0"
SaveSet0
End Sub

Private Sub LevelBox_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
Call LoadSkill
SaveSet0
If Val(AlphaImageChar.tag) = 13 And Index = 1 Then
    Text1(13).Text = CStr(LevelBox(Index).ListIndex)
End If

If Val(AlphaImageChar.tag) = 24 And Index = 2 Then
    Text1(12).Text = CStr(LevelBox(Index).ListIndex)
End If

End Sub




Private Sub ListBox1_Selected(Index As Long)
Dim i As Integer
Label2(1) = "��ǰ���ܣ�" + CurrCharSkill(Index, 3)
Label2(2).Caption = DMGTypetext(Val(CurrCharSkill(Index, 19)))
CurrSkill = CurrCharSkill(Index, 1)


If CurrSkill = "c6q1" Then CheckState(18).Value = Checked: Call LoadBuff(4)
If CurrSkill = "c24q1" Then CheckState(22).Value = Checked: Call LoadBuff(5)


Select Case Label2(2).Caption
    Case "����"
        Label2(2).ForeColor = vbBlack
    Case "��"
        Label2(2).ForeColor = vbRed
    Case "ˮ"
         Label2(2).ForeColor = RGB(0, 128, 255)
    Case "��"
         Label2(2).ForeColor = RGB(153, 217, 234)
    Case "��"
        Label2(2).ForeColor = RGB(128, 0, 128)
    Case "��"
        Label2(2).ForeColor = vbGreen
    Case "��"
        Label2(2).ForeColor = RGB(128, 64, 0)
    Case "��"
        Label2(2).ForeColor = RGB(128, 64, 0)
End Select
Label2(2).Left = Label2(1).Left + Label2(1).Width + 400

    LoadBuff 0


ReloadTip = True
SaveSet0
End Sub


Private Sub SelectBuffBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
SelectBuffBar(Index).tag = "1"
End Sub
Private Sub SelectBuffBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer, j As Integer
If SelectBuffBar(Index).tag = "1" Then
SelectBuffBar(Index).Cls

j = Int(SelectBuffBar(Index).ScaleWidth / SelectBuffBar(Index).LinkTimeout) + 1
i = Round(x / j)

If i < 0 Then i = 0
If i > SelectBuffBar(Index).LinkTimeout Then i = SelectBuffBar(Index).LinkTimeout


SelectBuffBar(Index).Line (0, 0)-(i * j, SelectBuffBar(Index).Height), , BF
BuffLabel(Index).tag = i
BuffLabel(Index).Caption = CStr(i) + "/" + CStr(SelectBuffBar(Index).LinkTimeout) + " " + SelectBuffLabel(SelectCount).tag

End If
End Sub
Private Sub SelectBuffBar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
SelectBuffBar(Index).tag = "0"
If InStr(1, SelectBuffLabel(Index).Caption, "��������") > 0 Then
        test.��ħ = IIf(Val(BuffLabel(Index).tag) = 0, "", "��"): LoadBuff 10
End If


End Sub

Private Sub RBox_Click(Index As Integer)
Dim i%
 If RBox(Index).Value = Checked Then
    RBoxFlag = Index
 Else
    RBoxFlag = Index - 1
 End If
 
 For i = 1 To 4
    If i <= RBoxFlag Then
        RBox(i).Value = Checked
    Else
        RBox(i).Value = Unchecked
    End If
 Next
 SaveSet0
End Sub

Private Sub ScrollBar1_Scroll()
Frame2.Top = -ScrollBar1.Value
End Sub

Private Sub ScrollBar2_Scroll()
ContainerBox.Top = -ScrollBar2.Value
End Sub

Private Sub SetCombo1_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
SaveSet Index
End Sub
Private Sub SetCombo2_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
SaveSet Index
End Sub
Private Sub SetCombo3_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
SaveSet Index
End Sub


Private Sub SetCopyButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call CommandButton1_Click
SetSwitch(SetCount).Value = SetSwitch(Index).Value



SetPic1(SetCount).tag = SetPic1(Index).tag
If Val(SetPic1(SetCount).tag) <> 0 Then
SetPic1(SetCount).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(SetPic1(Index).tag), 1) + ".jpg"
Else
SetPic1(SetCount).LoadImage_FromFile ImageTemp(0).Picture
End If
SetPic2(SetCount).tag = SetPic2(Index).tag
If Val(SetPic2(SetCount).tag) <> 0 Then
SetPic2(SetCount).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(SetPic2(Index).tag), 1) + ".jpg"
Else
SetPic2(SetCount).LoadImage_FromFile ImageTemp(0).Picture
End If
SetPic3(SetCount).tag = SetPic3(Index).tag
If Val(SetPic3(SetCount).tag) <> 0 Then
SetPic3(SetCount).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(SetPic3(Index).tag), 1) + ".jpg"
Else
SetPic3(SetCount).LoadImage_FromFile ImageTemp(0).Picture
End If
SetPic4(SetCount).tag = SetPic4(Index).tag
If Val(SetPic4(SetCount).tag) <> 0 Then
SetPic4(SetCount).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(SetPic4(Index).tag), 1) + ".jpg"
Else
SetPic4(SetCount).LoadImage_FromFile ImageTemp(0).Picture
End If
SetPic5(SetCount).tag = SetPic5(Index).tag
If Val(SetPic5(SetCount).tag) <> 0 Then
SetPic5(SetCount).LoadImage_FromFile App.Path + "\Res\Public\" + ArtList(Val(SetPic5(Index).tag), 1) + ".jpg"
Else
SetPic5(SetCount).LoadImage_FromFile ImageTemp(0).Picture
End If

SetCombo1(SetCount).ListIndex = SetCombo1(Index).ListIndex
SetCombo2(SetCount).ListIndex = SetCombo2(Index).ListIndex
SetCombo3(SetCount).ListIndex = SetCombo3(Index).ListIndex

SetTipLabel13(SetCount) = SetTipLabel13(Index)
SetText1(SetCount) = SetText1(Index)
SetText2(SetCount) = SetText2(Index)
SetText3(SetCount) = SetText3(Index)
SetText4(SetCount) = SetText4(Index)
SetText5(SetCount) = SetText5(Index)
SetText6(SetCount) = SetText6(Index)
SetText7(SetCount) = SetText7(Index)
Call SetSwitch_Click(SetCount, SetSwitch(Index).Value)

End Sub

Private Sub SetPic5_Click(Index As Integer, ByVal Button As Integer)
If Button = 1 Then ShowArtBoxA 5, Index
If Button = 2 Then SetPic5(Index).LoadImage_FromStdPicture ImageTemp(0): SetPic5(Index).tag = 0
End Sub
Private Sub SetPic4_Click(Index As Integer, ByVal Button As Integer)
If Button = 1 Then ShowArtBoxA 4, Index
If Button = 2 Then SetPic4(Index).LoadImage_FromStdPicture ImageTemp(0): SetPic4(Index).tag = 0
End Sub
Private Sub SetPic3_Click(Index As Integer, ByVal Button As Integer)
If Button = 1 Then ShowArtBoxA 3, Index
If Button = 2 Then SetPic3(Index).LoadImage_FromStdPicture ImageTemp(0): SetPic3(Index).tag = 0
End Sub
Private Sub SetPic2_Click(Index As Integer, ByVal Button As Integer)
If Button = 1 Then ShowArtBoxA 2, Index
If Button = 2 Then SetPic2(Index).LoadImage_FromStdPicture ImageTemp(0): SetPic2(Index).tag = 0
End Sub
Private Sub SetPic1_Click(Index As Integer, ByVal Button As Integer)
If Button = 1 Then ShowArtBoxA 1, Index
If Button = 2 Then SetPic1(Index).LoadImage_FromStdPicture ImageTemp(0): SetPic1(Index).tag = 0
End Sub
Private Sub SetPic5_MouseEnter(Index As Integer)
If Val(SetPic5(Index).tag) = 0 Then Exit Sub
ShowBox.Top = SetBox(Index).Top + SetBox(Index).Height - 210 + IIf(Index > 1, -3250, 0)
ShowBox.Left = SetPic5(Index).Left + 120
LoadArtShowBox Val(SetPic5(Index).tag)
ShowBox.Visible = True
End Sub
Private Sub SetPic5_MouseExit(Index As Integer)
ShowBox.Visible = False
End Sub
Private Sub SetPic4_MouseEnter(Index As Integer)
If Val(SetPic4(Index).tag) = 0 Then Exit Sub
ShowBox.Top = SetBox(Index).Top + SetBox(Index).Height - 210 + IIf(Index > 1, -3250, 0)
ShowBox.Left = SetPic4(Index).Left + 120
LoadArtShowBox Val(SetPic4(Index).tag)
ShowBox.Visible = True
End Sub
Private Sub SetPic4_MouseExit(Index As Integer)
ShowBox.Visible = False
End Sub
Private Sub SetPic3_MouseEnter(Index As Integer)
If Val(SetPic3(Index).tag) = 0 Then Exit Sub
ShowBox.Top = SetBox(Index).Top + SetBox(Index).Height - 210 + IIf(Index > 1, -3250, 0)
ShowBox.Left = SetPic3(Index).Left + 120
LoadArtShowBox Val(SetPic3(Index).tag)
ShowBox.Visible = True
End Sub
Private Sub SetPic3_MouseExit(Index As Integer)
ShowBox.Visible = False
End Sub
Private Sub SetPic2_MouseEnter(Index As Integer)
If Val(SetPic2(Index).tag) = 0 Then Exit Sub
ShowBox.Top = SetBox(Index).Top + SetBox(Index).Height - 210 + IIf(Index > 1, -3250, 0)
ShowBox.Left = SetPic2(Index).Left + 120
LoadArtShowBox Val(SetPic2(Index).tag)
ShowBox.Visible = True
End Sub
Private Sub SetPic2_MouseExit(Index As Integer)
ShowBox.Visible = False
End Sub
Private Sub SetPic1_MouseEnter(Index As Integer)
If Val(SetPic1(Index).tag) = 0 Then Exit Sub
ShowBox.Top = SetBox(Index).Top + SetBox(Index).Height - 210 + IIf(Index > 1, -3250, 0)
ShowBox.Left = SetPic1(Index).Left + 120
LoadArtShowBox Val(SetPic1(Index).tag)
ShowBox.Visible = True
End Sub
Private Sub SetPic1_MouseExit(Index As Integer)
ShowBox.Visible = False
End Sub











Private Sub SetSelectBox_SelectionMade(Index As Integer, ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
If SetSelectBox(0).ListIndex = SetSelectBox(1).ListIndex And SetSelectBox(1).ListIndex <> 1 Then
    Label1(2).Visible = True
Else
    Label1(2).Visible = False
End If
End Sub

Private Sub SetSwitch_Click(Index As Integer, Value As Boolean)
If Value Then
    SetTipLabel(Index).Caption = "�����ʥ���ﵥ�����"
    SetBox2(Index).Visible = False
    SetTipLabel12(Index).Visible = False
    SetTipLabel13(Index).Visible = False
Else
    SetTipLabel(Index).Caption = "��ʥ���������������"
    SetBox2(Index).Visible = True
    SetTipLabel12(Index).Visible = True
    SetTipLabel13(Index).Visible = True
End If
SaveSet Index
End Sub


Private Sub SetText1_GotFocus(Index As Integer)
SetText1(Index).SelStart = Len(SetText1(Index).Text)
End Sub
Private Sub SetText1_LostFocus(Index As Integer)
SetText1(Index).Text = CStr(Val(SetText1(Index).Text))
SaveSet (Index)
End Sub
Private Sub SetText2_GotFocus(Index As Integer)
SetText2(Index).SelStart = Len(SetText2(Index).Text)
End Sub
Private Sub SetText2_LostFocus(Index As Integer)
SetText2(Index).Text = CStr(Val(SetText2(Index).Text))
SaveSet (Index)
End Sub
Private Sub SetText3_GotFocus(Index As Integer)
SetText3(Index).SelStart = Len(SetText3(Index).Text)
End Sub
Private Sub SetText3_LostFocus(Index As Integer)
SetText3(Index).Text = CStr(Val(SetText3(Index).Text))
SaveSet (Index)
End Sub
Private Sub SetText4_GotFocus(Index As Integer)
SetText4(Index).SelStart = Len(SetText4(Index).Text)
End Sub
Private Sub SetText4_LostFocus(Index As Integer)
SetText4(Index).Text = CStr(Val(SetText4(Index).Text))
SaveSet (Index)
End Sub
Private Sub SetText5_GotFocus(Index As Integer)
SetText5(Index).SelStart = Len(SetText5(Index).Text)
End Sub
Private Sub SetText5_LostFocus(Index As Integer)
SetText5(Index).Text = CStr(Val(SetText5(Index).Text))
SaveSet (Index)
End Sub
Private Sub SetText6_GotFocus(Index As Integer)
SetText6(Index).SelStart = Len(SetText6(Index).Text)
End Sub
Private Sub SetText6_LostFocus(Index As Integer)
SetText6(Index).Text = CStr(Val(SetText6(Index).Text))
SaveSet (Index)
End Sub
Private Sub SetText7_GotFocus(Index As Integer)
SetText7(Index).SelStart = Len(SetText7(Index).Text)
End Sub
Private Sub SetText7_LostFocus(Index As Integer)
SetText7(Index).Text = CStr(Val(SetText7(Index).Text))
SaveSet (Index)
End Sub

Private Sub CheckState_Click(Index As Integer)
Dim i%
If Index < 8 Then
    Call LoadBuff(7)
If Index = 0 Then
    If CheckState(0).Value = Checked Then
        For i = 1 To 7
            CheckState(i).Value = Unchecked
        Next
    End If
    Call LoadBuff(7)
Else
    If CheckState(Index).Value = Checked Then CheckState(0).Value = Unchecked
End If
Else
LoadBuff 0
End If
SaveBuffFile Val(AlphaImageChar.tag)
End Sub


Private Sub LoadBuffFile(Index As Integer)
Dim t As String, i%, temp() As String, temp2() As String, d As String, j%, temp3() As String, temp4() As String
d = " 9 10 11 14 15 16 18 19 24 12 13 "
Text1ban = True

    If Dir(App.Path + "\Data\User\set") <> "" Then
                    Open App.Path + "\Data\User\set" For Binary As #1
                         t = StrConv(InputB(LOF(1), 1), vbUnicode)
                    Close #1
                Else
                    t = "5" + vbCrLf + "1" + vbCrLf + "1000" + vbCrLf + "800" + vbCrLf + "100" + vbCrLf + "200" + vbCrLf + "1" + vbCrLf + "1" + vbCrLf + "5"
    End If
                temp3 = Split(t, vbCrLf)
                j = 0
                For i = 0 To Text1Bound
                    If InStr(1, d, " " + CStr(i) + " ") > 0 And i <> 12 And i <> 13 Then
                         Text1(i).Text = temp3(j)
                        j = j + 1
                    End If
                Next
                
                                 If Dir(App.Path + "\Data\User\C24\set0") <> "" Then
                                    Open App.Path + "\Data\User\C24\set0" For Binary As #1
                                         t = StrConv(InputB(LOF(1), 1), vbUnicode)
                                    Close #1
                                    temp4 = Split(t, vbCrLf)
                                    Text1(12).Text = temp4(7)
                                 Else
                                    Text1(12).Text = "10"
                                 End If
                                 
                                 If Dir(App.Path + "\Data\User\C13\set0") <> "" Then
                                    Open App.Path + "\Data\User\C13\set0" For Binary As #1
                                         t = StrConv(InputB(LOF(1), 1), vbUnicode)
                                    Close #1
                                    temp4 = Split(t, vbCrLf)
                                    Text1(13).Text = temp4(6)
                                 Else
                                    Text1(13).Text = "10"
                                 End If
                                          
    If Dir(App.Path + "\Data\User\C" + CStr(Index) + "\set") <> "" Then
                Open App.Path + "\Data\User\C" + CStr(Index) + "\set" For Binary As #1
                     t = StrConv(InputB(LOF(1), 1), vbUnicode)
                Close #1
                temp = Split(t, vbCrLf)
                
                BuffComboBox1.ListIndex = Val(temp(0))
                Call BuffComboBox1_SelectionMade("", Val(temp(0)))
                BuffComboBox2.ListIndex = Val(temp(1))
                
                temp2 = Split(temp(2), vbTab)
                For i = 0 To Check1Bound
                    If temp2(i) = "1" Then
                        CheckState(i).Value = Checked
                    Else
                        CheckState(i).Value = Unchecked
                    End If
                Next
                
                temp2 = Split(temp(3), vbTab)
                j = 0
                For i = 0 To Text1Bound
                    If InStr(1, d, " " + CStr(i) + " ") <= 0 Then
                        Text1(i).Text = temp2(j)
                        j = j + 1
                    End If
                Next
    End If
    

    
Text1ban = False
End Sub
Private Sub SaveBuffFile(Index As Integer)
    Dim s As String, i%, d As String
        s = CStr(BuffComboBox1.ListIndex) + vbCrLf + CStr(BuffComboBox2.ListIndex) + vbCrLf
            For i = 0 To Check1Bound
                s = s + IIf(CheckState(i).Value = Checked, "1", "0") + vbTab
            Next
            s = s + vbCrLf
            d = " 9 10 11 14 15 16 18 19 24 12 13 "
            For i = 0 To Text1Bound
                If InStr(1, d, " " + CStr(i) + " ") <= 0 Then
                    s = s + Text1(i).Text + vbTab
                End If
            Next
            Open App.Path + "\Data\User\C" + CStr(Index) + "\set" For Output As #1
                Print #1, s;
            Close #1
            
            s = ""
            For i = 0 To Text1Bound
                If InStr(1, d, " " + CStr(i) + " ") > 0 And i <> 12 And i <> 13 Then
                    s = s + Text1(i).Text + vbCrLf
                End If
            Next
            Open App.Path + "\Data\User\set" For Output As #1
                Print #1, s;
            Close #1
End Sub

Sub LoadBuff(Index As Integer)
Dim t As String, i%, j%, temp As Single, v As Single
Select Case Index
    Case 0
        For i = 1 To 10
            LoadBuff i
        Next
    Case 1
        t = Mid(Label2(2).Caption, 1, 1)
        If t = "��" Then i = 1
        If t = "��" Then i = 2
        If t = "��" Then i = 3
        If t = "��" Then i = 4
        If t = "ˮ" Then i = 5
        If t = "��" Then i = 6
        If t = "��" Then i = 7
        If t = "��" Then i = 8
        If BuffComboBox2.Text = t Then
            j = Val(Enemy(BuffComboBox1.ListIndex + 1, 10))
        Else
            j = Val(Enemy(BuffComboBox1.ListIndex + 1, i + 1))
        End If
        
        
        Label2(3).tag = j - Val(Text1(0).Text) - GBC(8, 40) - GBC(9, 20) - GBC(10, 20) - GBC(43, 40)
        Label2(3).Caption = "���˵�ǰ���ԣ�" + t + "����" + CStr(Round(Label2(3).tag, 2)) + "%"
        If j = 10000 Then Label2(3).Caption = "���˵�ǰ���ԣ�" + t + "��������"
    Case 2
        Label2(5).tag = Val(Text1(1).Text) + GBC(11, 23) + GBC(12, 15) + GBC(13, 15)
        Label2(5).Caption = "���˱���������" + CStr(Round(Label2(5).tag, 2)) + "%"
    Case 3
        temp = 0
        BuffListTip(1) = ""
        v = Val(Text1(2).Text)
        If v > 0 Then temp = temp + v: BuffListTip(1) = BuffListTip(1) + "�����Զ���Ĺ������ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(14, 20)
        If v > 0 Then temp = temp + v: BuffListTip(1) = BuffListTip(1) + "��������4���׵Ĺ������ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(15, 20)
        If v > 0 Then temp = temp + v: BuffListTip(1) = BuffListTip(1) + "����ǧ��4���׵Ĺ������ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(16, 20)
        If v > 0 Then temp = temp + v: BuffListTip(1) = BuffListTip(1) + "��������Ӣ��̷�Ĺ������ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = 0
        If CheckState(17).Value = Checked Then v = (Val(Text1(10).Text)) * 5 + 15: t = "����"
        If CheckState(40).Value = Checked Then If v < ((Val(Text1(18).Text)) * 5 + 15) Then v = (Val(Text1(18).Text)) * 5 + 15: t = "����"
        If CheckState(41).Value = Checked Then If v < ((Val(Text1(19).Text)) * 5 + 15) Then v = (Val(Text1(19).Text)) * 5 + 15: t = "����"
        If v > 0 Then temp = temp + v: BuffListTip(1) = BuffListTip(1) + "����" + t + "֮��Ĺ������ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(37, 25)
        If v > 0 Then temp = temp + v: BuffListTip(1) = BuffListTip(1) + "����˫�����Ĺ������ӳɣ�" + CStr(v) + "%" + vbCrLf

        Label2(8).tag = Round(temp, 2)
        Label2(8).Caption = "�������ٷֱȼӳɣ�" + CStr(Round(temp, 2)) + "%"
    Case 4
        temp = 0
        BuffListTip(2) = ""
        v = Val(Text1(3).Text)
        If v > 0 Then temp = temp + v: BuffListTip(2) = BuffListTip(2) + "�����Զ���Ĺ������ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(18, 20)
        If v > 0 Then temp = temp + v: BuffListTip(2) = BuffListTip(2) + "���԰�����Q�Ĺ������ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(19, 20)
        If v > 0 Then temp = temp + v: BuffListTip(2) = BuffListTip(2) + "���Ծ������׵Ĺ������ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(35, 372)
        If v > 0 Then temp = temp + v: BuffListTip(2) = BuffListTip(2) + "��������ǽ�Ĺ������ӳɣ�" + CStr(v) + "" + vbCrLf
    
        Label2(10).tag = Round(temp, 2)
        Label2(10).Caption = "���������ּӳɣ�" + CStr(Round(temp, 2))
    Case 5
        temp = 0
        BuffListTip(3) = ""
        v = Val(Text1(4).Text)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "�����Զ�������ˣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(20, 20)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "������Ҷ�츳��Ԫ���˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(21, 35)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "��������4���׵�Ԫ���˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(22, 20)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "����Ī��Q�����ˣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(23, 20)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "�����׵罫��E�����ˣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(24, 20)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "���Ը���Q�ı�Ԫ���˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(36, 20)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "����Ԫ�ؾ��͵�" + Label2(2).Caption + "Ԫ���˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(44, 0)
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "���԰׳�֮����Ԫ���˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        
        If CheckState(10).Value = Checked Then temp = temp + 15: BuffListTip(3) = BuffListTip(3) + "����˫�ҹ��������ˣ�15%" + vbCrLf
        t = ListBox1.ItemKey(ListBox1.ListIndex)
        i = Val(Text1(10).Text)
        v = 0
            If (InStr(2, t, "a") > 0 Or InStr(2, t, "d") > 0 Or InStr(2, t, "c") > 0) And CheckState(17).Value = Checked Then
                v = 16 + (i - 1) * 4
            End If
        If v > 0 Then temp = temp + v: BuffListTip(3) = BuffListTip(3) + "���Կ���֮������ˣ�" + CStr(v) + "%" + vbCrLf
        
        Label2(12).tag = Round(temp, 2)
        Label2(12).Caption = "�˺��ӳɣ�" + CStr(Round(temp, 2)) + "%"
    Case 6
        temp = 0
        BuffListTip(4) = ""
        v = Val(Text1(5).Text)
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "�����Զ����Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(25, 20)
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "����ɰ���츳��Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(26, 200)
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "������Ҷ����2��Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(27, 200)
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "���Եϰ�������6��Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(28, 125)
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "���԰������츳��Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
        v = GBC(42, 120)
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "���Խ̹�4���׵�Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
        v = 0
        If CheckState(41).Value = Checked Then v = Val(Text1(19).Text) * 25 + 75
        If v > 0 Then temp = temp + v: BuffListTip(4) = BuffListTip(4) + "���Ա���֮���Ԫ�ؾ�ͨ�ӳɣ�" + CStr(v) + "" + vbCrLf
    
        Label2(14).tag = Round(temp, 2)
        Label2(14).Caption = "Ԫ�ؾ�ͨ�ӳɣ�" + CStr(Round(temp, 2))
    Case 7
        temp = 0
        BuffListTip(5) = ""
        v = Val(Text1(6).Text)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "�����Զ���ı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(29, 12)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "���Ժ����츳�ı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(30, 0)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "������ɯ�����츳�ı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(31, 12)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "���Ժ�������4�ı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(38, 15)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "����˫�������ı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(34, 20)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "�����ɵ����µı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(35, 12)
        If v > 0 Then temp = temp + v: BuffListTip(5) = BuffListTip(5) + "��������ǽ�ı����ʼӳɣ�" + CStr(v) + "%" + vbCrLf
    
        Label2(16).tag = Round(temp, 2)
        Label2(16).Caption = "�����ʼӳɣ�" + CStr(Round(temp, 2)) + "%"
    Case 8
        temp = 0
        BuffListTip(6) = ""
        v = Val(Text1(7).Text)
        If v > 0 Then temp = temp + v: BuffListTip(6) = BuffListTip(6) + "�����Զ���ı����˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(32, 60)
        If v > 0 Then temp = temp + v: BuffListTip(6) = BuffListTip(6) + "���Ծ�������6�ı����˺��ӳɣ�" + CStr(v) + "%" + vbCrLf
        v = GBC(34, 20)
        If v > 0 Then temp = temp + v: BuffListTip(6) = BuffListTip(6) + "�����ɵ����µı����˺��ӳɣ�" + CStr(v) + "%" + vbCrLf

        Label2(18).tag = Round(temp, 2)
        Label2(18).Caption = "�����˺��ӳɣ�" + CStr(Round(temp, 2)) + "%"
    Case 9
        Label2(20).tag = Val(Text1(8).Text) + GBC(33, 60)
        Label2(20).Caption = "Ԫ�س���Ч�ʼӳɣ�" + CStr(Round(Label2(20).tag, 2)) + "%"
    Case 10
        If (InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Or InStr(2, CurrSkill, "d") > 0) And test.��ħ <> "" Then
            Label2(2).Caption = test.��ħ
        Else
            For i = 1 To UBound(CurrCharSkill)
                If CurrCharSkill(i, 1) = CurrSkill Then
                    Label2(2).Caption = DMGTypetext(Val(CurrCharSkill(i, 19)))
                    Exit For
                End If
            Next
        End If
            LoadBuff 1
            CheckBox2(3).Caption = Mid(Label2(2).Caption, 1, 1) + "�˱�"
            Select Case Label2(2).Caption
                Case "����"
                    Label2(2).ForeColor = vbBlack
                Case "��"
                    Label2(2).ForeColor = vbRed
                Case "ˮ"
                     Label2(2).ForeColor = RGB(0, 128, 255)
                Case "��"
                     Label2(2).ForeColor = RGB(153, 217, 234)
                Case "��"
                    Label2(2).ForeColor = RGB(128, 0, 128)
                Case "��"
                    Label2(2).ForeColor = vbGreen
                Case "��"
                    Label2(2).ForeColor = RGB(128, 64, 0)
                Case "��"
                    Label2(2).ForeColor = RGB(128, 64, 0)
            End Select
            
End Select
End Sub
Private Function GBC(Index As Integer, v As Integer) As Single
Dim i As Integer, t As String, temp() As String, tempc As Chars, templ As Integer, tempco As Integer, templ2 As Integer, templ3 As Integer, v1 As Variant
GBC = 0
If CheckState(Index).Value = Unchecked Then Exit Function


GBC = IIf(CheckState(Index).Value = Checked, v, 0)
If Index = 8 Then
    If Mid(Label2(2).Caption, 1, 1) <> "��" And Mid(Label2(2).Caption, 1, 1) <> "��" And Mid(Label2(2).Caption, 1, 1) <> "��" Then
        GBC = GBC
    Else
        GBC = 0
    End If
End If


If Index = 10 Then
    If Mid(Label2(2).Caption, 1, 1) <> "��" Then GBC = 0
End If

If Index = 16 Then '����
    i = Val(Text1(9).Text)
    If i < 0 Then i = 1: Text1(9).Text = "1"
    If i > 5 Then i = 5: Text1(9).Text = "5"
    GBC = 24 + (i - 1) * 6
End If

If Index = 17 Then '�Թ�
    i = Val(Text1(10).Text)
    If i < 0 Then i = 1: Text1(10).Text = "1"
    If i > 5 Then i = 5: Text1(10).Text = "5"
    GBC = 20 + (i - 1) * 5
End If

If Index = 18 Then '������
    If Dir(App.Path + "\Data\User\C7", vbDirectory) <> "" Then  '�����ļ���
                Open App.Path + "\Data\User\C7\set0" For Binary As #1
                     t = StrConv(InputB(LOF(1), 1), vbUnicode)
                Close #1
                temp = Split(t, vbCrLf)
                templ = Val(temp(0))  '1���ȼ�
                tempco = Val(temp(1)) '2������
                tempc.cWeapon = (Val(temp(2))) '3������
                tempc.cNumber = 7
                templ2 = Val(temp(3)) '4�������ȼ�
                templ3 = Val(temp(7)) 'Q�ȼ�
                v1 = Array(56, 60.2, 64.4, 70, 74.2, 78.4, 84, 89.6, 95.2, 100.8, 106.4, 112, 119, 126, 133)
                CreatChar tempc, templ, templ2
                GBC = Round(tempc.ATK * (v1(templ3 - 1) + IIf(tempco >= 0, 20, 0)) / 100, 2)
    Else
        GBC = 1202
    End If
End If

If Index = 19 Then '����
    If Dir(App.Path + "\Data\User\C16", vbDirectory) <> "" Then  '�����ļ���
                Open App.Path + "\Data\User\C16\set0" For Binary As #1
                     t = StrConv(InputB(LOF(1), 1), vbUnicode)
                Close #1
                temp = Split(t, vbCrLf)
                templ = Val(temp(0))  '1���ȼ�
                tempco = Val(temp(1)) '2������
                tempc.cWeapon = (Val(temp(2))) '3������
                tempc.cNumber = 16
                templ2 = Val(temp(3)) '4�������ȼ�
                templ3 = Val(temp(7)) 'Q�ȼ�
                v1 = Array(42.96, 46.18, 49.4, 53.7, 56.92, 60.14, 64.44, 68.74, 73.03, 77.33, 81.62, 85.92, 91.29, 96.66, 102.03)
                CreatChar tempc, templ, templ2
                GBC = Round(tempc.ATK * v1(templ3 - 1) / 100, 2)
    Else
        GBC = 793
    End If
End If


If Index = 20 Then '��Ҷ����
    i = Val(Text1(11).Text)
    If i < 0 Then i = 0: Text1(11).Text = "0"
    If CheckState(26).Value = Checked Then i = i + 200
    GBC = Round(i * 0.04, 2)
    If Mid(Label2(2).Caption, 1, 1) = "��" Or Mid(Label2(2).Caption, 1, 1) = "��" Or Mid(Label2(2).Caption, 1, 1) = "��" Then GBC = 0
End If

If Index = 21 Then
    If Mid(Label2(2).Caption, 1, 1) = "��" Or Mid(Label2(2).Caption, 1, 1) = "��" Or Mid(Label2(2).Caption, 1, 1) = "��" Then GBC = 0
End If

If Index = 22 Then 'Ī��Q
    i = Val(Text1(12).Text)
    If i < 0 Then i = 0: Text1(12).Text = "0"
    If i > 10 Then i = 10
    GBC = 40 + 2 * i
End If


If Index = 23 Then '����E
    i = Val(Text1(13).Text)
    If i < 0 Then i = 0: Text1(13).Text = "0"
    If i > 9 Then i = 9
    GBC = Val(CharList(Val(AlphaImageChar.tag), 3)) * (0.21 + 0.01 * i)
        t = ListBox1.ItemKey(ListBox1.ListIndex)
            If InStr(2, t, "q") <= 0 Then GBC = 0
End If


If Index = 24 Then
    If Mid(Label2(2).Caption, 1, 1) <> "��" Then GBC = 0
End If

If Index = 25 Then 'ɰ�Ǿ�ͨ
    i = Val(Text1(14).Text)
    If i < 0 Then i = 0: Text1(14).Text = "0"
    GBC = Round(i * 0.2) + 50
End If


If Index = 30 Then '��ɯ���Ǳ���
    i = Val(Text1(15).Text)
    If i < 0 Then i = 0: Text1(15).Text = "0"
    If i > 100 Then i = 100
    GBC = Round(i * 0.15, 2)
End If

If Index = 32 Then
    If Mid(Label2(2).Caption, 1, 1) <> "��" Then GBC = 0
End If

If Index = 33 Then '��������
    i = Val(Text1(16).Text)
    If i < 0 Then i = 0: Text1(16).Text = "0"
    GBC = 20 + i * 0.1
End If

If Index = 36 Then
    If Mid(Label2(2).Caption, 1, 1) = "��" Then GBC = 0
End If

If Index = 38 Then
  If CheckState(3).Value = Unchecked And CheckState(7).Value = Unchecked Then GBC = 0
End If

If Index = 43 Then
    If Mid(Label2(2).Caption, 1, 1) <> "��" Then GBC = 0
End If

If Index = 44 Then '�׳�֮��
    i = Val(Text1(24).Text)
    If i < 0 Then i = 1: Text1(24).Text = "1"
    If i > 5 Then i = 5: Text1(24).Text = "5"
    GBC = 10 + (i - 1) * 2.5
End If

End Function






Private Sub SetTipLabel12_Click(Index As Integer)
Set Container3.Container = Frame2
Container3.Top = SetBox(Index).Top + 600
Container3.Left = SetTipLabel12(Index).Left - 1000
Container3.Visible = True
CommandButton3.tag = Index
End Sub



Private Sub Text1_Change(Index As Integer)
If Text1ban = True Then Exit Sub
LoadBuff 0
SaveBuffFile Val(AlphaImageChar.tag)
End Sub


Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Container5.Visible = False
End Sub

Private Sub Text3_Change()
Dim i%
i = Int(Val(Text3))
If i < 1 Then i = 1
If i > 40 Then i = 40
Text3.Text = CStr(i)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim i As Integer
                i = Val(Timer1.tag)
                LevelBox(i).ListIndex = BoxTemp(i)
                If LevelBox(i).ListIndex = BoxTemp(i) Then Timer1.tag = CStr(i + 1)
                
                If Timer1.tag = "3" Then
                                ListBox1.ListIndex = BoxTemp(3)
                                Call ListBox1_Selected(Val(BoxTemp(3)))  '9��ѡ����
                                Timer1.Enabled = False
                End If
                
                
End Sub

Private Sub WeaponBox_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
SaveSet0
End Sub

Private Sub zTimCtn1_Timer() '��̬Ч�� �߳�һ������Ч�ʸ��ߵ� Timer �ؼ�����ϵͳ���� Timer��֧�ֵ���ʱ���ܣ�
On Error Resume Next
    Dim ml As Single, mt As Single
    If mObj1.Left < mToLeft1 Then
        ml = mToLeft1 - mObj1.Left
        ml = ml / 9
        mObj1.Left = mObj1.Left + ml
    ElseIf mObj1.Left > mToLeft1 Then
        ml = mObj1.Left - mToLeft1
        ml = ml / 9
        mObj1.Left = mObj1.Left - ml
    End If
    
    If mObj1.Top < mToTop1 Then
        mt = mToTop1 - mObj1.Top
        mt = mt / 9
        mObj1.Top = mObj1.Top + mt
    ElseIf mObj1.Top > mToTop1 Then
        mt = mObj1.Top - mToTop1
        mt = mt / 9
        mObj1.Top = mObj1.Top - mt
    End If
    
    If Round(mObj1.Left) = Round(mToLeft1) And Round(mObj1.Top) = Round(mToTop1) Then
        zTimCtn1.Enabled = False
    End If
End Sub



Private Sub LoadControl() 'д�ɷ�������ʽ������ã�������ά��
On Error Resume Next
Dim i%, j%, sumi%, sumj%
    FrmMain.Width = 10995
    FrmMain.Height = 7920
    DMGTypetext(1) = "��"
    DMGTypetext(2) = "ˮ"
    DMGTypetext(3) = "��"
    DMGTypetext(4) = "��"
    DMGTypetext(5) = "��"
    DMGTypetext(6) = "��"
    DMGTypetext(7) = "��"
    DMGTypetext(8) = "����"
        ArtTypetext(1) = "��ʿ"
        ArtTypetext(2) = "��Ů"
        ArtTypetext(3) = "�԰�"
        ArtTypetext(4) = "ħŮ"
        ArtTypetext(5) = "����"
        ArtTypetext(6) = "�ɻ�"
        ArtTypetext(7) = "����"
        ArtTypetext(8) = "���"
        ArtTypetext(9) = "ƽ��"
        ArtTypetext(10) = "ǧ��"
        ArtTypetext(11) = "����"
        ArtTypetext(12) = "ˮ��"
        ArtTypetext(13) = "��Ե"
        ArtTypetext(14) = "�Ƕ�"
        ArtTypetext(15) = "׷��"
        ArtTypetext(16) = "����"
        ArtTypetext(17) = "����"
        ArtTypetext(18) = "����"
    
    
    
    
    ContainerBox.Left = 0
    ContainerBox.Top = 0
    
    i = 1
    Do While Dir(App.Path + "\Res\Public\S_" + CStr(i) + ".jpg") <> ""
        Load ImageTemp(i)
        ImageTemp(i).Picture = LoadPicture(App.Path + "\Res\Public\S_" + CStr(i) + ".jpg")
        i = i + 1
    Loop
    
    i = 1
    Do While Dir(App.Path + "\Res\Public\A" + CStr(i) + "_1.jpg") <> ""
        ImageTemp2(0).Picture = LoadPicture(App.Path + "\Res\Public\A" + CStr(i) + "_1.jpg")
        Load ImageTemp2(i)
        Set ImageTemp2(i).Picture = Nothing
        ImageTemp2(i).PaintPicture ImageTemp2(0).Picture, 0, 0, 21, 21, 0, 0, 70, 70
        ImageTemp2(i).Visible = False
        i = i + 1
    Loop
    
    
    'Call CommandButton1_Click
    
    For i = 0 To Check1Bound
        CheckState(i).Font.SIZE = CheckState(i).Font.SIZE - 1
        CheckState(i).BackColor = 16244694
        Text1(i).BackColor = 16244694
    Next
    ScrollBar2.Max = ContainerBox.Height - Container1(2).Height
    ContainerBox.Top = 0
    Container1(0).BackColor = vbWhite
    
    
    
    
    
    'AlphaImage ͼƬ�ؼ�������ǿ����Ҫ��������UI�زĶѵ���ֱ����ʾPNG�ȸ�ʽͼƬ���������Ͽ���ʵ���κν���Ч�������Ҽ����Բ˵������ؼ�ͼƬ��
    pngCloseBG.Opacity = 0
    pngMinimizeBG.Opacity = 0
    


    For i = 2 To UBound(Enemy)
        BuffComboBox1.AddItem Enemy(i, 1)
        BuffComboBox2.AddItem Enemy(i, 1)
    Next
    BuffComboBox1.ListIndex = 1
    BuffComboBox1_SelectionMade "", 1
    
    
    WeaponBox.AddItem 1
    For i = 20 To 90 Step 10
        If i <> 30 Then WeaponBox.AddItem i
    Next
    WeaponBox.ListIndex = 8
    For i = 0 To 1
        SetSelectBox(i).AddItem "��"
            For j = 1 To UBound(ArtTypetext)
                SetSelectBox(i).AddItem ArtTypetext(j) + "2", , ImageTemp2(j).Image
            Next
        SetSelectBox(i).ListIndex = 1
    Next
        Label1(2).Visible = False
    
    
    For i = 0 To 2
        For j = 1 To 15
            LevelBox(i).AddItem j
        Next
        LevelBox(i).ListIndex = 10
    Next
    

    CBoxFlag = 5
    RBoxFlag = 4
    LevelBar.Line (0, 0)-(192, LevelBar.Height), , BF
    lblTab(4).tag = 96


    i = 0
    j = 0
    
        For i = 1 To 90
            j = j + 1
            LevelText(j) = CStr(i) + "��"
                If i Mod 10 = 0 And i <> 10 And i <> 30 Then
                    j = j + 1
                    LevelText(j) = CStr(i) + "����ͻ��"
                End If
        Next
 

End Sub


Sub LoadChar(Index As Integer)
On Error Resume Next
Dim i%, j%, sumi%, sumj%
Dim t As String, temp() As String, temp2() As String

                Open App.Path + "\Data\Data\C" + CStr(Index) + "_2" For Binary As #1
                   t = StrConv(InputB(LOF(1), 1), vbUnicode)
                 Close #1
                   temp = Split(t, vbCrLf)
                   sumi = UBound(temp) + 1
                   sumj = 19

    ReDim CurrCharSkill(1 To sumi, 1 To sumj)
        For i = 1 To sumi
            temp2 = Split(temp(i - 1), vbTab)
            For j = 1 To sumj
                CurrCharSkill(i, j) = temp2(j - 1)
            Next
        Next
        
AlphaImageChar.LoadImage_FromFile App.Path + "\Res\Public\C" + CStr(Index) + ".png"
AlphaImageChar.tag = Index
lblTab(10).Caption = CharList(Index, 1)

For i = 0 To 4
    CheckBox1(i).Value = Unchecked
    CheckBox2(i).Value = Unchecked
Next
For i = 0 To 6
    CheckBox3(i).Value = Unchecked
    CheckBox4(i).Value = Unchecked
Next

Select Case CharList(Index, 1)
    Case "����"
        CheckBox1(0).Value = Checked
        CheckBox2(3).Value = Checked
        CheckBox3(3).Value = Checked
        CheckBox4(1).Value = Checked
        CheckBox4(3).Value = Checked
        CheckBox4(4).Value = Checked
    Case Else
        CheckBox1(1).Value = Checked
        CheckBox2(3).Value = Checked
        CheckBox3(3).Value = Checked
        CheckBox4(0).Value = Checked
        CheckBox4(3).Value = Checked
        CheckBox4(4).Value = Checked
End Select

 
 

    test.��ħ = ""

    If CharList(Index, 1) = "��" Then test.��ħ = "��"
    If CharList(Index, 1) = "����" Then test.��ħ = "��"
    If CharList(Index, 1) = "��¬��" Then test.��ħ = "��"
    If CharList(Index, 1) = "����类�" Then test.��ħ = "��"
    

    
    
    If CharList(Index, 5) <> "" Then '������ʾ
        lblTab(11).Caption = "*" + CharList(Index, 5)
    Else
        lblTab(11).Caption = ""
    End If
redo:
    If Dir(App.Path + "\Data\User\C" + CStr(Index), vbDirectory) <> "" Then '�����ļ���
                Open App.Path + "\Data\User\C" + CStr(Index) + "\set0" For Binary As #1
                     t = StrConv(InputB(LOF(1), 1), vbUnicode)
                Close #1
                temp = Split(t, vbCrLf)
                lblTab(4).tag = Val(temp(0))  '1���ȼ�
                lblTab(4).Caption = "��ɫ�ȼ���" + LevelText(Val(temp(0)))
                LevelBar.Cls
                LevelBar.Line (0, 0)-(Val(temp(0)) * 2, LevelBar.Height), , BF
                CBoxFlag = Val(temp(1)) '2������
                LoadWeapon (Val(temp(2))) '3������
                WeaponBox.ListIndex = Val(temp(3)) '4�������ȼ�
                RBoxFlag = Val(temp(4)) '5������
                For i = 0 To 5
                   If i <= CBoxFlag Then
                       CBox(i).Value = Checked
                   Else
                       CBox(i).Value = Unchecked
                   End If
                Next
    
    
                For i = 0 To 4
                   If i <= RBoxFlag Then
                       RBox(i).Value = Checked
                   Else
                       RBox(i).Value = Unchecked
                   End If
                Next
                
                
                
                
                LevelBox(0).ListIndex = Val(temp(5))
                LevelBox(1).ListIndex = Val(temp(6))
                LevelBox(2).ListIndex = Val(temp(7))

                                ListBox1.ListIndex = Val(temp(8))
                                Call ListBox1_Selected(Val(temp(8)))  '9��ѡ����
                Call LoadSkill '���ر���
                
                'BoxTemp(0) = Val(temp(5))
                'BoxTemp(1) = Val(temp(6))
                'BoxTemp(2) = Val(temp(7))
                'BoxTemp(3) = Val(temp(8))
                'Timer1.tag = "0"
                'Timer1.Enabled = True

                
                For i = 1 To 100
                    SetBox(i).Visible = False
                Next
                SetCount = 0
                
                i = 1
                Do While Dir(App.Path + "\Data\User\C" + CStr(Index) + "\set" + CStr(i)) <> ""
                    Open App.Path + "\Data\User\C" + CStr(Index) + "\set" + CStr(i) For Binary As #1
                         t = StrConv(InputB(LOF(1), 1), vbUnicode)
                    Close #1
                    temp = Split(t, vbCrLf)
                    Call CommandButton1_Click
                    
                    
                    If temp(0) = "1" Then
                        SetSwitch(i).Value = True
                        SetBox2(i).Visible = False
                    Else
                        SetSwitch(i).Value = False
                         SetBox2(i).Visible = True
                        
                    End If
                    
                    If temp(1) = "0" Then
                        SetPic1(i).LoadImage_FromStdPicture ImageTemp(0)
                    Else
                        SetPic1(i).LoadImage_FromFile App.Path + "\res\public\" + ArtList(Val(temp(1)), 1) + ".jpg"
                    End If
                        SetPic1(i).tag = Val(temp(1))
                        
                    If temp(2) = "0" Then
                        SetPic2(i).LoadImage_FromStdPicture ImageTemp(0)
                    Else
                        SetPic2(i).LoadImage_FromFile App.Path + "\res\public\" + ArtList(Val(temp(2)), 1) + ".jpg"
                    End If
                        SetPic2(i).tag = Val(temp(2))
                        
                    If temp(3) = "0" Then
                        SetPic3(i).LoadImage_FromStdPicture ImageTemp(0)
                    Else
                        SetPic3(i).LoadImage_FromFile App.Path + "\res\public\" + ArtList(Val(temp(3)), 1) + ".jpg"
                    End If
                        SetPic3(i).tag = Val(temp(3))
                        
                    If temp(4) = "0" Then
                        SetPic4(i).LoadImage_FromStdPicture ImageTemp(0)
                    Else
                        SetPic4(i).LoadImage_FromFile App.Path + "\res\public\" + ArtList(Val(temp(4)), 1) + ".jpg"
                    End If
                        SetPic4(i).tag = Val(temp(4))
                        
                    If temp(5) = "0" Then
                        SetPic5(i).LoadImage_FromStdPicture ImageTemp(0)
                    Else
                        SetPic5(i).LoadImage_FromFile App.Path + "\res\public\" + ArtList(Val(temp(5)), 1) + ".jpg"
                    End If
                        SetPic5(i).tag = Val(temp(5))
                        
                    SetCombo1(i).ListIndex = Val(temp(6))
                    SetCombo2(i).ListIndex = Val(temp(7))
                    SetCombo3(i).ListIndex = Val(temp(8))
                    SetText1(i).Text = temp(9)
                    SetText2(i).Text = temp(10)
                    SetText3(i).Text = temp(11)
                    SetText4(i).Text = temp(12)
                    SetText5(i).Text = temp(13)
                    SetText6(i).Text = temp(14)
                    SetText7(i).Text = temp(15)
                    SetTipLabel13(i).Caption = temp(16)
                    SetTipLabel13(i).Caption = SetTipLabel13(i).Caption + temp(17)
                    
                    
                    i = i + 1
                Loop
                
                LoadBuffFile Index
                
    Else '�������ļ���
        MkDir (App.Path + "\Data\User\C" + CStr(Index))
        If CharList(Index, 2) = "1" Then t = "1"
        If CharList(Index, 2) = "2" Then t = "49"
        If CharList(Index, 2) = "3" Then t = "35"
        If CharList(Index, 2) = "4" Then t = "22"
        If CharList(Index, 2) = "5" Then t = "17"

            Open App.Path + "\Data\User\C" + CStr(Index) + "\set0" For Output As #1
                Print #1, "96" + vbCrLf + "5" + vbCrLf + t + vbCrLf + "8" + vbCrLf + "4" + vbCrLf + "10" + vbCrLf + "10" + vbCrLf + "10" + vbCrLf + "1";
            Close #1
            
            Open App.Path + "\Data\User\C" + CStr(Index) + "\set1" For Output As #1
                Print #1, "1" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "1" + vbCrLf + "1" + vbCrLf + "1" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "0" + vbCrLf + "";
            Close #1
        GoTo redo
    End If
    




End Sub


Private Sub LoadSkill() '��ʾ����
Dim i%, j%, tip As String, M As String, v As Variant
    ListBox1.Clear
    For i = 1 To UBound(CurrCharSkill)
        tip = ""
        If CurrCharSkill(i, 1) <> "" Then

            
            If CurrCharSkill(i, 1) = "c1e2" Then tip = " ������"
            If CurrCharSkill(i, 1) = "c8s1" Then '�������ܺ���
                v = Array(1386, 1525, 1675, 1837, 2010, 2195, 2392, 2600, 2819, 3050, 3293, 3547, 3813, 4090, 4379)
                tip = "����ֵ + " + CStr(v(LevelBox(1).ListIndex - 1))
            End If
            
            
            M = CStr(GetBonus(CurrCharSkill(i, 1))) + "%" + tip
            
            
            ListBox1.AddItem CurrCharSkill(i, 3), M, CurrCharSkill(i, 1), ImageTemp(Val(CurrCharSkill(i, 2)))
        End If
    Next
    If ListBox1.ListIndex = 0 Then ListBox1.ListIndex = 1
End Sub

 Sub LoadWeapon(Index As Integer)
    AlphaImageWeap.LoadImage_FromFile App.Path + "\Res\Public\W_" + CStr(Index) + ".png"
    AlphaImageWeap.tag = Index
    lblTab(9).Caption = WeaponList(Index, 1)
End Sub


Sub SaveSet0()
            Open App.Path + "\Data\User\C" + AlphaImageChar.tag + "\set0" For Output As #1
                Print #1, lblTab(4).tag + vbCrLf + CStr(CBoxFlag) + vbCrLf + AlphaImageWeap.tag + vbCrLf + CStr(WeaponBox.ListIndex) + vbCrLf + CStr(RBoxFlag) + vbCrLf + CStr(LevelBox(0).ListIndex) + vbCrLf + CStr(LevelBox(1).ListIndex) + vbCrLf + CStr(LevelBox(2).ListIndex) + vbCrLf + CStr(ListBox1.ListIndex);
            Close #1
End Sub











Private Sub pngMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) '��Ҫ�ڸ����ؼ� MouseDown ������ʹ��
    MoveForm Me '�����ƶ��ޱ߿���
End Sub
Private Sub lblLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'ͬ��
     MoveForm Me '�����ƶ��ޱ߿���
End Sub

Private Sub pngMenu_Click(ByVal Button As Integer) '��ʾ CMenu �˵�
    If Button = 1 Then
        C_Menu.Show '�����ʾ
    Else
        C_Sort.Show '�Ҽ���ʾ
    End If
End Sub

Private Sub pngMinimize_Click(ByVal Button As Integer)
    Me.WindowState = vbMinimized '������С��
End Sub

Private Sub pngMinimize_MouseEnter() '��� Enter �ǿؼ��ڲ��������� Move Ч����ͬ��Enter Ч�ʸ��ߣ�
    pngMinimizeBG.FadeInOut 100, 8
End Sub

Private Sub pngMinimize_MouseExit() 'Exit ͬ�ϣ�Ϊ����Ƴ��ؼ���Χʱ����
    pngMinimizeBG.FadeInOut 0, 8
End Sub

Private Sub pngClose_Click(ByVal Button As Integer) 'Tray ���̿ؼ����ڳ���ر�ʱ���ж��Ƿ�Ϊ����ģʽ
    '��ȡ���� [ ��С�������� ]
    If False Then '���������Ƴɶ�ȡ ini �����ļ���modIni ģ�飬ʹ�÷����뿴�·�ע�ʹ��� ������[ ����дһ��������ģ�飬�ڳ�������ʱ��ȡ���ã�����ȫ�ֱ������Լ�������������϶�Ӧֵ ]
        Me.Hide                       '�����ȡ ini ���� FrmMain.SET_TRAY = ��ȡ�������ԣ�Ȼ����������ر�ʱȡ���ȫ�ֱ��� SET_TRAY �ж�һ�£������ XX ��ô�������������ô����
        Tray1.Show '��ʾ����
        mnuShow.Caption = "��ʾ����"
        If MODEL_TRAY Then '��������ֻ��ʾһ��
            MODEL_TRAY = False
            Tray1.ShowBubble "��ܰ��ʾ", "BSkin Demo ���ں�̨����", NIIF_GUID 'NIIF_NONE ��ͼ�꣬NIIF_INFO ��Ϣͼ�꣬NIIF_WARNING ����ͼ�꣬NIIF_ERROR ����ͼ�꣬NIIF_GUID ���̵�ͼ��
        End If
    Else
        Unload Me '�رմ��壬����رպ���̲��ܽ�������ʹ�� End
    End If
End Sub

'ini �����ļ����ô��루�ж�ѡ�Ȼ�󴢴��Ӧ�Ĳ�����
'If optSetExitEnd.Value = True Then WriteIniParam APP_DATA & INI_SETTING, "Common", "AppExit", "End"
'If optSetExitMin.Value = True Then WriteIniParam APP_DATA & INI_SETTING, "Common", "AppExit", "Min"

'ini �����ļ���ȡ���루��ȡ���ò�����Ȼ����ж�Ӧ���ã�
'optSetExitEnd.Value = IIf(GetIniParam(APP_DATA & INI_SETTING, "Common", "AppExit") = "End", True, False)
'optSetExitMin.Value = IIf(GetIniParam(APP_DATA & INI_SETTING, "Common", "AppExit") = "Min", True, False)

'ini �����ļ��Զ�������������
'[Common]
'AppExit=End


'pngCloseBG.FadeInOut 100, 8 ��һ�������� Opacity ��͸���� 0-100 ֵԽ��͸����Խ�ͣ��ڶ����������ٶ� 1-20 ֵԽ��仯�ٶ�Խ��

Private Sub pngClose_MouseEnter() '�������
    pngCloseBG.FadeInOut 100, 8 'ͼ�񽥱�
End Sub

Private Sub pngClose_MouseExit() '����Ƴ�
    pngCloseBG.FadeInOut 0, 8 'ͼ�񽥱�
End Sub

Private Sub pngTab_Click(Index As Integer, ByVal Button As Integer) '�˵�����ѡ����Ч
    zMove1 AlphaImage1, pngTab(Index).Left - 200, AlphaImage1.Top, True '�ȼ����ƶ�������Ч�ʸ��ߵ� Timer �ؼ�
    
    Dim i As String, j As Integer, tempc As Chars, t As String, YSZJ As Boolean, DUN As Boolean, t1 As String, flag As Boolean
    
    
    For j = 0 To Container1.Count - 1 '����ȫ����������
        With Container1(j)
            .Visible = False
            .BackColor = vbWhite
            .Left = Container1(0).Left 'ÿ���������͵�һ����������
            .Top = Container1(0).Top '������������ҳ��ĵ���
        End With
    Next
            lblTab(12).Visible = False
            pngTab(4).Visible = False
    
    If Index <> 1 Then Unload FrmAbout
    If Index <> 0 Then Unload FrmChar
    
    
    Select Case Index '��ʾ��ѡ�Ĺ����������м���ҳ��� Case �������������������ʾ��Ӧ��ҳ�棩
        Case 0
            Container1(Index).Visible = True 'Index Ϊ��Ӧ�ؼ�������������֪ʶ

        Case 1
            Container1(Index).Visible = True 'ÿ�������ڶ����Լ��Ĵ����߼������� Case ��Ӧ��ʾ���������д�������ķ���

            
        Case 2
            Container1(Index).Visible = True

            'Call LoadListView
        Case 3
            Container1(Index).Visible = True
            If ReloadTip Then
            ClearSelectBuff
            ReloadTip = False
            YSZJ = False
            DUN = False
            If CurrSkill = "c3a3" Or CurrSkill = "c3a4" Or CurrSkill = "c9a4" Or CurrSkill = "c9c2" Then YSZJ = True

            
            
            
            i = CharList(Val(AlphaImageChar.tag), 1)
                Select Case i
                    Case "������"
                        If CurrSkill = "c1e2" Then AddSelectBuff i + "�츳2��ɲ��֮��������ֵС��50%�ĵ�����ɶ����˺�", 2, 0, "Ŀ������ֵС��50%"
                        If InStr(2, CurrSkill, "d") > 0 And CBoxFlag >= 3 Then AddSelectBuff i + "����4�����������������������乥���˺�", 2, 0, "��������������"
                        
                    Case "����"
                        If CurrSkill <> "c3a1" And CurrSkill <> "c3a2" Then AddSelectBuff i + "�츳2������ͥ�����ڼ���ͨ�������к��û�Ԫ���˺��ӳ�", 1, 10, " ��Ч��"
                        If CBoxFlag >= 0 Then AddSelectBuff i + "����1�����������Ӱ���µĵ��˺��ù������ӳ�", 2, , "���������Ӱ���µĵ��˺�"
                        If CBoxFlag >= 1 Then AddSelectBuff i + "����2����Ԫ���˺���ɱ������û�Ԫ���˺��ӳ�", 2, , "��Ԫ���˺���ɱ�����"
                    Case "����"
                        AddSelectBuff i + "�츳2����׼�������������ù������ӳ�", 2, 0, "��׼������������"
                    Case "�Ű���"
                        AddSelectBuff i + "����2������E����״̬ʱ���ˮԪ���˺��ӳ�", 2, 0, "���ڸ���֮����"
                    Case "����"
                        If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "�츳2���ͷ�����˺�E������ͨ����/�ػ��˺��ӳ�", 2, 0, "ʩ������˺�E��"
                    Case "����"
                            AddSelectBuff i + "���Ƿ���Ԫ��ս��������˰�����״̬��", 2, 0, "ʩ��Ԫ��ս����"
                            YSZJ = True
                            BuffCheck(SelectCount).Value = Checked
                            AddSelectBuff i + "�츳3��Ѫ������50%ʱ��û�Ԫ���˺��ӳ�", 2, 0, "����ֵ����50%"
                    Case "��"
                        AddSelectBuff i + "�����ھ�������ڼ���", 1, 15, " ��������(��δ��������Ϊ0)"
                        SelectBuffBar(SelectCount).Cls
                        SelectBuffBar(SelectCount).tag = "1"
                        Call SelectBuffBar_MouseMove(SelectCount, 1, 0, 15, 1)
                        SelectBuffBar(SelectCount).tag = "0"
                    Case "��¬��"
                        AddSelectBuff i + "���Ƿ���Ԫ�ر����������ħ״̬��", 2, 0, "ʩ��Ԫ�ر�����"
                        BuffCheck(SelectCount).Value = Checked
                        If CBoxFlag >= 0 Then AddSelectBuff i + "����1��������ֵ����50%�ĵ�����ɶ����˺�", 2, 0, "Ŀ������ֵ����50%"
                        If CBoxFlag >= 1 Then AddSelectBuff i + "����2���ܻ�����߹�����", 1, 3, " ��Ч��%"
                        If CBoxFlag >= 3 And InStr(2, CurrSkill, "e") > 0 Then AddSelectBuff i + "����4���н����ʩ��Ԫ��ս��ʱ������˺�", 2, , "ʩ��Ԫ��ս����": YSZJ = True
                        If CBoxFlag >= 5 And InStr(2, CurrSkill, "a") > 0 Then AddSelectBuff i + "����6��ʩ��Ԫ��ս���������ͨ�����˺�", 2, , "ʩ��Ԫ��ս����": YSZJ = True
                        
                        
                    Case "����类�"
                        AddSelectBuff i + "���Ƿ�ʹ�������̣��������ħ״̬�����е���", 2, 0, "ʹ�����������е���"
                        BuffCheck(SelectCount).Value = Checked
                        If CBoxFlag = 5 And InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "����6�������ػ��Ƿ񴥷�������̤", 2, 0, "������̤����"
                        If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "�츳2��ʹ��Ԫ��ս���������ͨ����/�ػ��˺�", 2, 0, "ʩ��Ԫ��ս����": YSZJ = True
                    
                    Case "�׵罫��"
                        If InStr(2, CurrSkill, "q") > 0 Then AddSelectBuff i + "����Ը����֮�ֵ�Ը������", 1, 60, " ��Ը��"
                    
                    
                End Select
                
            
            
            
            i = WeaponList(Val(AlphaImageWeap.tag), 1)
                Select Case i
                    Case "����"
                        AddSelectBuff "���𰣺���е��˺������������", 1, 7, "��Ч��"
                    Case "����֮ǹ"
                        AddSelectBuff "����֮ǹ������е���ʱ�������", 1, 2, "��ߵ�������"
                    Case "ϻ������"
                        If InStr(2, CurrSkill, "q") > 0 Or InStr(2, CurrSkill, "e") > 0 Then AddSelectBuff "ϻ�����£���Ч����", 2, , "��ͨ�������е��˺�"
                        If InStr(2, CurrSkill, "a") > 0 Then AddSelectBuff "ϻ�����£���Ч����", 2, , "Ԫ��ս��/Ԫ�ر������е��˺�"
                    Case "ǧ�ҹŽ�"
                        AddSelectBuff "ǧ�ҹŽ��������������³�Աʱ�������", 1, 4, "�������³�Ա����"
                    Case "ǧ�ҳ�ǹ"
                        AddSelectBuff "ǧ�ҳ�ǹ�������������³�Աʱ�������", 1, 4, "�������³�Ա����"
                    Case "������"
                        AddSelectBuff "�����򣺴���ˮ���Ԫ�ط�Ӧ�����ӹ������ӳ�", 2, , "����ˮ��ط�Ӧ"
                    Case "�������"
                        AddSelectBuff "������񣺻��ܵ��˺��ù������ӳ�", 1, 3, "���ܵ�������"
                    Case "����ͼ��"
                        AddSelectBuff "����ͼ�ף�����Ԫ�ط�Ӧ����Ԫ���˺��ӳ�", 1, 2, "����Ԫ�ط�Ӧ"
                    Case "�����ط�¼"
                        AddSelectBuff "�����ط�¼������˺��󣬶ѵ�������", 1, 5, "�ѵ�����"
                    Case "���Ҵ�"
                        AddSelectBuff "���Ҵ󽣣�����˺��󣬶ѵ�������", 1, 5, "�ѵ�����"
                    Case "������ǹ"
                        AddSelectBuff "������ǹ������˺��󣬶ѵ�������", 1, 5, "�ѵ�����"
                    Case "���ҳ���"
                        AddSelectBuff "���ҳ���������˺��󣬶ѵ�������", 1, 5, "�ѵ�����"
                    Case "���ҳ���"
                        AddSelectBuff "���ҳ���������˺��󣬶ѵ�������", 1, 5, "�ѵ�����"
                    Case "����ľ���ʫ"
                        AddSelectBuff "����ľ���ʫ����̺����ӹ�����", 2, 0, "���/�����̺�"
                    Case "�ཿɹ��¼�"
                        If InStr(2, CurrSkill, "c") > 0 Then
                            AddSelectBuff "�ཿɹ��¼�����ͨ����/�ػ����к�������", 2, 2, "�ػ����е��˺�", "��ͨ�������е��˺�"
                        Else
                            AddSelectBuff "�ཿɹ��¼����ػ����к��ù������ӳ�", 2, , "�ػ����е��˺�"
                        End If
                    Case "�ķ�ԭ��"
                        AddSelectBuff "�ķ�ԭ�䣺�����ڳ�ʱ���Ԫ���˺��ӳ�", 1, 4, "��Ч��"
                    Case "����֮��"
                        AddSelectBuff "����֮�������е��˺��ù������ӳ�", 1, 5, "��Ч��", "���ڻ���״̬��": DUN = True
                    Case "��Ӱ����"
                        If InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff "��Ӱ����������ֵ����70%ʱ����ػ��˺��ӳ�", 2, , "����ֵ����70%"
                    Case "���������"""
                        AddSelectBuff "�������������ͨ�������ػ����е��˺��ù������ӳ�", 1, 4, "��Ч��"
                    Case "����ն��"
                        AddSelectBuff "����ն�������ܵ��˺��ù������ӳ�", 1, 3, "���ܵ�������"
                    Case "���Ҵ�ǹ"
                        AddSelectBuff "���Ҵ�ǹ�����ܵ��˺��ù������ӳ�", 1, 3, "���ܵ�������"
                    Case "����ս��"
                        AddSelectBuff "����ս�������ܵ��˺��ù������ӳ�", 1, 3, "���ܵ�������"
                    Case "���ҳ���"
                        AddSelectBuff "���ҳ��������ܵ��˺��ù������ӳ�", 1, 3, "���ܵ�������"
                   Case "�ǽ�"
                        AddSelectBuff "�ǽ��������ڳ�ʱ����˺��ӳ� ", 1, 5, "��Ч��"
                   Case "��Ӱ��"
                        AddSelectBuff "��Ӱ�������е��˺��ù������ͷ������ӳ� ", 1, 4, "��Ч��"
                   Case "�ǵ�ĩ·"
                        AddSelectBuff "�ǵ�ĩ·����������ֵС��30%�ĵ��˺��ù������ӳ� ", 2, , "��������ֵС��30%�ĵ��˺�"
                   
                   Case "�޹�֮��"
                        AddSelectBuff "�޹�֮�������е��˺��ù������ӳ� ", 1, 5, "��Ч��", "���ڻ���״̬��": DUN = True
                   
                   Case "��������֮ʱ"
                        AddSelectBuff "��������֮ʱ��������Ч���ù������͹��ټӳ� ", 2, , "����֮�败��"
                   Case "������֮��"
                        AddSelectBuff "������֮�ģ�������Ҫ����ɵ��˺�", 2, , "���Ҫ��"
                   Case "����"
                        If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "���������е���ʱ����˺��������˺�����", 2, , "0.3��������"
                   Case "��������"
                        AddSelectBuff i + "������λ�ں�̨ʱ����˺��ӳ�", 1, 10, "��Ч��"
                   Case "���ֹ�"
                        AddSelectBuff i + "����ͨ����/�ػ����е��˺��ù������͹��ټӳ�", 1, 4, "��Ч��"
                   Case "�������"
                        AddSelectBuff i + "���ػ�����Ҫ�����ù����������ټӳ�", 2, , "����Ҫ����"
                  Case "�绨֮��"
                    If InStr(2, CurrSkill, "e") = 0 Then AddSelectBuff i + "��ʩ��Ԫ��ս�����ù������ӳ�", 2, , "ʩ��Ԫ��ս����": YSZJ = True
                  Case "��ҹ������"
                      If InStr(2, CurrSkill, "e") > 0 Then AddSelectBuff i + "����ͨ�������к���Ԫ��ս���˺��ӳ�", 2, , "��ͨ�������к�"
                      If InStr(2, CurrSkill, "a") > 0 Then AddSelectBuff i + "��Ԫ��ս�����к�����ͨ�����˺��ӳ�", 2, , "Ԫ��ս�����к�": YSZJ = True
                  Case "��Ī˹֮��"
                        If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "����ʸ�����ʱ������˺�", 1, 5, "��������������"
                   Case "��ĩ�֮̾ʫ"
                        AddSelectBuff i + "��������Ч���ù�������Ԫ�ؾ�ͨ�ӳ� ", 2, , "����֮�败��"
                   Case "����֮����"
                        AddSelectBuff i + "����÷���֮��ӡ�������ͨ�����˺�", 1, 3, "����֮��ӡ����"
                   Case "������"
                         AddSelectBuff i + "�����������Ԫ�ط�Ӧ�����ӹ������ӳ�", 2, , "��������ط�Ӧ"
                   Case "������"
                        AddSelectBuff i + "������ֵ����70%ʱ��ñ����ʼӳ�", 2, , "����ֵ����70%"
                    
                   Case "��������"
                        AddSelectBuff i + "��ʩ��Ԫ�ر������ù����������ټӳ�", 2, , "ʩ��Ԫ�ر�����"
                   Case "�����"
                        AddSelectBuff i + "�����Ԫ���˺������˺��ӳ�", 1, 2, "��Ч��"
                   Case "��Ӱ��"
                        AddSelectBuff i + "�����е��˺��ù������ͷ������ӳ� ", 1, 4, "��Ч��"
                   
                   Case "��֮��"
                        AddSelectBuff i + "�����е��˺��ù������ӳ� ", 1, 5, "��Ч��", "���ڻ���״̬��": DUN = True
                        
                   Case "�Թ�����֮��"
                        AddSelectBuff i + "��������Ч���ù��������˺��ӳ� ", 2, , "����֮�败��"
                        
                   Case "����֮�ع�"
                        AddSelectBuff i + "���������֮��ӡ�����Ԫ���˺��ӳ�", 1, 3, "����֮��ӡ����"
                   Case "������"
                        If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "�����Ԫ��΢��������ͨ����/�ػ��˺��ӳ�", 2, , "���Ԫ��΢����"
                   Case "��������"
                        If InStr(2, CurrSkill, "a") > 0 Or InStr(2, CurrSkill, "c") > 0 Then AddSelectBuff i + "��ʩ��Ԫ��ս��������ͨ����/�ػ��˺��ӳ�", 1, 2, "ʩ��Ԫ��ս������"
                   Case "���֮��"
                        AddSelectBuff i + "�����е��˺��ù������ӳ� ", 1, 5, "��Ч��", "���ڻ���״̬��": DUN = True
                        
                   Case "�S��֮����"
                        AddSelectBuff i + "��ʩ��Ԫ�ر�������Ԫ�س���Ч�ʼӳ� ", 2, , "ʩ��Ԫ�ر�����"
                        If InStr(2, CurrSkill, "q") > 0 Then BuffCheck(SelectCount).Value = Checked
                        
                   Case "�׳�֮��"
                        AddSelectBuff i + "�����������Ԫ�ط�Ӧ����Ԫ���˺��ӳ� ", 2, , "��������ط�Ӧ"
                        
                   Case "��������"
                        AddSelectBuff i + "���ǳ�ʱ��üӳɣ�1����������2��Ԫ���˺���3��Ԫ�ؾ�ͨ�� ", 1, 3, " ������һ��Ч��"
                        
                   Case "��Ħ֮��"
                    flag = False
                        For j = 1 To SelectCount
                            If BuffCheck(j).Caption = "����ֵ����50%" Then flag = True
                        Next
                        If flag = False Then AddSelectBuff i + "������ֵ����50%ʱ��ö��⹥�����ӳ� ", 2, , "����ֵ����50%"
                End Select

            t = ""
            For j = 1 To SetCount
                t = t + AddArt(tempc, j)
            Next
            
            If InStr(1, t, "���4") > 0 And DUN = False Then
                AddSelectBuff "��ɵ������ļ��ף����ڻ���״̬ʱ������� ", 2, , " ���ڻ���״̬��"
            End If
            
            If InStr(1, t, "��ʿ4") > 0 And InStr(2, CurrSkill, "c") > 0 Then
                AddSelectBuff "ȾѪ����ʿ���ļ��ף����ܵ��˺��ػ�����˺��ӳ� ", 2, , " ���ܵ��˺�"
            End If
            
            
            If InStr(1, t, "ħŮ4") > 0 Then
                AddSelectBuff "ʥ������ҵ���֮ħŮ��ʩ��Ԫ��ս������ ", 1, 3, " ��"
            End If
            
            If InStr(1, t, "�԰�4") > 0 Then
                AddSelectBuff "�԰�֮���ļ��ף�ʩ��Ԫ��ս��������˺����� ", 1, 2, " ��"
            End If
            
            If (InStr(1, t, "ˮ��4") > 0 Or InStr(1, t, "׷��4") > 0) And YSZJ = False And InStr(1, t, "ħŮ4") <= 0 Then
                t1 = ""
                If InStr(1, t, "ˮ��4") > 0 Then t1 = "����֮��"
                If InStr(1, t, "׷��4") > 0 Then
                    If t1 = "" Then
                        t1 = "����֮��"
                    Else
                        t1 = t1 + "/׷��֮ע��"
                    End If
                End If
                AddSelectBuff "ʥ����" + t1 + "��ʩ��Ԫ��ս���������� ", 2, , " ʩ��Ԫ��ս����"
            End If


            End If


            
            
            

        Case 4
            Container1(Index).Visible = True
            lblTab(12).Visible = True
            pngTab(4).Visible = True
    End Select
End Sub
Private Sub ClearSelectBuff()
On Error GoTo Outs
Dim i%
            For i = 1 To 100
                 FrmMain.SelectBuffBox(i).Visible = False
            Next
Outs:
            SelectCount = 0
End Sub
Private Sub AddSelectBuff(tip As String, mode As Integer, Optional Count As Integer, Optional tip2 As String, Optional tip3 As String)
On Error GoTo Outs
    SelectCount = SelectCount + 1
    Load SelectBuffBox(SelectCount)
    SelectBuffBox(SelectCount).Left = SelectBuffBox(0).Left
    SelectBuffBox(SelectCount).Top = 480 + (SelectCount - 1) * (SelectBuffBox(SelectCount).Height)
    SelectBuffBox(SelectCount).Visible = True
    
    Load SelectBuffLabel(SelectCount)
    Set SelectBuffLabel(SelectCount).Container = SelectBuffBox(SelectCount)
    SelectBuffLabel(SelectCount).Left = SelectBuffLabel(0).Left
    SelectBuffLabel(SelectCount).Top = SelectBuffLabel(0).Top
    SelectBuffLabel(SelectCount).Visible = True
    
    
    Load SelectBuffBar(SelectCount)
    Set SelectBuffBar(SelectCount).Container = SelectBuffBox(SelectCount)
    SelectBuffBar(SelectCount).Left = SelectBuffBar(0).Left
    SelectBuffBar(SelectCount).Top = SelectBuffBar(0).Top
    

    Load BuffLabel(SelectCount)
    Set BuffLabel(SelectCount).Container = SelectBuffBox(SelectCount)
    BuffLabel(SelectCount).Left = BuffLabel(0).Left
    BuffLabel(SelectCount).Top = BuffLabel(0).Top
    

    Load BuffCheck(SelectCount)
    Set BuffCheck(SelectCount).Container = SelectBuffBox(SelectCount)
    BuffCheck(SelectCount).Left = BuffCheck(0).Left
    BuffCheck(SelectCount).Top = BuffCheck(0).Top
    
    Load BuffCheck2(SelectCount)
    Set BuffCheck2(SelectCount).Container = SelectBuffBox(SelectCount)
    BuffCheck2(SelectCount).Left = BuffCheck2(0).Left
    BuffCheck2(SelectCount).Top = BuffCheck2(0).Top
    
Outs:
    SelectBuffBox(SelectCount).Visible = True
    SelectBuffLabel(SelectCount).Caption = tip
    If mode = 1 Then '����ģʽ
        SelectBuffBar(SelectCount).Visible = True
        SelectBuffBar(SelectCount).LinkTimeout = Count
        BuffLabel(SelectCount).Visible = True
        BuffCheck(SelectCount).Visible = False
        BuffCheck2(SelectCount).Visible = False
        BuffLabel(SelectCount).Caption = "0/" + CStr(Count) + " " + tip2
        SelectBuffBar(SelectCount).Cls
        SelectBuffLabel(SelectCount).tag = tip2
        If IsMissing(tip3) = False And tip3 <> "" Then
            BuffCheck2(SelectCount).Visible = True
            BuffCheck2(SelectCount).Caption = tip3
        End If
    Else
        SelectBuffBar(SelectCount).Visible = False
        BuffLabel(SelectCount).Visible = False
        BuffCheck(SelectCount).Visible = True
        BuffCheck(SelectCount).Caption = tip2
        If Count = 2 Then
            BuffCheck2(SelectCount).Visible = True
            BuffCheck2(SelectCount).Caption = tip3
        End If
    End If

End Sub




























































'ȡɫģ��


'��ȡHEX
Private Function GetHex(intVal As Long) As String
On Error GoTo ErrExit
    Dim strHex As String
    
    strHex = Hex(intVal)
    If Len(strHex) = 1 Then strHex = "0" & strHex
    GetHex = strHex
ErrExit:
End Function

'Tray ���̿ؼ�����Ҫ���������Ϸ��� Tray �ؼ������� �˵��༭�� ����������˵�ѡ��
Private Sub Tray1_PopupMenu() '���̲˵�
    PopupMenu mnuApp '�������õĲ˵����˵��༭�� ���޸ģ�
End Sub

Private Sub Tray1_Click() '���̵��������������ֱ����ʾ���壬�Ҽ��������ʾ�˵�����Ӧ�Ĳ˵������� �˵��༭�� ���޸ģ����� ���˵��� ��ѡ����Ҫ�� �ɼ� ��ѡȥ����
    Me.WindowState = vbNormal
    Me.Show
    mnuShow.Caption = "���ؽ���"
End Sub

Private Sub mnuShow_Click() '��ʾ���ؽ���
    If mnuShow.Caption = "���ؽ���" Then
        Me.Hide
        Tray1.Show
        mnuShow.Caption = "��ʾ����"
    Else
        Me.WindowState = vbNormal
        Me.Show
        mnuShow.Caption = "���ؽ���"
    End If
End Sub

Private Sub mnuSetting_Click() '��������
    'FrmSetting.Show , Me 'ĳ������ҳ��
End Sub

Private Sub mnuExit_Click() '�˳�����
    Unload Me
End Sub


'�˵�����
Private Sub C_Sort_MenuClick(ByVal MenuIndex As Long)
    Debug.Print "�����˵� [" & MenuIndex & "] �KeyΪ [" & C_Sort.Key(MenuIndex) & "] "
End Sub

'�˵�����
Private Sub C_Menu_MenuClick(ByVal MenuIndex As Long)
    Select Case C_Menu.Key(MenuIndex)
        Case "setting"
            'ϵͳ����
        Case "download"
            '�½�����
        Case "unpack"
            '�ļ���ѹ
        Case "update"
            '������
        Case "about"
            FrmAbout.Show , Me '���ڳ���
        Case "exit"
            Unload Me '�˳�����
    End Select
End Sub

'�����Ҽ��˵�
Private Sub C_Subclass_SubclassProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lhWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
    Select Case uMsg
        Case WM_TASKMENU
            bHandled = True
            lReturn = 0
            C_Sort.Show
    End Select
End Sub

