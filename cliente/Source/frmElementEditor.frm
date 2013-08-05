VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmElementEditor 
   Caption         =   "Editor de Elementos"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   ScaleHeight     =   3315
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   397
      TabMaxWidth     =   1852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Elementos"
      TabPicture(0)   =   "frmElementEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblStrong"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblWeak"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlStrong"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlWeak"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.HScrollBar scrlWeak 
         Height          =   255
         Left            =   840
         Max             =   1000
         TabIndex        =   6
         Top             =   2040
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrong 
         Height          =   255
         Left            =   840
         Max             =   1000
         TabIndex        =   3
         Top             =   1320
         Value           =   1
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Aceptar"
         PictureNormal   =   "frmElementEditor.frx":001C
         PictureHot      =   "frmElementEditor.frx":0800
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton Command1 
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Cancelar"
         PictureNormal   =   "frmElementEditor.frx":0FE4
         PictureHot      =   "frmElementEditor.frx":1938
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblWeak 
         AutoSize        =   -1  'True
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1680
         TabIndex        =   8
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Debil contra:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   7
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblStrong 
         AutoSize        =   -1  'True
         Caption         =   "Nada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fuerte contra:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmElementEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ElementEditorOk
End Sub
Private Sub Command1_Click()
    Call ElementEditorCancel
End Sub

Private Sub Form_Load()
    scrlStrong.max = MAX_ELEMENTS
    scrlWeak.max = MAX_ELEMENTS
End Sub

Private Sub scrlStrong_Change()
    lblStrong.Caption = Element(scrlStrong.value).name
End Sub

Private Sub scrlWeak_Change()
    lblWeak.Caption = Element(scrlWeak.value).name
End Sub
