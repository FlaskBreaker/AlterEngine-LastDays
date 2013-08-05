VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNotice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atributo de Noticia"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmNotice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Noticia"
      TabPicture(0)   =   "frmNotice.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Title"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lstSound"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.ListBox lstSound 
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
         Height          =   3150
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Title 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1200
         Width           =   4335
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   4080
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
         PictureNormal   =   "frmNotice.frx":0FDE
         PictureHot      =   "frmNotice.frx":17C2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdCancel 
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   4560
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
         PictureNormal   =   "frmNotice.frx":1FA6
         PictureHot      =   "frmNotice.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Titulo:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texto:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sonido:"
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
         Left            =   210
         TabIndex        =   3
         Top             =   1560
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    NoticeTitle = Title.Text
    NoticeText = Text.Text
    NoticeSound = lstSound.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Call ListSounds(App.Path & "\SFX\", 2)

    Title.Text = NoticeTitle
    Text.Text = NoticeText
    lstSound.Text = NoticeSound
End Sub

Private Sub lstSound_Click()
    Call PlaySound(lstSound.Text)
End Sub
