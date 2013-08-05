VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atributo de Sonido"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   2895
   ControlBox      =   0   'False
   Icon            =   "frmSound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8281
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
      TabCaption(0)   =   "Sonido"
      TabPicture(0)   =   "frmSound.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstSound"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   3600
         Width           =   1815
         _ExtentX        =   3201
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
         PictureNormal   =   "frmSound.frx":0FDE
         PictureHot      =   "frmSound.frx":17C2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdCancel 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   3960
         Width           =   1815
         _ExtentX        =   3201
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
         PictureNormal   =   "frmSound.frx":1FA6
         PictureHot      =   "frmSound.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    SoundFileName = lstSound.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Call ListSounds(App.Path & "\SFX\", 1)
    lstSound.Text = SoundFileName
End Sub

Private Sub lstSound_Click()
    Call PlaySound(lstSound.Text)
End Sub
