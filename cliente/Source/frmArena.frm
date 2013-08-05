VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmArena 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atributos de Arena"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmArena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3836
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   441
      TabMaxWidth     =   2822
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Añadir el spawn"
      TabPicture(0)   =   "frmArena.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblY"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblX"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMap"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrlY"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlX"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlMap"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   2520
         Max             =   30
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
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
         PictureNormal   =   "frmArena.frx":0FDE
         PictureHot      =   "frmArena.frx":17C2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdCancel 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
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
         PictureNormal   =   "frmArena.frx":1FA6
         PictureHot      =   "frmArena.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         Caption         =   "Mapa: 0"
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
         TabIndex        =   6
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X: 0"
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
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         Caption         =   "Y: 0"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlMap.value = 0
    scrlX.value = 0
    scrlY.value = 0

    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    scrlMap.max = MAX_MAPS
    scrlX.max = MAX_MAPX
    scrlY.max = MAX_MAPY

    If Arena1 < scrlMap.min Then
        Arena1 = scrlMap.min
    End If

    scrlMap.value = Arena1

    If Arena2 < scrlX.min Then
        Arena2 = scrlX.min
    End If

    scrlX.value = Arena2

    If Arena3 < scrlY.min Then
        Arena3 = scrlY.min
    End If

    scrlY.value = Arena3
End Sub

Private Sub scrlMap_Change()
    lblMap.Caption = "Mapa: " & scrlMap.value
    Arena1 = scrlMap.value
End Sub

Private Sub scrlX_Change()
    lblX.Caption = "X: " & scrlX.value
    Arena2 = scrlX.value
End Sub

Private Sub scrlY_Change()
    lblY.Caption = "Y: " & scrlY.value
    Arena3 = scrlY.value
End Sub
