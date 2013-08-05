VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapWarp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mover jugador a..."
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "frmMapWarp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
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
      TabCaption(0)   =   "Mover"
      TabPicture(0)   =   "frmMapWarp.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblX"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlX"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlY"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMap"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.TextBox txtMap 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Text            =   "1"
         Top             =   480
         Width           =   3855
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   720
         Max             =   30
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   720
         Max             =   30
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   1800
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
         PictureNormal   =   "frmMapWarp.frx":0FDE
         PictureHot      =   "frmMapWarp.frx":17C2
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
         TabIndex        =   10
         Top             =   1800
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
         PictureNormal   =   "frmMapWarp.frx":1FA6
         PictureHot      =   "frmMapWarp.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblY 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblX 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mapa :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMapWarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    EditorWarpMap = Val(txtMap.Text)
    EditorWarpX = scrlX.value
    EditorWarpY = scrlY.value
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    scrlX.max = MAX_MAPX
    scrlY.max = MAX_MAPY

    If EditorWarpX < scrlX.min Then
        EditorWarpX = scrlX.min
    End If
    scrlX.value = EditorWarpX
    If EditorWarpY < scrlY.min Then
        EditorWarpY = scrlY.min
    End If
    scrlY.value = EditorWarpY
    txtMap.Text = EditorWarpMap
End Sub

Private Sub scrlX_Change()
    lblX.Caption = STR(scrlX.value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = STR(scrlY.value)
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.Text) <= 0 Or Val(txtMap.Text) > MAX_MAPS Then
        txtMap.Text = "1"
    End If
End Sub

