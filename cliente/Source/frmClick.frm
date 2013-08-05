VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmClick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atributo: Un Click"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2566
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   3175
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Seleccionar Script"
      TabPicture(0)   =   "frmClick.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblScript"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlClick"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.HScrollBar scrlClick 
         Height          =   255
         Left            =   360
         Max             =   100
         TabIndex        =   1
         Top             =   600
         Width           =   4095
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   960
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
         PictureNormal   =   "frmClick.frx":001C
         PictureHot      =   "frmClick.frx":0800
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
         TabIndex        =   5
         Top             =   960
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
         PictureNormal   =   "frmClick.frx":0FE4
         PictureHot      =   "frmClick.frx":1938
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblScript 
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
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Script:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ClickScript = scrlClick.value
    Unload Me
End Sub

Private Sub Form_Load()
    If ClickScript < scrlClick.min Then
        ClickScript = scrlClick.min
    End If
    scrlClick.value = ClickScript
    lblScript.Caption = scrlClick.value
    SendScriptTile ClickScript
End Sub

Private Sub scrlClick_Change()
    lblScript.Caption = scrlClick.value
    Call SendScriptTile(ClickScript)
End Sub
