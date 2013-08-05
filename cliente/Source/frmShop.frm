VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmShop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tienda"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "frmShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
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
      TabCaption(0)   =   "Establecer"
      TabPicture(0)   =   "frmShop.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNum"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlNum"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   3
         Top             =   720
         Value           =   1
         Width           =   4335
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   1200
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
         PictureNormal   =   "frmShop.frx":0FDE
         PictureHot      =   "frmShop.frx":17C2
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
         Top             =   1200
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
         PictureNormal   =   "frmShop.frx":1FA6
         PictureHot      =   "frmShop.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tienda Numero:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlNum.value = 1
    Unload Me
End Sub

Private Sub cmdOk_Click()
    EditorShopNum = scrlNum.value
    scrlNum.value = 1
    Unload Me
End Sub

Private Sub Form_Load()
    lblNum.Caption = scrlNum.value & " - " & Trim$(Shop(scrlNum.value).name)
    If EditorShopNum < scrlNum.min Then
        EditorShopNum = scrlNum.min
    End If
    scrlNum.value = EditorShopNum
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = scrlNum.value & " - " & Trim$(Shop(scrlNum.value).name)
End Sub
