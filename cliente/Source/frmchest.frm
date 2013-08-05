VERSION 5.00
Begin VB.Form frmChest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Cofres"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmchest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox SSTab1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   840
         Max             =   500
         Min             =   1
         TabIndex        =   6
         Top             =   1320
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   840
         Max             =   500
         Min             =   1
         TabIndex        =   1
         Top             =   960
         Value           =   1
         Width           =   3255
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1920
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
         PictureNormal   =   "frmchest.frx":0FC2
         PictureHot      =   "frmchest.frx":17A6
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
         Top             =   1920
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
         PictureNormal   =   "frmchest.frx":1F8A
         PictureHot      =   "frmchest.frx":28DE
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblamount 
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
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Width           =   615
      End
      Begin VB.Label lblItem 
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
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Objeto"
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
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmChest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    scrlItem.max = MAX_ITEMS
    lblName.Caption = Trim$(Item(scrlItem.value).name)

    If ChestItemNum < scrlItem.min Then
        ChestItemNum = scrlItem.min
    End If
    
    If ChestItemAmount < scrlAmount.min Then
        ChestItemAmount = scrlAmount.min
    End If
    
    scrlItem.value = ChestItemNum
    scrlAmount.value = ChestItemAmount
End Sub

Private Sub cmdOk_Click()
    ChestItemNum = scrlItem.value
    ChestItemAmount = scrlAmount.value
    Unload Me
End Sub

Private Sub scrlAmount_Change()
    lblAmount.Caption = STR(scrlAmount.value)
End Sub

Private Sub scrlItem_Change()
    lblItem.Caption = STR(scrlItem.value)
    lblName.Caption = Trim$(Item(scrlItem.value).name)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

