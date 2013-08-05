VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEditArrows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Flechas"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frmEditArrows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   5847
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Editar Flecha"
      TabPicture(0)   =   "frmEditArrows.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRange"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblArrow"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAmount"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOk"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlArrow"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Picture1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlRange"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlAmount"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   2520
         Max             =   500
         Min             =   1
         TabIndex        =   10
         Top             =   1800
         Value           =   1
         Width           =   2535
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2055
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   120
         Max             =   30
         Min             =   1
         TabIndex        =   7
         Top             =   1800
         Value           =   1
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3720
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   2
         Top             =   360
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   3
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picArrows 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   4
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.HScrollBar scrlArrow 
         Height          =   255
         Left            =   2520
         Max             =   500
         Min             =   1
         TabIndex        =   1
         Top             =   960
         Value           =   1
         Width           =   2535
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   2640
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
         PictureNormal   =   "frmEditArrows.frx":0FDE
         PictureHot      =   "frmEditArrows.frx":17C2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton Command1 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   2640
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
         PictureNormal   =   "frmEditArrows.frx":1FA6
         PictureHot      =   "frmEditArrows.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblAmount 
         Caption         =   "Cantidad:"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblArrow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flecha:"
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
         TabIndex        =   6
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Rango:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   795
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   960
      TabIndex        =   11
      Top             =   1800
      Width           =   45
   End
End
Attribute VB_Name = "frmEditArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ArrowEditorOk
End Sub

Private Sub Command1_Click()
    Call ArrowEditorCancel
End Sub

Private Sub scrlArrow_Change()
    lblArrow.Caption = "Flecha: " & scrlArrow.value
    frmEditArrows.picArrows.top = (scrlArrow.value * 32) * -1
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Rango: " & scrlRange.value
End Sub
Private Sub scrlAmount_Change()
    lblAmount.Caption = "Cantidad: " & scrlAmount
End Sub
