VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSpriteChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atributo de Cambiar Script"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmSpriteChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
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
      TabCaption(0)   =   "Sprite"
      TabPicture(0)   =   "frmSpriteChange.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSprite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCost"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblItem"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancel"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlSprite"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlCost"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlItem"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "picSprite"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   600
         Width           =   480
         Begin VB.PictureBox picSprites 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   11
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   360
         Max             =   30
         TabIndex        =   7
         Top             =   1560
         Width           =   4335
      End
      Begin VB.HScrollBar scrlCost 
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         Max             =   30000
         TabIndex        =   2
         Top             =   2400
         Width           =   4335
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   2760
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
         PictureNormal   =   "frmSpriteChange.frx":0FDE
         PictureHot      =   "frmSpriteChange.frx":17C2
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
         TabIndex        =   13
         Top             =   2760
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
         PictureNormal   =   "frmSpriteChange.frx":1FA6
         PictureHot      =   "frmSpriteChange.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "Sin Coste"
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
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Objeto:"
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
         Left            =   510
         TabIndex        =   8
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
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
         Height          =   165
         Left            =   1080
         TabIndex        =   6
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   615
         TabIndex        =   5
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sprite:"
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
         Left            =   585
         TabIndex        =   4
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
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
         Height          =   165
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmSpriteChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmSpriteChange.Visible = False
End Sub

Private Sub cmdOk_Click()
    SpritePic = scrlSprite.value
    SpriteItem = scrlItem.value
    SpritePrice = scrlCost.value
    scrlCost.value = 0
    scrlSprite.value = 0
    scrlItem.value = 0
    frmSpriteChange.Visible = False
End Sub

Private Sub Form_Load()

    If SpritePic < scrlSprite.min Then
        SpritePic = scrlSprite.min
    End If
    scrlSprite.value = SpritePic
    If SpriteItem < scrlItem.min Then
        SpriteItem = scrlItem.min
    End If
    scrlItem.value = SpriteItem
    If SpritePrice < scrlCost.min Then
        SpritePrice = scrlCost.min
    End If
    scrlCost.value = SpritePrice

    If SpriteSize = 1 Then
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.top = (scrlSprite.value * 64) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.value * 64, SRCCOPY)
    Else
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.top = (scrlSprite.value * PIC_Y) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.value * PIC_Y, SRCCOPY)
    End If

End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = scrlCost.value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.value = 0 Then
        lblItem.Caption = "Sin Coste"
        Exit Sub
    Else
        lblItem.Caption = scrlItem.value & " - " & Trim$(Item(scrlItem.value).name)
    End If

    If Item(scrlItem.value).Type = ITEM_TYPE_CURRENCY Then
        scrlCost.Enabled = True
    Else
        scrlCost.Enabled = False
    End If
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = scrlSprite.value
    If SpriteSize = 1 Then
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.top = (scrlSprite.value * 64) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.value * 64, SRCCOPY)
    Else
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.top = (scrlSprite.value * PIC_Y) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.value * PIC_Y, SRCCOPY)
    End If
End Sub

Private Sub Timer1_Timer()
    If SpriteSize = 1 Then
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.top = (scrlSprite.value * 64) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.value * 64, SRCCOPY)
    Else
        frmSpriteChange.picSprites.Left = (3 * PIC_X) * -1
        frmSpriteChange.picSprites.top = (scrlSprite.value * PIC_Y) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, 3 * PIC_X, scrlSprite.value * PIC_Y, SRCCOPY)
    End If
End Sub
