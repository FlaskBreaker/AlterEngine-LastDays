VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecciona para editar"
   ClientHeight    =   3675
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   5550
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
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstIndex 
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
      Height          =   2340
      ItemData        =   "frmIndex.frx":0CCA
      Left            =   240
      List            =   "frmIndex.frx":0CCC
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6165
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
      TabCaption(0)   =   "Selección"
      TabPicture(0)   =   "frmIndex.frx":0CCE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   2880
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
         PictureNormal   =   "frmIndex.frx":0CEA
         PictureHot      =   "frmIndex.frx":14CE
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
         TabIndex        =   3
         Top             =   2880
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
         PictureNormal   =   "frmIndex.frx":1CB2
         PictureHot      =   "frmIndex.frx":2606
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    EditorIndex = lstIndex.ListIndex + 1

    If InItemsEditor Then
        Call SendData("edititem" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InNpcEditor Then
        Call SendData("editnpc" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InShopEditor Then
        Call SendData("editshop" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InElementEditor Then
        Call SendData("editelement" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InSpellEditor Then
        Call SendData("editspell" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InEmoticonEditor Then
        Call SendData("editemoticon" & SEP_CHAR & EditorIndex & END_CHAR)
    End If

    If InArrowEditor Then
        Call SendData("editarrow" & SEP_CHAR & EditorIndex & END_CHAR)
    End If
    
    If InQuestEditor = True Then
        Call SendData("EDITQUEST" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
    End If

    Unload frmIndex
End Sub

Private Sub cmdCancel_Click()
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InElementEditor = False
    InSpellEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InQuestEditor = False
    Unload frmIndex
End Sub
