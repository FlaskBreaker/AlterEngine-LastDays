VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " AlterEngine | Panel de Administración"
   ClientHeight    =   5865
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   6120
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "frmAdmin"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Eclipse.jcbutton btnClose 
      Height          =   495
      Left            =   3840
      TabIndex        =   26
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      BackColor       =   16641248
      Caption         =   "Cerrar Panel"
      PictureNormal   =   "frmAdmin.frx":0FC2
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   529
      TabMaxWidth     =   2999
      BackColor       =   -2147483648
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jugadores"
      TabPicture(0)   =   "frmAdmin.frx":1916
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPlayerName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblValue"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnWarpToMe"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnWarpMeTo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnKick"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnBan"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "btnSetSprite"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnSetAccess"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdGive"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdTake"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtValue"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPlayerName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtItem"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Mundo"
      TabPicture(1)   =   "frmAdmin.frx":1932
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMapNumber"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "btnWarpTo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "btnLocation"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "btnRespawn"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "btnSetNone"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "btnSetRain"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "btnSetThunder"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "btnSetSnow"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtMap"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Editores"
      TabPicture(2)   =   "frmAdmin.frx":194E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btnEditItem"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "btnEditSpell"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnEditShops"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "btnEditEmoticon"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnEditElement"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "btnEditArrow"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "btnEditQuests"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "btnEditNPC"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "btnEditMap"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtMap 
         Alignment       =   2  'Center
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
         Left            =   -74400
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtPlayerName 
         Alignment       =   2  'Center
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
         Left            =   480
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin Eclipse.jcbutton btnEditMap 
         Height          =   495
         Left            =   -71760
         TabIndex        =   10
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de Mapa"
         PictureNormal   =   "frmAdmin.frx":196A
         PictureHot      =   "frmAdmin.frx":2B2E
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditNPC 
         Height          =   495
         Left            =   -71760
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de NPCS    "
         PictureNormal   =   "frmAdmin.frx":3CF2
         PictureHot      =   "frmAdmin.frx":49EE
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditQuests 
         Height          =   495
         Left            =   -71760
         TabIndex        =   12
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de Quests"
         PictureNormal   =   "frmAdmin.frx":56EA
         PictureHot      =   "frmAdmin.frx":637E
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditArrow 
         Height          =   495
         Left            =   -71760
         TabIndex        =   13
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de Flechas"
         PictureNormal   =   "frmAdmin.frx":7012
         PictureHot      =   "frmAdmin.frx":8066
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditElement 
         Height          =   495
         Left            =   -74520
         TabIndex        =   14
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editar Elementos"
         PictureNormal   =   "frmAdmin.frx":90BA
         PictureHot      =   "frmAdmin.frx":A10E
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditEmoticon 
         Height          =   495
         Left            =   -74520
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editar Emoticonos"
         PictureNormal   =   "frmAdmin.frx":B162
         PictureHot      =   "frmAdmin.frx":BFC2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditShops 
         Height          =   495
         Left            =   -74520
         TabIndex        =   16
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de Tiendas"
         PictureNormal   =   "frmAdmin.frx":CE22
         PictureHot      =   "frmAdmin.frx":E1AE
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditSpell 
         Height          =   495
         Left            =   -74520
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de Hechizos"
         PictureNormal   =   "frmAdmin.frx":F53A
         PictureHot      =   "frmAdmin.frx":1058E
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnEditItem 
         Height          =   495
         Left            =   -74520
         TabIndex        =   18
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Editor de Objetos"
         PictureNormal   =   "frmAdmin.frx":115E2
         PictureHot      =   "frmAdmin.frx":12516
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnSetSnow 
         Height          =   375
         Left            =   -71520
         TabIndex        =   19
         Top             =   2400
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
         Caption         =   "Nieve"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnSetThunder 
         Height          =   375
         Left            =   -71520
         TabIndex        =   20
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
         Caption         =   "Truenos y lluvia"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnSetRain 
         Height          =   375
         Left            =   -71520
         TabIndex        =   21
         Top             =   1440
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
         Caption         =   "Lluvia"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnSetNone 
         Height          =   375
         Left            =   -71520
         TabIndex        =   22
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
         Caption         =   "Nada"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnRespawn 
         Height          =   375
         Left            =   -74280
         TabIndex        =   23
         Top             =   2400
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
         Caption         =   "Respawnear"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnLocation 
         Height          =   375
         Left            =   -74280
         TabIndex        =   24
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
         Caption         =   "Localización"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnWarpTo 
         Height          =   375
         Left            =   -74280
         TabIndex        =   25
         Top             =   1440
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
         Caption         =   "Ir a el mapa"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdTake 
         Height          =   375
         Left            =   4440
         TabIndex        =   27
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "Quitar Objeto"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdGive 
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   2880
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "Dar Objeto"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnSetAccess 
         Height          =   375
         Left            =   3720
         TabIndex        =   29
         Top             =   1800
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
         Caption         =   "Dar Privilegios"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnSetSprite 
         Height          =   375
         Left            =   3720
         TabIndex        =   30
         Top             =   1320
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
         Caption         =   "Cambiar Sprite"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnBan 
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   2880
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
         Caption         =   "Banear Jugador"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnKick 
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   2400
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
         Caption         =   "Expulsar Jugador"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnWarpMeTo 
         Height          =   375
         Left            =   720
         TabIndex        =   33
         Top             =   1440
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
         Caption         =   "Ir hacia el jugador"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton btnWarpToMe 
         Height          =   375
         Left            =   720
         TabIndex        =   34
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
         Caption         =   "Traer el jugador"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Objeto:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Climatologia:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71520
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblValue 
         Caption         =   "Introduce un Valor:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblMapNumber 
         Caption         =   "Numero de Mapa:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74400
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblPlayerName 
         Caption         =   "Nombre del Jugador:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   120
      Picture         =   "frmAdmin.frx":1344A
      Top             =   120
      Width           =   5880
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEditMap_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendRequestEditMap
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditQuests_Click()
Call InitQuestEditor
End Sub

Private Sub btnSetSprite_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            If LenB(txtValue.Text) >= 0 Then
                If IsNumeric(txtValue.Text) Then
                    If Not Val(txtValue.Text) < 0 Then
                        Call SendSetPlayerSprite(txtPlayerName.Text, txtValue.Text)
                    End If
                End If
            End If
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnBan_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call SendBan(txtPlayerName.Text)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditItem_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditItem
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditShops_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditShop
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditSpell_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditSpell
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnKick_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MONITER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call SendKick(txtPlayerName.Text)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnLocation_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        BLoc = Not BLoc
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnRespawn_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendMapRespawn
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnWarpMeTo_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call WarpMeTo(txtPlayerName.Text)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnWarpTo_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Len(txtMap.Text) <> 0 Then
            If GetPlayerMap(MyIndex) <> Val(txtMap.Text) Then
                Call WarpTo(Val(txtMap.Text), GetPlayerX(MyIndex), GetPlayerY(MyIndex))
            Else
                Call AddText("Ya estas en ese mapa. No puedes ir otra vez.", BRIGHTRED)
            End If
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnWarpToMe_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName.Text) <> 0 Then
            Call WarpToMe(txtPlayerName.Text)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub PlayerInfo_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If LenB(txtPlayerName) <> 0 Then
            Call SendData("getstats" & SEP_CHAR & txtPlayerName.Text & END_CHAR)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditArrow_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditArrow
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditEmoticon_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditEmoticon
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditNPC_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditNPC
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnEditElement_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditElement
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnSetAccess_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
        If LenB(txtPlayerName.Text) <> 0 Then
            If LenB(txtValue.Text) <> 0 Then
                If Val(txtValue.Text) < 0 Or Val(txtValue.Text) > 5 Then
                    Call AddText("El rango valido debe estar entre 0 y 5.", BRIGHTRED)
                Else
                    Call SendSetAccess(txtPlayerName.Text, txtValue.Text)
                End If
            End If
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub btnSetNone_Click()
    Call SendData("weather" & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub btnSetRain_Click()
    Call SendData("weather" & SEP_CHAR & 1 & END_CHAR)
End Sub

Private Sub btnSetSnow_Click()
    Call SendData("weather" & SEP_CHAR & 2 & END_CHAR)
End Sub

Private Sub btnSetThunder_Click()
    Call SendData("weather" & SEP_CHAR & 3 & END_CHAR)
End Sub

Private Sub btnClose_Click()
    frmAdmin.Visible = False
End Sub

' Función para enviar objeto a jugadores
' Añadido en AE v1.1 por Stream

Private Sub cmdGive_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
        If Trim(txtPlayerName.Text) <> "" And Trim(txtItem.Text) <> "" Then
            Call SendGiveItem(Trim(txtPlayerName.Text), Trim(txtItem.Text))
            Call AddText("El objeto fue enviado al jugador.", Green)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

Private Sub cmdTake_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
        If Trim(txtPlayerName.Text) <> "" And Trim(txtItem.Text) <> "" Then
            Call SendTakeItem(Trim(txtPlayerName.Text), Trim(txtItem.Text))
            Call AddText("Has cogido un objeto de un jugador!", Green)
        End If
    Else
        Call AddText("No estás autorizado para realizar esta acción.", BRIGHTRED)
    End If
End Sub

