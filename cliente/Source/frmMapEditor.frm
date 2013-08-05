VERSION 5.00
Begin VB.Form frmMapEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AlterEngine | Editor de Mapas"
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Edición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1335
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   8175
      Begin Eclipse.jcbutton cmdProp 
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Propiedades"
         PictureNormal   =   "frmMapEditor.frx":0000
         PictureHot      =   "frmMapEditor.frx":0954
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdFill 
         Height          =   495
         Left            =   1800
         TabIndex        =   50
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Pintar Entero"
         PictureNormal   =   "frmMapEditor.frx":12A8
         PictureHot      =   "frmMapEditor.frx":1EC4
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdGrid 
         Height          =   495
         Left            =   3600
         TabIndex        =   51
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Caption         =   "Rejilla"
         PictureNormal   =   "frmMapEditor.frx":2AE0
         PictureHot      =   "frmMapEditor.frx":3858
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmddaynight 
         Height          =   495
         Left            =   5040
         TabIndex        =   52
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Dia/Noche"
         PictureNormal   =   "frmMapEditor.frx":45D0
         PictureHot      =   "frmMapEditor.frx":5254
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdSave 
         Height          =   375
         Left            =   5040
         TabIndex        =   53
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Guardar"
         PictureNormal   =   "frmMapEditor.frx":5ED8
         PictureHot      =   "frmMapEditor.frx":66BC
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdScreeny 
         Height          =   495
         Left            =   6600
         TabIndex        =   54
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Modo Fotografia"
         PictureNormal   =   "frmMapEditor.frx":6EA0
         PictureHot      =   "frmMapEditor.frx":7B30
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdExit 
         Height          =   375
         Left            =   6600
         TabIndex        =   55
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Salir"
         PictureNormal   =   "frmMapEditor.frx":87C0
         PictureHot      =   "frmMapEditor.frx":9114
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
   Begin VB.CommandButton cmdED 
      Caption         =   "Eye Dropper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   9240
      Width           =   1215
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   5505
      LargeChange     =   10
      Left            =   8040
      Max             =   512
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5520
      Left            =   3600
      ScaleHeight     =   368
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   0
      Top             =   1440
      Width           =   4440
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6480
         Left            =   0
         ScaleHeight     =   432
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   448
         TabIndex        =   1
         Top             =   0
         Width           =   6720
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4125
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   3435
      Begin VB.OptionButton optMinusStat 
         Caption         =   "Estado minimo "
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
         Left            =   1560
         TabIndex        =   45
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optClick 
         Caption         =   "Click"
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
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optKill 
         Caption         =   "Matar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   43
         Top             =   3120
         Width           =   810
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Curar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   42
         Top             =   2640
         Width           =   915
      End
      Begin VB.OptionButton optRoofBlock 
         Caption         =   "Tejado/Bloquear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optRoof 
         Caption         =   "Tejado"
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
         TabIndex        =   40
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optWalkThru 
         Caption         =   "Pasar Atraves"
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
         TabIndex        =   39
         Top             =   3360
         Width           =   1335
      End
      Begin VB.OptionButton OptGHook 
         Caption         =   "Piedra Grapple"
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
         Left            =   1560
         TabIndex        =   38
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optGuildBlock 
         Caption         =   "Bloquear Clan"
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
         TabIndex        =   37
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optHouse 
         Caption         =   "Casa de Jugador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   36
         Top             =   2880
         Width           =   1410
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   35
         Top             =   2400
         Width           =   1170
      End
      Begin VB.OptionButton optScripted 
         Caption         =   "Scripted"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1050
      End
      Begin VB.OptionButton optClassChange 
         Caption         =   "Cambiar Clase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   33
         Top             =   1200
         Width           =   1200
      End
      Begin VB.OptionButton optChest 
         Caption         =   "Cofre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   32
         Top             =   3360
         Width           =   720
      End
      Begin VB.OptionButton optNotice 
         Caption         =   "Noticia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   31
         Top             =   1440
         Width           =   1155
      End
      Begin VB.OptionButton optDoor 
         Caption         =   "Puerta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   30
         Top             =   1680
         Width           =   960
      End
      Begin VB.OptionButton optSign 
         Caption         =   "Aviso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1080
      End
      Begin VB.OptionButton optSprite 
         Caption         =   "Cambiar Sprite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   28
         Top             =   1920
         Width           =   1440
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Reproducir Sonido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1410
      End
      Begin VB.OptionButton optArena 
         Caption         =   "Arena"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   1170
      End
      Begin VB.OptionButton optCBlock 
         Caption         =   "Bloquear Clase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   25
         Top             =   2160
         Width           =   1410
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Tienda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   810
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Abrir con Llave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Bloquear"
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
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Mover a"
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
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Vaciar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   3720
         Width           =   1335
      End
      Begin VB.OptionButton optItem 
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
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Bloquear Npc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Llave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Capas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1680
      Begin VB.OptionButton optF2Anim 
         Caption         =   "Animación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1080
      End
      Begin VB.OptionButton optFringe2 
         Caption         =   "Superior 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1320
      End
      Begin VB.OptionButton optFAnim 
         Caption         =   "Animación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optM2Anim 
         Caption         =   "Animación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1245
      End
      Begin VB.OptionButton optMask2 
         Caption         =   "Mascara 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1365
      End
      Begin VB.OptionButton optGround 
         Caption         =   "Suelo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optMask 
         Caption         =   "Mascara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAnim 
         Caption         =   "Animación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optFringe 
         Caption         =   "Superior"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Vaciar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   5
         Top             =   3000
         Width           =   975
      End
   End
   Begin VB.Frame frmtile 
      Caption         =   "Tileset"
      Height          =   1095
      Left            =   2040
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
      Begin VB.HScrollBar TilesetValue 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   47
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label TilesetLabel 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
   End
   Begin Eclipse.jcbutton cmdtype 
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   56
      Top             =   6240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
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
      Caption         =   "Iluminación Dinamica"
      PictureNormal   =   "frmMapEditor.frx":9A68
      PictureHot      =   "frmMapEditor.frx":A7DC
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Eclipse.jcbutton cmdtype 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   57
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Capas"
      PictureNormal   =   "frmMapEditor.frx":B550
      PictureHot      =   "frmMapEditor.frx":BFB4
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Eclipse.jcbutton cmdtype 
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   58
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Atributos"
      PictureNormal   =   "frmMapEditor.frx":CA18
      PictureHot      =   "frmMapEditor.frx":D9BC
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KeyShift As Boolean

Private Sub cmdED_Click()
    If Me.MousePointer = 2 Or frmMirage.MousePointer = 2 Then
        Me.MousePointer = 1
        frmMirage.MousePointer = 1
    Else
        Me.MousePointer = 2
        frmMirage.MousePointer = 2
    End If
End Sub

Private Sub cmdExit_Click()
    Dim X As Long

    X = MsgBox("¿Estas seguro de descartar los cambios?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    Call EditorCancel
End Sub

Private Sub cmdFill_Click()
    Dim Y As Long
    Dim X As Long

    X = MsgBox("¿Estas seguro de querer pintar todo el mapa?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    If MapEditorSelectedType = 1 Then
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                    If Me.optGround.value Then
                        .Ground = EditorTileY * TilesInSheets + EditorTileX
                        .GroundSet = EditorSet
                    End If
                    If Me.optMask.value Then
                        .mask = EditorTileY * TilesInSheets + EditorTileX
                        .MaskSet = EditorSet
                    End If
                    If Me.optAnim.value Then
                        .Anim = EditorTileY * TilesInSheets + EditorTileX
                        .AnimSet = EditorSet
                    End If
                    If Me.optMask2.value Then
                        .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                        .Mask2Set = EditorSet
                    End If
                    If Me.optM2Anim.value Then
                        .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                        .M2AnimSet = EditorSet
                    End If
                    If Me.optFringe.value Then
                        .Fringe = EditorTileY * TilesInSheets + EditorTileX
                        .FringeSet = EditorSet
                    End If
                    If Me.optFAnim.value Then
                        .FAnim = EditorTileY * TilesInSheets + EditorTileX
                        .FAnimSet = EditorSet
                    End If
                    If Me.optFringe2.value Then
                        .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                        .Fringe2Set = EditorSet
                    End If
                    If Me.optF2Anim.value Then
                        .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                        .F2AnimSet = EditorSet
                    End If
                End With
            Next X
        Next Y
    ElseIf MapEditorSelectedType = 2 Then
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                    If Me.optBlocked.value Then
                        .Type = TILE_TYPE_BLOCKED
                    End If
                    If Me.optWarp.value Then
                        .Type = TILE_TYPE_WARP
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optHeal.value Then
                        .Type = TILE_TYPE_HEAL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optKill.value Then
                        .Type = TILE_TYPE_KILL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optItem.value Then
                        .Type = TILE_TYPE_ITEM
                        .Data1 = ItemEditorNum
                        .Data2 = ItemEditorValue
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optNpcAvoid.value Then
                        .Type = TILE_TYPE_NPCAVOID
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optKey.value Then
                        .Type = TILE_TYPE_KEY
                        .Data1 = KeyEditorNum
                        .Data2 = KeyEditorTake
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optKeyOpen.value Then
                        .Type = TILE_TYPE_KEYOPEN
                        .Data1 = KeyOpenEditorX
                        .Data2 = KeyOpenEditorY
                        .Data3 = 0
                        .String1 = KeyOpenEditorMsg
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optShop.value Then
                        .Type = TILE_TYPE_SHOP
                        .Data1 = EditorShopNum
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optCBlock.value Then
                        .Type = TILE_TYPE_CBLOCK
                        .Data1 = EditorItemNum1
                        .Data2 = EditorItemNum2
                        .Data3 = EditorItemNum3
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optArena.value Then
                        .Type = TILE_TYPE_ARENA
                        .Data1 = Arena1
                        .Data2 = Arena2
                        .Data3 = Arena3
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optSound.value Then
                        .Type = TILE_TYPE_SOUND
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = SoundFileName
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optSprite.value Then
                        .Type = TILE_TYPE_SPRITE_CHANGE
                        .Data1 = SpritePic
                        .Data2 = SpriteItem
                        .Data3 = SpritePrice
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optSign.value Then
                        .Type = TILE_TYPE_SIGN
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = SignLine1
                        .String2 = SignLine2
                        .String3 = SignLine3
                    End If
                    If Me.optDoor.value Then
                        .Type = TILE_TYPE_DOOR
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optNotice.value Then
                        .Type = TILE_TYPE_NOTICE
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = NoticeTitle
                        .String2 = NoticeText
                        .String3 = NoticeSound
                    End If
                    If Me.optChest.value Then
                        .Type = TILE_TYPE_CHEST
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optClassChange.value Then
                        .Type = TILE_TYPE_CLASS_CHANGE
                        .Data1 = ClassChange
                        .Data2 = ClassChangeReq
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optScripted.value Then
                        .Type = TILE_TYPE_SCRIPTED
                        .Data1 = ScriptNum
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optGuildBlock.value Then
                        .Type = TILE_TYPE_GUILDBLOCK
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = GuildBlock
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optBank.value Then
                        .Type = TILE_TYPE_BANK
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.OptGHook.value Then
                        .Type = TILE_TYPE_HOOKSHOT
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                End With
            Next X
        Next Y
    ElseIf MapEditorSelectedType = 3 Then
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light = EditorTileY * TilesInSheets + EditorTileX
            Next X
        Next Y
    End If
End Sub

Private Sub cmdGrid_Click()
    If GridMode = 0 Then
        GridMode = 1
    Else
        GridMode = 0
    End If
End Sub

Private Sub cmdScreeny_Click()
    If ScreenMode = 0 Then
        ScreenMode = 1
    Else
        ScreenMode = 0
    End If
End Sub

Private Sub cmddaynight_Click()
    If NightMode = 0 Then
        NightMode = 1
    Else
        NightMode = 0
    End If
End Sub

Private Sub cmdProp_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub cmdSave_Click()
    Dim X As Long

    X = MsgBox("¿Estas seguro de guardar estos cambios?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    Call EditorSend
End Sub

Private Sub cmdtype_Click(index As Integer)
    If index = 1 Then
        MapEditorSelectedType = 1

        Me.fraAttribs.Visible = False
        Me.fraLayers.Visible = True
        Me.frmtile.Visible = True
    ElseIf index = 2 Then
        MapEditorSelectedType = 2

        Me.shpSelected.Width = 32
        Me.shpSelected.Height = 32

        Me.fraLayers.Visible = False
        Me.frmtile.Visible = False
        Me.fraAttribs.Visible = True
    Else
        MapEditorSelectedType = 3

        Me.fraAttribs.Visible = False
        Me.fraLayers.Visible = False
        Me.frmtile.Visible = False
        'Me.Option1(10).Value = True

        Me.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles10.bmp")

        EditorSet = 10

        scrlPicture.max = Int((picBackSelect.Height - picBack.Height) / PIC_Y)
    End If
End Sub

Private Sub deshacer_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = False
    End If
End Sub

Private Sub optClick_Click()
    frmClick.Show vbModal
End Sub

Private Sub optChest_Click()
frmChest.Show vbModal
End Sub

Private Sub optGuildBlock_Click()
    frmGuildBlock.Show vbModal
    frmGuildBlock.txtGuild.Text = vbNullString
End Sub



Private Sub optMinusStat_Click()
    frmMinusStat.Show
    frmMinusStat.scrlNum1.value = MinusHp
    frmMinusStat.lblNum1.Caption = MinusHp
    frmMinusStat.scrlNum2.value = MinusMp
    frmMinusStat.lblNum2.Caption = MinusMp
    frmMinusStat.scrlNum3.value = MinusSp
    frmMinusStat.lblNum3.Caption = MinusSp
    frmMinusStat.Text1.Text = MessageMinus
End Sub

Private Sub optRoof_Click()
    frmRoofTile.Show vbModal
End Sub

Private Sub picBackSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub picBackSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = False
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Not KeyShift Then
            Call EditorChooseTile(Button, Shift, X, Y)

            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(X / PIC_X)
            EditorTileY = Int(Y / PIC_Y)

            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If

            If Int(EditorTileY * PIC_Y) >= shpSelected.top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If

    If MapEditorSelectedType = 2 Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If

    EditorTileX = Int((shpSelected.Left + PIC_X) / PIC_X)
    EditorTileY = Int((shpSelected.top + PIC_Y) / PIC_Y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Not KeyShift Then
            Call EditorChooseTile(Button, Shift, X, Y)

            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(X / PIC_X)
            EditorTileY = Int(Y / PIC_Y)

            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If

            If Int(EditorTileY * PIC_Y) >= shpSelected.top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If

    If MapEditorSelectedType = 2 Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If

    EditorTileX = Int(shpSelected.Left / PIC_X)
    EditorTileY = Int(shpSelected.top / PIC_Y)
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
End Sub

Private Sub optArena_Click()
    frmArena.Show vbModal
End Sub

Private Sub optCBlock_Click()
    frmBClass.scrlNum1.max = Max_Classes
    frmBClass.scrlNum2.max = Max_Classes
    frmBClass.scrlNum3.max = Max_Classes
    frmBClass.Show vbModal
End Sub

Private Sub optClassChange_Click()
    frmClassChange.scrlClass.max = Max_Classes
    frmClassChange.scrlReqClass.max = Max_Classes
    frmClassChange.Show vbModal
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModal
End Sub

Private Sub optItem_Click()
    frmMapItem.scrlItem.value = 1
    frmMapItem.Show vbModal
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModal
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModal
End Sub

Private Sub optNotice_Click()
    frmNotice.Show vbModal
End Sub

Private Sub optScripted_Click()
    frmScript.Show vbModal
End Sub

Private Sub optShop_Click()
    frmShop.scrlNum.max = MAX_SHOPS
    frmShop.Show vbModal
End Sub

Private Sub optSign_Click()
    frmSign.Show vbModal
End Sub

Private Sub optSound_Click()
    frmSound.Show vbModal
End Sub

Private Sub optSprite_Click()
    If SpriteSize = 1 Then
        frmSpriteChange.picSprite.Height = 960
    End If

    frmSpriteChange.scrlItem.max = MAX_ITEMS
    frmSpriteChange.Show vbModal
End Sub

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub

Private Sub optHouse_Click()
    frmHouse.scrlItem.max = MAX_ITEMS
    frmHouse.Show vbModal
End Sub

Private Sub TilesetValue_Change()

TilesetLabel.Caption = TilesetValue.value

If FileExists("GFX\Tiles" & TilesetValue.value & ".bmp") = True Then

    frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & TilesetValue.value & ".bmp")

    EditorSet = TilesetValue.value
    
End If
End Sub
