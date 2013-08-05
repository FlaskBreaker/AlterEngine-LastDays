VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de AlterEngine"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   670
      TabCaption(0)   =   "Configuración Principal"
      TabPicture(0)   =   "frmConfiguracion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "nombrejuego"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "webjuego"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "puertojuego"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "jcbutton2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Parametros"
      TabPicture(1)   =   "frmConfiguracion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label18"
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(2)=   "Label20"
      Tab(1).Control(3)=   "Label21"
      Tab(1).Control(4)=   "Label22"
      Tab(1).Control(5)=   "Label23"
      Tab(1).Control(6)=   "Label24"
      Tab(1).Control(7)=   "Label25"
      Tab(1).Control(8)=   "Label26"
      Tab(1).Control(9)=   "Label44"
      Tab(1).Control(10)=   "Label45"
      Tab(1).Control(11)=   "Label46"
      Tab(1).Control(12)=   "jcbutton4"
      Tab(1).Control(13)=   "jcbutton3"
      Tab(1).Control(14)=   "scripting"
      Tab(1).Control(15)=   "scripterrores"
      Tab(1).Control(16)=   "paperdoll"
      Tab(1).Control(17)=   "guardartiempo"
      Tab(1).Control(18)=   "sizedenpc"
      Tab(1).Control(19)=   "customnpc"
      Tab(1).Control(20)=   "nivelminimopk"
      Tab(1).Control(21)=   "mostrarnivel"
      Tab(1).Control(22)=   "clases"
      Tab(1).Control(23)=   "Frame4"
      Tab(1).Control(24)=   "Text1"
      Tab(1).Control(25)=   "Text2"
      Tab(1).Control(26)=   "Text3"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "Maximos"
      TabPicture(2)   =   "frmConfiguracion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label29"
      Tab(2).Control(1)=   "Label30"
      Tab(2).Control(2)=   "Label31"
      Tab(2).Control(3)=   "Label32"
      Tab(2).Control(4)=   "Label33"
      Tab(2).Control(5)=   "Label34"
      Tab(2).Control(6)=   "Label35"
      Tab(2).Control(7)=   "Label36"
      Tab(2).Control(8)=   "Label37"
      Tab(2).Control(9)=   "Label38"
      Tab(2).Control(10)=   "Label39"
      Tab(2).Control(11)=   "Label40"
      Tab(2).Control(12)=   "Label41"
      Tab(2).Control(13)=   "Label42"
      Tab(2).Control(14)=   "maxscriptedhechizos"
      Tab(2).Control(15)=   "maxgrupos"
      Tab(2).Control(16)=   "maxnivel"
      Tab(2).Control(17)=   "maxelementos"
      Tab(2).Control(18)=   "maxemoticonos"
      Tab(2).Control(19)=   "maxmiembrosclan"
      Tab(2).Control(20)=   "maxclanes"
      Tab(2).Control(21)=   "maxobjetosenmapa"
      Tab(2).Control(22)=   "maxmapas"
      Tab(2).Control(23)=   "maxhechizos"
      Tab(2).Control(24)=   "maxtiendas"
      Tab(2).Control(25)=   "maxnpcs"
      Tab(2).Control(26)=   "maxobjetos"
      Tab(2).Control(27)=   "maxjugadores"
      Tab(2).Control(28)=   "Frame5"
      Tab(2).ControlCount=   29
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   97
         Top             =   5400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   96
         Top             =   6120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   95
         Top             =   5760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame5 
         Caption         =   "Consejo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -73800
         TabIndex        =   91
         Top             =   960
         Width           =   5895
         Begin VB.Label Label43 
            Caption         =   $"frmConfiguracion.frx":0054
            Height          =   495
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   5655
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Información"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1335
         Left            =   -73920
         TabIndex        =   72
         Top             =   600
         Width           =   5775
         Begin VB.Label Label28 
            Caption         =   "Para activar siempre será el valor 1, para desactivar el valor 0."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   5535
         End
         Begin VB.Label Label27 
            Caption         =   "En la sección parametros solo sirve para activar/desactivar sistemas y configuraciones en AE."
            Height          =   495
            Left            =   360
            TabIndex        =   73
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Nombre de los estados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1215
         Left            =   720
         TabIndex        =   54
         Top             =   2880
         Width           =   6375
         Begin VB.TextBox stat1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   58
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox stat2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   57
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox stat3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   56
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox stat4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   55
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Estado 4:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   3360
            TabIndex        =   62
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Estado 3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   3360
            TabIndex        =   61
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Estado 2:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   360
            TabIndex        =   60
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Estado 1:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   360
            TabIndex        =   59
            Top             =   360
            Width           =   975
         End
      End
      Begin Server.jcbutton jcbutton2 
         Height          =   375
         Left            =   6720
         TabIndex        =   53
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "?"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "El tamaño recomendado es 30x30. El tamaño maximo optimo recomendable es 50x50."
         TooltipType     =   1
         TooltipIcon     =   1
         TooltipTitle    =   "Ayuda"
         TooltipBackColor=   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Configuración de Mapas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   840
         TabIndex        =   46
         Top             =   1320
         Width           =   6135
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   102
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox scrollY 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   52
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox scrollX 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   51
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox scroll 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   48
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label47 
            Caption         =   "Codigo de Encriptacion:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   495
            Left            =   3120
            TabIndex        =   101
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Tamaño mapa Y:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3360
            TabIndex        =   50
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Tamaño mapa X:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   480
            TabIndex        =   49
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Mapas con scroll:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   480
            TabIndex        =   47
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Regeneraciones de Estados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   2535
         Left            =   240
         TabIndex        =   27
         Top             =   4200
         Width           =   7575
         Begin VB.TextBox timermagia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            TabIndex        =   35
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox regeneracionmagia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   34
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox vidatimer 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            TabIndex        =   33
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox regereacionvida 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   32
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox stamregeneracion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   31
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox timerstam 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            TabIndex        =   30
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox npcregenen 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   29
            Top             =   2160
            Width           =   375
         End
         Begin Server.jcbutton jcbutton1 
            Height          =   375
            Left            =   7200
            TabIndex        =   28
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16765357
            Caption         =   "?"
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            CaptionEffects  =   0
            ToolTip         =   "Las regeneraciones son si cada X segundos, el jugador recupera poco a poco esos stats."
            TooltipType     =   1
            TooltipIcon     =   1
            TooltipTitle    =   "Ayuda"
            TooltipBackColor=   0
         End
         Begin VB.Label Label4 
            Caption         =   "Regeneración de vida:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   720
            TabIndex        =   42
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "Tiempo de regeneración:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3720
            TabIndex        =   41
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "Regeneración de magia:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   480
            TabIndex        =   40
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "Tiempo de regeneración:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3720
            TabIndex        =   39
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "Regeneración de estamina:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "Tiempo de regeneración:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3720
            TabIndex        =   37
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label10 
            Caption         =   "Regeneración de vida en NPCS:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   2160
            Width           =   3375
         End
      End
      Begin VB.TextBox puertojuego 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox webjuego 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox nombrejuego 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox clases 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   23
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox mostrarnivel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   22
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox nivelminimopk 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   21
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox customnpc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         TabIndex        =   20
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox sizedenpc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   19
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox guardartiempo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         TabIndex        =   18
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox paperdoll 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   17
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox scripterrores 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         TabIndex        =   16
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox scripting 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         TabIndex        =   15
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox maxjugadores 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   14
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox maxobjetos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   13
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox maxnpcs 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   12
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox maxtiendas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   11
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox maxhechizos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   10
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox maxmapas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   9
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox maxobjetosenmapa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   8
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox maxclanes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   7
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox maxmiembrosclan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72000
         TabIndex        =   6
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox maxemoticonos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   5
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox maxelementos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   4
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox maxnivel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   3
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox maxgrupos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   2
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox maxscriptedhechizos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68040
         TabIndex        =   1
         Top             =   4560
         Width           =   735
      End
      Begin Server.jcbutton jcbutton3 
         Height          =   255
         Left            =   -71760
         TabIndex        =   75
         Top             =   3480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "?"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Pon 0 para sprites normales 32x64, o 1 para que lea sprites por partes de los archivos head.bmp, legs.bmp, bodys.bmp."
         TooltipType     =   1
         TooltipIcon     =   1
         TooltipTitle    =   "Ayuda"
         TooltipBackColor=   0
      End
      Begin Server.jcbutton jcbutton4 
         Height          =   255
         Left            =   -68160
         TabIndex        =   76
         Top             =   3480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "?"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Pon 1 para Sprites de 32x64 o 0 para sprites de 32x32."
         TooltipType     =   1
         TooltipIcon     =   1
         TooltipTitle    =   "Ayuda"
         TooltipBackColor=   0
      End
      Begin VB.Label Label46 
         Caption         =   "Max Legs:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69960
         TabIndex        =   100
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label45 
         Caption         =   "Max Bodys:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69960
         TabIndex        =   99
         Top             =   6120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "Max Heads:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69960
         TabIndex        =   98
         Top             =   5760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label42 
         Caption         =   "Hechizos Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69960
         TabIndex        =   90
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label41 
         Caption         =   "Maximo de objetos en mapa:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70920
         TabIndex        =   89
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label Label40 
         Caption         =   "Hechizos por script Max.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70680
         TabIndex        =   88
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label Label39 
         Caption         =   "Emoticonos Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70320
         TabIndex        =   87
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label38 
         Caption         =   "Grupos Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69840
         TabIndex        =   86
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label37 
         Caption         =   "Nivel Maximo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69600
         TabIndex        =   85
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label36 
         Caption         =   "Elementos Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70200
         TabIndex        =   84
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label35 
         Caption         =   "Jugadores en clan Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -74880
         TabIndex        =   83
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label34 
         Caption         =   "Clanes Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   82
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label33 
         Caption         =   "Tiendas Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   81
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label32 
         Caption         =   "Mapas Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73680
         TabIndex        =   80
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label31 
         Caption         =   "NPCS Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73560
         TabIndex        =   79
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Objetos Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   78
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Jugadores Maximos:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -74040
         TabIndex        =   77
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label26 
         Caption         =   "Sistema de clases:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70680
         TabIndex        =   71
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label25 
         Caption         =   "Mostrar nivel en el nombre:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -71520
         TabIndex        =   70
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label Label24 
         Caption         =   "Nivel minimo para matar un jugador:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -72360
         TabIndex        =   69
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label Label23 
         Caption         =   "Tipo de Sprite:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73800
         TabIndex        =   68
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label22 
         Caption         =   "Tamaño de Sprites:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70680
         TabIndex        =   67
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label21 
         Caption         =   "Guardar Tiempo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -73920
         TabIndex        =   66
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Paperdoll:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -69720
         TabIndex        =   65
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Errores de Script:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -70440
         TabIndex        =   64
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Sistema de Scripts:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   -74280
         TabIndex        =   63
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del juego:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Web del juego:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   3480
         TabIndex        =   44
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Puerto del juego:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   5520
         TabIndex        =   43
         Top             =   600
         Width           =   1935
      End
   End
   Begin Server.jcbutton guardarconfig 
      Height          =   495
      Left            =   1560
      TabIndex        =   93
      Top             =   7080
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Guardar Configuración"
      PictureNormal   =   "frmConfiguracion.frx":00E8
      PictureHot      =   "frmConfiguracion.frx":0A3C
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton cmdCancel 
      Height          =   495
      Left            =   4560
      TabIndex        =   94
      Top             =   7080
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Cancelar"
      PictureNormal   =   "frmConfiguracion.frx":1390
      PictureHot      =   "frmConfiguracion.frx":1CE4
      CaptionEffects  =   4
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub customnpc_Change()
If Not IsNumeric(customnpc.Text) Then
customnpc.Text = 0
End If
If customnpc.Text = 1 Then
Label44.Visible = True
Label45.Visible = True
Label46.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text1.Text = MAX_HEAD
Text2.Text = MAX_BODY
Text3.Text = MAX_LEGS
nivelminimopk.Top = 5520
Label24.Top = 5520
Label25.Top = 6000
mostrarnivel.Top = 6000
Label26.Top = 6480
clases.Top = 6480
Else
Label44.Visible = False
Label45.Visible = False
Label46.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Label24.Top = 3960
Label25.Top = 4440
Label26.Top = 4920
nivelminimopk.Top = 3960
mostrarnivel.Top = 4440
clases.Top = 4920
End If

End Sub

Private Sub Form_Load()

    If FileExists("\Configuracion.ini") Then
        
    ' Configuración Principal
    nombrejuego.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "GameName")
    webjuego.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "WebSite")
    puertojuego.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Port")
    scroll.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Scrolling")
    scrollX.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ScrollX")
    scrollY.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ScrollY")
    stat1.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat1")
    stat2.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat2")
    stat3.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat3")
    stat4.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat4")
    ' Estados
    regereacionvida.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "HPRegen")
    vidatimer.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "HPTimer")
    regeneracionmagia.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MPRegen")
    timermagia.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MPTimer")
    stamregeneracion.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SPRegen")
    timerstam.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SPTimer")
    npcregenen.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "NPCRegen")
    ' Parametros
    scripting.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Scripting")
    scripterrores.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ScriptErrors")
    guardartiempo.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SaveTime")
    paperdoll.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PaperDoll")
    customnpc.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Custom")
    sizedenpc.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SpriteSize")
    nivelminimopk.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PKMinLvl")
    mostrarnivel.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Level")
    clases.Text = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Classes")
    ' Maximos
    maxjugadores.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_PLAYERS")
    maxelementos.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_ELEMENTS")
    maxobjetos.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_ITEMS")
    maxnivel.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_LEVEL")
    maxnpcs.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_NPCS")
    maxgrupos.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_PARTY_MEMBERS")
    maxmapas.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_MAPS")
    maxemoticonos.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_EMOTICONS")
    maxtiendas.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SHOPS")
    maxscriptedhechizos.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SCRIPTSPELLS")
    maxclanes.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_GUILDS")
    maxobjetosenmapa.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_MAP_ITEMS")
    maxmiembrosclan.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_GUILD_MEMBERS")
    maxhechizos.Text = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SPELLS")
    
    Else
        MsgBox "No se ha encontrado el archivo", vbInformation
    End If

End Sub

Private Sub guardarconfig_Click()

Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "GameName", nombrejuego.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "WebSite", webjuego.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Port", puertojuego.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Scrolling", scroll.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "ScrollX", scrollX.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat1", stat1.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat2", stat2.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat3", stat3.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat4", stat4.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "HPRegen", regereacionvida.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "HPTimer", vidatimer.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "MPRegen", regeneracionmagia.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "MPTimer", timermagia.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "SPRegen", stamregeneracion.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "NPCRegen", npcregenen.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "SPTimer", timerstam.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Scripting", scripting.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "ScriptErrors", scripterrores.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "SaveTime", guardartiempo.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "PaperDoll", paperdoll.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Custom", customnpc.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "SpriteSize", sizedenpc.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "PKMinLvl", nivelminimopk.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Level", mostrarnivel.Text)
Call PutVar(App.Path & "\Configuracion.ini", "CONFIG", "Classes", clases.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_PLAYERS", maxjugadores.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_ELEMENTS", maxelementos.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_ITEMS", maxobjetos.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_LEVEL", maxnivel.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_NPCS", maxnpcs.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_PARTY_MEMBERS", maxgrupos.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_MAPS", maxmapas.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_EMOTICONS", maxemoticonos.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SHOPS", maxtiendas.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SCRIPTSPELLS", maxscriptedhechizos.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_GUILDS", maxclanes.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_MAP_ITEMS", maxobjetosenmapa.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_GUILD_MEMBERS", maxmiembrosclan.Text)
Call PutVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SPELLS", maxhechizos.Text)
Call PutVar(App.Path & "\Configuacion.ini", "MAX", "MAX_HEAD", Text1.Text)
Call PutVar(App.Path & "\Configuacion.ini", "MAX", "MAX_BODY", Text2.Text)
Call PutVar(App.Path & "\Configuacion.ini", "MAX", "MAX_LEGS", Text3.Text)
If Text4.Text <> "" Then
Call SaveSetting(App.EXEName, "Clave", "Clave", Text4.Text)
Call SendDataToAll("SPASS2" & SEP_CHAR & Text4.Text & END_CHAR)
End If

MsgBox "La configuración fue guardada correctamente.", vbInformation

Me.Visible = False

End Sub

