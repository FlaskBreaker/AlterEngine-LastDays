VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmQuestEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de Quest"
   ClientHeight    =   6855
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab3 
      Height          =   3255
      Left            =   240
      TabIndex        =   21
      Top             =   2880
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5741
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Dar objeto al empezar"
      TabPicture(0)   =   "frmQuestEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstartitem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstartval"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkstart"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlstartnum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlstartval"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Objeto para la quest"
      TabPicture(1)   =   "frmQuestEditor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblquestitem"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblquestval"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "scrlquestvalue"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "scrlquestitem"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Recompensa"
      TabPicture(2)   =   "frmQuestEditor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblrewval"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblrewitem"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label18"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "scrlrewitem"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "scrlrewval"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.HScrollBar scrlrewval 
         Height          =   255
         Left            =   -72960
         Min             =   1
         TabIndex        =   39
         Top             =   2280
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlrewitem 
         Height          =   255
         Left            =   -72960
         Max             =   500
         Min             =   1
         TabIndex        =   38
         Top             =   1560
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlquestitem 
         Height          =   255
         Left            =   -72840
         Max             =   500
         Min             =   1
         TabIndex        =   32
         Top             =   1560
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlquestvalue 
         Height          =   255
         Left            =   -72840
         Min             =   1
         TabIndex        =   31
         Top             =   2280
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlstartval 
         Height          =   255
         Left            =   2040
         Min             =   1
         TabIndex        =   28
         Top             =   2760
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar scrlstartnum 
         Height          =   255
         Left            =   2040
         Max             =   500
         Min             =   1
         TabIndex        =   25
         Top             =   2040
         Value           =   1
         Width           =   3975
      End
      Begin VB.CheckBox chkstart 
         Caption         =   "Dar objeto al inicio"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Valor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   43
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Objeto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   42
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblrewitem 
         Caption         =   "lblrewitem"
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
         Left            =   -72960
         TabIndex        =   41
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label lblrewval 
         Height          =   255
         Left            =   -72360
         TabIndex        =   40
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label14 
         Caption         =   "El objeto que se entregara como recompensa al jugador tras completar la quest."
         Height          =   735
         Left            =   -73800
         TabIndex        =   37
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label lblquestval 
         Height          =   255
         Left            =   -72240
         TabIndex        =   36
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblquestitem 
         Caption         =   "Label9"
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
         Left            =   -72840
         TabIndex        =   35
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label13 
         Caption         =   "Objeto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72840
         TabIndex        =   34
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label12 
         Caption         =   "Valor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72840
         TabIndex        =   33
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "El objeto que debe entregar el jugador al NPC para completar la quest."
         Height          =   255
         Left            =   -73320
         TabIndex        =   30
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label lblstartval 
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Valor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Objeto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label lblstartitem 
         Caption         =   "Label9"
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
         Left            =   2040
         TabIndex        =   24
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Label8 
         Caption         =   $"frmQuestEditor.frx":0054
         Height          =   495
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   6255
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2655
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4683
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Antes/Despues"
      TabPicture(0)   =   "frmQuestEditor.frx":0101
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "jcbutton1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtbefore"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtafter"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ayuditaquest"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Principio/Final"
      TabPicture(1)   =   "frmQuestEditor.frx":011D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "jcbutton3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "jcbutton2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtstart"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtend"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Durante"
      TabPicture(2)   =   "frmQuestEditor.frx":0139
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtnotitem"
      Tab(2).Control(1)=   "txtduring"
      Tab(2).Control(2)=   "jcbutton4"
      Tab(2).Control(3)=   "jcbutton5"
      Tab(2).Control(4)=   "Label7"
      Tab(2).Control(5)=   "Label6"
      Tab(2).ControlCount=   6
      Begin Eclipse.jcbutton ayuditaquest 
         Height          =   375
         Left            =   4200
         TabIndex        =   47
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "jcbutton"
         CaptionEffects  =   0
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   120
         Top             =   2160
      End
      Begin VB.TextBox txtnotitem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74880
         TabIndex        =   20
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox txtduring 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74880
         TabIndex        =   18
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtend 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         TabIndex        =   16
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox txtstart 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74760
         TabIndex        =   14
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtafter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtbefore 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   3975
      End
      Begin Eclipse.jcbutton jcbutton1 
         Height          =   375
         Left            =   4200
         TabIndex        =   48
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "jcbutton"
         CaptionEffects  =   0
      End
      Begin Eclipse.jcbutton jcbutton2 
         Height          =   375
         Left            =   -71160
         TabIndex        =   49
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "jcbutton"
         CaptionEffects  =   0
      End
      Begin Eclipse.jcbutton jcbutton3 
         Height          =   375
         Left            =   -71160
         TabIndex        =   50
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "jcbutton"
         CaptionEffects  =   0
      End
      Begin Eclipse.jcbutton jcbutton4 
         Height          =   375
         Left            =   -70800
         TabIndex        =   51
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "jcbutton"
         CaptionEffects  =   0
      End
      Begin Eclipse.jcbutton jcbutton5 
         Height          =   375
         Left            =   -70800
         TabIndex        =   52
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "jcbutton"
         CaptionEffects  =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Conversación para decirle al jugador que NO tiene el objeto:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   "Conversación para preguntarle al jugador si tiene el objeto:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "Conversación cuando finaliza la quest:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Conversación para ofrecer la quest al jugador:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Conversación una vez que la quest ya fue completada:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Conversación de antes de que comiences la quest:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmQuestEditor.frx":0155
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLevel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label50"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbExpReward"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstclass"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkcls"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrllvl"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chklvl"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlQuestExpReward"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.HScrollBar scrlQuestExpReward 
         Height          =   255
         Left            =   1560
         Max             =   32000
         TabIndex        =   45
         Top             =   2280
         Value           =   1
         Width           =   1935
      End
      Begin VB.CheckBox chklvl 
         Caption         =   "Nivel requerido para la quest"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.HScrollBar scrllvl 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         Max             =   500
         Min             =   1
         TabIndex        =   5
         Top             =   720
         Value           =   1
         Width           =   1095
      End
      Begin VB.CheckBox chkcls 
         Caption         =   "Clase requerida para la quest"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.ListBox lstclass 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lbExpReward 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   3600
         TabIndex        =   46
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experiencia a dar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   165
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   1350
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Caption         =   "500"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin Eclipse.jcbutton cmdOk 
      Height          =   495
      Left            =   2280
      TabIndex        =   53
      Top             =   6240
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
      BackColor       =   14935011
      Caption         =   "Aceptar y Guardar"
      PictureNormal   =   "frmQuestEditor.frx":0171
      PictureHot      =   "frmQuestEditor.frx":0955
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Eclipse.jcbutton salirsinguardar 
      Height          =   495
      Left            =   4920
      TabIndex        =   54
      Top             =   6240
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
      BackColor       =   14935011
      Caption         =   "Salir sin guardar"
      PictureNormal   =   "frmQuestEditor.frx":1139
      PictureHot      =   "frmQuestEditor.frx":1A8D
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmQuestEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check33_Click()

End Sub

Private Sub Check36_Click()

End Sub

Private Sub cmdOk_Click()
Call QuestEditorOk
End Sub

Private Sub chkcls_Click()
If frmQuestEditor.chkcls.value = 1 Then
frmQuestEditor.lstclass.Enabled = True
Else
frmQuestEditor.lstclass.Enabled = False
End If
End Sub

Private Sub chklvl_Click()
If frmQuestEditor.chklvl.value = 1 Then
frmQuestEditor.scrllvl.Enabled = True
Else
frmQuestEditor.scrllvl.Enabled = False
End If
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub chkstart_Click()
If chkstart.value = 1 Then
        frmQuestEditor.scrlstartnum.Enabled = True
        frmQuestEditor.scrlstartval.Enabled = True
        frmQuestEditor.lblstartitem = frmQuestEditor.scrlstartnum.value & ":" & Item(frmQuestEditor.scrlstartnum.value).name
        frmQuestEditor.lblstartval = Quest(EditorIndex).Startval
    Else
        frmQuestEditor.scrlstartnum.value = 1
        frmQuestEditor.scrlstartval.value = 1
        frmQuestEditor.scrlstartnum.Enabled = False
        frmQuestEditor.scrlstartval.Enabled = False
        frmQuestEditor.lblstartitem = "Desactivado"
        frmQuestEditor.lblstartval = "Desactivado"
End If
End Sub

Private Sub ExitMenu_Click(index As Integer)
Call QuestEditorCancel
End Sub

Private Sub Form_Load()
frmQuestEditor.scrlstartnum.max = MAX_ITEMS
frmQuestEditor.lblrewitem.Caption = scrlrewitem.value & ":" & Item(scrlrewitem.value).name
frmQuestEditor.lblquestitem.Caption = frmQuestEditor.scrlquestitem.value & ":" & Item(scrlquestitem.value).name
End Sub

Private Sub salirsinguardar_Click()
Call QuestEditorCancel
End Sub

Private Sub SaveMenu_Click(index As Integer)
Call QuestEditorOk
End Sub

Private Sub scrllvl_Change()
frmQuestEditor.lblLevel.Caption = scrllvl.value
End Sub

Private Sub scrlQuestExpReward_Change()
frmQuestEditor.lbExpReward.Caption = frmQuestEditor.scrlQuestExpReward.value
End Sub

Private Sub scrlquestitem_Change()
frmQuestEditor.lblquestitem.Caption = frmQuestEditor.scrlquestitem.value & ":" & Item(scrlquestitem.value).name
End Sub

Private Sub scrlquestvalue_Change()
frmQuestEditor.lblquestval.Caption = frmQuestEditor.scrlquestvalue.value
End Sub

Private Sub scrlrewitem_Change()
frmQuestEditor.lblrewitem.Caption = scrlrewitem.value & ":" & Item(scrlrewitem.value).name
End Sub

Private Sub scrlrewval_Change()
frmQuestEditor.lblrewval.Caption = frmQuestEditor.scrlrewval.value
End Sub


Private Sub scrlstartnum_Change()
frmQuestEditor.lblstartitem.Caption = scrlstartnum.value & ":" & Item(scrlstartnum.value).name
End Sub

Private Sub scrlstartval_Change()
frmQuestEditor.lblstartval.Caption = frmQuestEditor.scrlstartval.value
End Sub

Private Sub Timer1_Timer()

If frmQuestEditor.chkstart.value = 1 Then
If Item(frmQuestEditor.scrlstartnum.value).Type = 16 Or Item(frmQuestEditor.scrlstartnum.value).Stackable > 0 Then
frmQuestEditor.scrlstartval.Enabled = True
frmQuestEditor.lblstartval.Caption = frmQuestEditor.scrlstartval.value
Else
frmQuestEditor.scrlstartval.Enabled = False
frmQuestEditor.lblstartval.Caption = "1"
End If
End If

If Item(frmQuestEditor.scrlstartnum.value).Type = 16 Or Item(frmQuestEditor.scrlstartnum.value).Stackable > 0 Then
frmQuestEditor.scrlquestvalue.Enabled = True
frmQuestEditor.lblquestval.Caption = frmQuestEditor.scrlquestvalue.value
Else
frmQuestEditor.scrlquestvalue.Enabled = False
frmQuestEditor.lblquestval.Caption = "1"
End If

If Item(frmQuestEditor.scrlstartnum.value).Type = 16 Or Item(frmQuestEditor.scrlstartnum.value).Stackable > 0 Then
frmQuestEditor.scrlrewval.Enabled = True
frmQuestEditor.lblrewval.Caption = frmQuestEditor.scrlrewval.value
Else
frmQuestEditor.scrlrewval.Enabled = False
frmQuestEditor.lblrewval.Caption = "1"
End If

End Sub
