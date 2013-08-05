VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Npcs (Index: 0)"
   ClientHeight    =   8385
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9705
   ClipControls    =   0   'False
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
   Icon            =   "frmNpcEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   559
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   1200
      Top             =   7680
   End
   Begin VB.Frame Frame4 
      Caption         =   "Quest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   4920
      TabIndex        =   57
      Top             =   3360
      Width           =   4695
      Begin VB.HScrollBar scrlquest 
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1200
         TabIndex        =   59
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Configuración"
      ClipControls    =   0   'False
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
      Height          =   3135
      Left            =   4920
      TabIndex        =   41
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chksstill 
         Caption         =   "Inmovil"
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
         Left            =   2280
         TabIndex        =   67
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkNight 
         Caption         =   "Night"
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
         Left            =   3840
         TabIndex        =   47
         Top             =   600
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Day"
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
         Left            =   3240
         TabIndex        =   46
         Top             =   600
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.ComboBox cmbBehavior 
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
         ItemData        =   "frmNpcEditor.frx":0FC2
         Left            =   240
         List            =   "frmNpcEditor.frx":0FDB
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtSpawnSecs 
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
         Height          =   255
         Left            =   240
         MaxLength       =   10
         TabIndex        =   44
         Text            =   "0"
         Top             =   600
         Width           =   1695
      End
      Begin VB.HScrollBar scrlScript 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   43
         Top             =   2520
         Width           =   4215
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   42
         Top             =   1920
         Value           =   1
         Width           =   4215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Movimiento:"
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
         Left            =   2280
         TabIndex        =   66
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Comportamiento:"
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
         TabIndex        =   54
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Tiempo de Respawn (Segundos)"
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
         TabIndex        =   53
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Spawn según tiempo:"
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
         Left            =   3240
         TabIndex        =   52
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label lblScript 
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
         TabIndex        =   51
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblScriptNum 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   50
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Elemento:"
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
         TabIndex        =   49
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label lblElement 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ninguno"
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
         Left            =   3900
         TabIndex        =   48
         Top             =   1680
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Soltar Objetos"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   2895
      Left            =   4920
      TabIndex        =   25
      Top             =   4680
      Width           =   4695
      Begin VB.HScrollBar scrlChance 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   55
         Top             =   2400
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   28
         Top             =   1800
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   27
         Top             =   1200
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlDropItem 
         Height          =   255
         Left            =   240
         Max             =   5
         Min             =   1
         TabIndex        =   26
         Top             =   480
         Value           =   1
         Width           =   4215
      End
      Begin VB.Label lblChance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1 In X"
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
         Left            =   3240
         TabIndex        =   56
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Posibilidad:"
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
         TabIndex        =   36
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   4005
         TabIndex        =   35
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   4000
         TabIndex        =   33
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label9 
         Caption         =   "Numero:"
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
         TabIndex        =   32
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblItemName 
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
         Left            =   960
         TabIndex        =   31
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblDropItem 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   4005
         TabIndex        =   30
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label13 
         Caption         =   "Soltar :"
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
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información General"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox BigNpc 
         Caption         =   "NPC Grande"
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
         Left            =   2040
         TabIndex        =   62
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Opt32 
         Caption         =   "32x32"
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
         Left            =   2280
         TabIndex        =   61
         Top             =   2280
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Opt64 
         Caption         =   "64x32"
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
         Left            =   2280
         TabIndex        =   60
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtName 
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
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtAttackSay 
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
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   720
         Width           =   3615
      End
      Begin VB.HScrollBar ExpGive 
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   6960
         Width           =   4215
      End
      Begin VB.HScrollBar StartHP 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   6360
         Width           =   4215
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   6
         Top             =   1320
         Width           =   4215
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   5
         Top             =   3360
         Width           =   4215
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   4
         Top             =   3960
         Width           =   4215
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   3
         Top             =   4560
         Width           =   4215
      End
      Begin VB.HScrollBar scrlSPEED 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   2
         Top             =   5160
         Width           =   4215
      End
      Begin VB.HScrollBar scrlMAGI 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   1
         Top             =   5760
         Width           =   4215
      End
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   63
         Top             =   2160
         Width           =   480
         Begin VB.PictureBox picSprites 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   12000
            Left            =   120
            ScaleHeight     =   800
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   456
            TabIndex        =   64
            Top             =   120
            Width           =   6840
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   3240
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   78
         TabIndex        =   65
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Conversar:"
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
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblExpGiven 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   23
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label lblNpcExp 
         Caption         =   "Experiencia:"
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
         TabIndex        =   22
         Top             =   6720
         Width           =   735
      End
      Begin VB.Label lblStartHP 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   21
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblNpcHP 
         Caption         =   "Puntos de Golpe:"
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
         TabIndex        =   19
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblSprite 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   18
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblNpcSprite 
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
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   16
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblNpcSight 
         BackStyle       =   0  'Transparent
         Caption         =   "Vista del NPC (Por Tile):"
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
         TabIndex        =   15
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   14
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblNpcStr 
         Caption         =   "Fuerza:"
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
         TabIndex        =   13
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   12
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label lblNpcDef 
         Caption         =   "Defensa:"
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
         TabIndex        =   11
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   10
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblNpcSpd 
         Caption         =   "Velocidad:"
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
         TabIndex        =   9
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         TabIndex        =   8
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label lblNpcMagic 
         Caption         =   "Magia:"
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
         TabIndex        =   7
         Top             =   5520
         Width           =   495
      End
   End
   Begin Eclipse.jcbutton cmdOk 
      Height          =   495
      Left            =   2760
      TabIndex        =   68
      Top             =   7800
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
      Caption         =   "Aceptar"
      PictureNormal   =   "frmNpcEditor.frx":1045
      PictureHot      =   "frmNpcEditor.frx":1829
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Eclipse.jcbutton cmdCancel 
      Height          =   495
      Left            =   4920
      TabIndex        =   69
      Top             =   7800
      Width           =   2295
      _ExtentX        =   4048
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
      PictureNormal   =   "frmNpcEditor.frx":200D
      PictureHot      =   "frmNpcEditor.frx":2961
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BigNpc_Click()
    frmNpcEditor.ScaleMode = 3

    If BigNpc.value = Checked Then
        picSprite.Width = 960
        picSprite.Height = 960
        picSprite.top = 1900
        picSprite.Left = 3360

        picSprites.Picture = LoadPicture(App.Path & "\GFX\BigSprites.bmp")
    Else
        picSprite.Width = 480
        picSprite.Left = 3600

        If Opt64.value Then
            picSprite.Height = 960
            picSprite.top = 1900
        Else
            picSprite.Height = 480
            picSprite.top = 2160
        End If

        picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")
    End If
End Sub

Private Sub chkDay_Click()
    If chkNight.value = Unchecked Then
        If chkDay.value = Unchecked Then
            chkDay.value = Checked
        End If
    End If
End Sub

Private Sub chkNight_Click()
    If chkDay.value = Unchecked Then
        If chkNight.value = Unchecked Then
            chkNight.value = Checked
        End If
    End If
End Sub

Private Sub ExpGive_Change()
    lblExpGiven.Caption = CStr(ExpGive.value)
End Sub

Private Sub Form_Load()
    frmNpcEditor.Caption = "Editor de NPCS (Index: " & EditorIndex & ")"

    scrlElement.max = MAX_ELEMENTS
    scrlDropItem.max = MAX_NPC_DROPS

    If SpriteSize = 1 Then
        picSprite.Height = 960
    End If

    picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    If Opt64.value = Checked Then
        picSprite.Height = 960
        picSprite.top = 1900
    Else
        picSprite.Height = 480
        picSprite.top = 2160
    End If
End Sub

Private Sub Opt32_Click()
    If Not BigNpc.value = Checked Then
        picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

        picSprite.Height = 480
        picSprite.top = 2160
    End If
End Sub

Private Sub Opt64_Click()
    If Not BigNpc.value = Checked Then
        picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

        picSprite.Height = 960
        picSprite.top = 1900
    End If
End Sub

Private Sub scrlDropItem_Change()
    scrlNum.value = Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemNum
    scrlValue.value = Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemValue
    scrlChance.value = Npc(EditorIndex).ItemNPC(scrlDropItem.value).chance

    lblDropItem.Caption = CStr(scrlDropItem.value)
End Sub

Private Sub scrlElement_Change()
    lblElement.Caption = CStr(scrlElement.value)
End Sub

Private Sub scrlquest_Change()
If frmNpcEditor.scrlquest.value > 0 Then
frmNpcEditor.Label21.Caption = "Quest: " & Trim(Quest(frmNpcEditor.scrlquest.value).name)
Else
frmNpcEditor.Label21.Caption = "Ninguna"
End If
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = CStr(scrlSprite.value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.value)
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = CStr(scrlSTR.value)
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = CStr(scrlDEF.value)
End Sub

Private Sub scrlSPEED_Change()
    lblSPEED.Caption = CStr(scrlSPEED.value)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = CStr(scrlMAGI.value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = CStr(scrlNum.value)
    lblItemName.Caption = vbNullString

    If scrlNum.value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.value).name)
    End If

    Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemNum = scrlNum.value
End Sub

Private Sub scrlValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemValue = scrlValue.value

    lblValue.Caption = CStr(scrlValue.value)
End Sub

Private Sub cmdOk_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub StartHP_Change()
    lblStartHP.Caption = CStr(StartHP.value)
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub

Private Sub scrlChance_Change()
    lblChance.Caption = "1 entre " & scrlChance.value
    Npc(EditorIndex).ItemNPC(scrlDropItem.value).chance = scrlChance.value
End Sub


Private Sub cmbBehavior_Click()
    If cmbBehavior.ListIndex = NPC_BEHAVIOR_SCRIPTED Then
        scrlScript.Enabled = True
    Else
        scrlScript.Enabled = False
    End If
End Sub

Private Sub scrlScript_Change()
    lblScriptNum.Caption = CStr(scrlScript.value)
End Sub

