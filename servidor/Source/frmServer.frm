VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AlterEngine"
   ClientHeight    =   5775
   ClientLeft      =   420
   ClientTop       =   840
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   724
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet tolpene 
      Left            =   6480
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5625
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9922
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   370
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Conversación"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmrChatLogs"
      Tab(0).Control(1)=   "picCMsg"
      Tab(0).Control(2)=   "SSTab2"
      Tab(0).Control(3)=   "fraChatOpt"
      Tab(0).Control(4)=   "lblLogTime"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Jugadores"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBanPlayerReason"
      Tab(1).Control(1)=   "cmdKickPlayerReason"
      Tab(1).Control(2)=   "picMessage"
      Tab(1).Control(3)=   "picKick"
      Tab(1).Control(4)=   "picWarp"
      Tab(1).Control(5)=   "picJail"
      Tab(1).Control(6)=   "picStats"
      Tab(1).Control(7)=   "picBan"
      Tab(1).Control(8)=   "Check1"
      Tab(1).Control(9)=   "lvUsers"
      Tab(1).Control(10)=   "Frame5"
      Tab(1).Control(11)=   "Command66"
      Tab(1).Control(12)=   "cmdGiveAccess"
      Tab(1).Control(13)=   "TPO"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Panel de Control"
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "versionnueva"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "News"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Socket(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "PlayerTimer"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "tmrShutdown"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "tmrGameAI"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "tmrSpawnMapItems"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "tmrPlayerSave"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Frame6"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Frame9"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "picExp"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "picMap"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "picWeather"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Timer1"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Script"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Time"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "picWarpAll"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "AlterEngine"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblTopic"
      Tab(3).Control(1)=   "lblContent"
      Tab(3).Control(2)=   "Image1"
      Tab(3).Control(3)=   "lstTopics"
      Tab(3).Control(4)=   "txtTopic"
      Tab(3).Control(5)=   "Frame7"
      Tab(3).Control(6)=   "Frame8"
      Tab(3).ControlCount=   7
      Begin VB.Frame Frame8 
         Caption         =   "Creditos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2415
         Left            =   -69600
         TabIndex        =   204
         Top             =   3000
         Width           =   5055
         Begin VB.Label Label12 
            Caption         =   "AE Team: 6dragon6 - Ludipe - Ellesar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   210
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label11 
            Caption         =   "Queda totalmente prohibido vender total o parcialmente                  cualquier rasgo de AlterEngine."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   207
            Top             =   1680
            Width           =   4815
         End
         Begin VB.Label Label7 
            Caption         =   "Javier ""Stream"" Cantó | javiloveyou@gmail.com"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   206
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label6 
            Caption         =   "Basado en Eclipse Source"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1200
            TabIndex        =   205
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ultima Versión"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2415
         Left            =   -74760
         TabIndex        =   197
         Top             =   3000
         Width           =   4815
         Begin VB.TextBox ae_versionactual 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   200
            Text            =   "v1.8.2"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox ae_actualizaciones 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   198
            Top             =   600
            Width           =   1815
         End
         Begin Server.jcbutton tieneslaultima 
            Height          =   615
            Left            =   720
            TabIndex        =   202
            Top             =   1320
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            ButtonStyle     =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   15199212
            Caption         =   "Tienes la ultima versión."
            PictureNormal   =   "frmServer.frx":170FA
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            PicturePushOnHover=   -1  'True
            CaptionEffects  =   0
            ToolTip         =   "Tienes la ultima versión :)"
            TooltipType     =   1
            TooltipIcon     =   1
            TooltipBackColor=   0
         End
         Begin Server.jcbutton notienes 
            Height          =   615
            Left            =   720
            TabIndex        =   203
            Top             =   1320
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            ButtonStyle     =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   15199212
            Caption         =   "Debes actualizar tu versión!"
            PictureNormal   =   "frmServer.frx":178DE
            PictureHot      =   "frmServer.frx":18232
            PictureEffectOnOver=   0
            PictureEffectOnDown=   0
            PicturePushOnHover=   -1  'True
            CaptionEffects  =   0
            ToolTip         =   "Ves a www.alterengine.net para actualizar a la ultima versión de AlterEngine."
            TooltipType     =   1
            TooltipIcon     =   1
            TooltipBackColor=   0
         End
         Begin VB.Label Label3 
            Caption         =   "Versión que tu usas:"
            Height          =   375
            Left            =   240
            TabIndex        =   201
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Ultima Versión:"
            Height          =   375
            Left            =   2760
            TabIndex        =   199
            Top             =   360
            Width           =   1215
         End
      End
      Begin Server.jcbutton cmdBanPlayerReason 
         Height          =   375
         Left            =   -66480
         TabIndex        =   176
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Banear Jugador  "
         PictureNormal   =   "frmServer.frx":18B86
         PictureHot      =   "frmServer.frx":194DA
         CaptionEffects  =   4
         ColorScheme     =   2
      End
      Begin Server.jcbutton cmdKickPlayerReason 
         Height          =   375
         Left            =   -66480
         TabIndex        =   175
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
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
         PictureNormal   =   "frmServer.frx":19E2E
         PictureHot      =   "frmServer.frx":1A782
         CaptionEffects  =   4
         ColorScheme     =   2
      End
      Begin VB.PictureBox picWarpAll 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   480
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlMY 
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMX 
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMM 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   61
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin Server.jcbutton Command37 
            Height          =   375
            Left            =   120
            TabIndex        =   168
            Top             =   2040
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
            Caption         =   "Mover"
            PictureNormal   =   "frmServer.frx":1B0D6
            PictureHot      =   "frmServer.frx":1BA2A
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command38 
            Height          =   375
            Left            =   1680
            TabIndex        =   169
            Top             =   2040
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
            PictureNormal   =   "frmServer.frx":1C37E
            PictureHot      =   "frmServer.frx":1CCD2
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label lblMY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label lblMX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblMM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.PictureBox picMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -72360
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   247
         TabIndex        =   144
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtPlayerMsg 
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
            Left            =   120
            TabIndex        =   145
            Top             =   360
            Width           =   3435
         End
         Begin Server.jcbutton cmdServMsg 
            Height          =   375
            Left            =   240
            TabIndex        =   185
            Top             =   720
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
            Caption         =   "Enviar"
            PictureNormal   =   "frmServer.frx":1D626
            PictureHot      =   "frmServer.frx":1DF7A
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdMsgCancel 
            Height          =   375
            Left            =   2040
            TabIndex        =   186
            Top             =   720
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":1E8CE
            PictureHot      =   "frmServer.frx":1F222
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label lblMessage 
            Caption         =   "Mensaje a enviar al jugador:"
            Height          =   240
            Left            =   120
            TabIndex        =   146
            Top             =   120
            Width           =   2775
         End
      End
      Begin VB.PictureBox picKick 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -72360
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   141
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtKickReason 
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
            Left            =   120
            TabIndex        =   143
            Top             =   360
            Width           =   3075
         End
         Begin VB.CheckBox chkKickReason 
            Caption         =   "Con una razón"
            Height          =   240
            Left            =   120
            TabIndex        =   142
            Top             =   120
            Width           =   1695
         End
         Begin Server.jcbutton cmdServKick 
            Height          =   375
            Left            =   120
            TabIndex        =   187
            Top             =   720
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
            Caption         =   "Expulsar"
            PictureNormal   =   "frmServer.frx":1FB76
            PictureHot      =   "frmServer.frx":204CA
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdKickCancel 
            Height          =   375
            Left            =   1680
            TabIndex        =   188
            Top             =   720
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":20E1E
            PictureHot      =   "frmServer.frx":21772
            CaptionEffects  =   4
            ColorScheme     =   2
         End
      End
      Begin VB.PictureBox picWarp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   -72120
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   131
         Top             =   1320
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlWarpMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   136
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.HScrollBar scrlWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CheckBox chkWarpReason 
            Caption         =   "Con una razón"
            Height          =   240
            Left            =   120
            TabIndex        =   133
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtWarpReason 
            Height          =   285
            Left            =   120
            TabIndex        =   132
            Top             =   2280
            Width           =   3135
         End
         Begin Server.jcbutton cmdServWarp 
            Height          =   375
            Left            =   120
            TabIndex        =   189
            Top             =   2640
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
            Caption         =   "Mover"
            PictureNormal   =   "frmServer.frx":220C6
            PictureHot      =   "frmServer.frx":22A1A
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdWarpCancel 
            Height          =   375
            Left            =   1800
            TabIndex        =   190
            Top             =   2640
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":2336E
            PictureHot      =   "frmServer.frx":23CC2
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label lblWarpMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   139
            Top             =   120
            Width           =   585
         End
         Begin VB.Label lblWarpX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   138
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblWarpY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   137
            Top             =   1320
            Width           =   285
         End
      End
      Begin VB.Frame Time 
         Caption         =   "Tiempo del juego"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1095
         Left            =   360
         TabIndex        =   122
         Top             =   4320
         Width           =   9975
         Begin VB.TextBox txtTimeS 
            Height          =   285
            Left            =   1320
            TabIndex        =   128
            Text            =   "1"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtTimeM 
            Height          =   285
            Left            =   720
            TabIndex        =   127
            Text            =   "1"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtTimeH 
            Height          =   285
            Left            =   120
            TabIndex        =   126
            Text            =   "1"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox GameTimeSpeed 
            Height          =   285
            Left            =   7320
            TabIndex        =   123
            Text            =   "1"
            Top             =   720
            Width           =   495
         End
         Begin Server.jcbutton cmdSetTime 
            Height          =   255
            Left            =   1920
            TabIndex        =   161
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
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
            Caption         =   "Poner en hora"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command68 
            Height          =   255
            Left            =   7920
            TabIndex        =   162
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
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
            Caption         =   "Cambiar Velocidad"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command69 
            Height          =   375
            Left            =   3960
            TabIndex        =   163
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "Desactivar Tiempo"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   3600
            TabIndex        =   125
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Velocidad:"
            Height          =   255
            Left            =   6240
            TabIndex        =   124
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.PictureBox Script 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   271
         TabIndex        =   118
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
         Begin CodeSenseCtl.CodeSense ServerScript 
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmServer.frx":24616
            TabIndex        =   119
            Top             =   240
            Width           =   3705
         End
         Begin Server.jcbutton Command72 
            Height          =   495
            Left            =   240
            TabIndex        =   170
            Top             =   1920
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
            Caption         =   "Ejecutar"
            PictureNormal   =   "frmServer.frx":2477C
            PictureHot      =   "frmServer.frx":250D0
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command71 
            Height          =   495
            Left            =   2040
            TabIndex        =   171
            Top             =   1920
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":25A24
            PictureHot      =   "frmServer.frx":26378
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Script:"
            Height          =   195
            Left            =   120
            TabIndex        =   120
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.TextBox txtTopic 
         Height          =   3570
         Left            =   -72120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   116
         Top             =   6600
         Width           =   7575
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6960
         Top             =   360
      End
      Begin VB.Timer tmrChatLogs 
         Interval        =   1000
         Left            =   -65160
         Top             =   0
      End
      Begin VB.PictureBox picWeather 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   2640
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   101
         Top             =   7440
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton Command65 
            Caption         =   "Snow"
            Height          =   255
            Left            =   1680
            TabIndex        =   109
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command64 
            Caption         =   "Rain"
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command63 
            Caption         =   "Thunder"
            Height          =   255
            Left            =   1680
            TabIndex        =   107
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command62 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   1080
            Width           =   1335
         End
         Begin VB.HScrollBar scrlRainIntensity 
            Height          =   255
            Left            =   120
            Max             =   50
            Min             =   1
            TabIndex        =   104
            Top             =   360
            Value           =   25
            Width           =   2895
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1560
            TabIndex        =   102
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Weather: None"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   720
            Width           =   1710
         End
         Begin VB.Label lblRainIntensity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intensity: 25"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.PictureBox picCMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   -74400
         ScaleHeight     =   1905
         ScaleWidth      =   3345
         TabIndex        =   15
         Top             =   5760
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtMsg 
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
            Top             =   960
            Width           =   3075
         End
         Begin VB.TextBox txtTitle 
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
            MaxLength       =   13
            TabIndex        =   20
            Top             =   360
            Width           =   3075
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   17
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   255
            Left            =   1680
            TabIndex        =   16
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   360
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3015
         Left            =   -74520
         TabIndex        =   91
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   353
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Principal"
         TabPicture(0)   =   "frmServer.frx":26CCC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtText(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtChat"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Difusión"
         TabPicture(1)   =   "frmServer.frx":26CE8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtText(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Global"
         TabPicture(2)   =   "frmServer.frx":26D04
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtText(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Mapa"
         TabPicture(3)   =   "frmServer.frx":26D20
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtText(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Privado"
         TabPicture(4)   =   "frmServer.frx":26D3C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtText(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Admin"
         TabPicture(5)   =   "frmServer.frx":26D58
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtText(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Emoticonos"
         TabPicture(6)   =   "frmServer.frx":26D74
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "txtText(6)"
         Tab(6).ControlCount=   1
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   6
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   99
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   5
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   4
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   97
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   3
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   96
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   2
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   95
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   1
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Top             =   360
            Width           =   9135
         End
         Begin VB.TextBox txtChat 
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
            TabIndex        =   93
            Top             =   2640
            Width           =   9375
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2250
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   360
            Width           =   9375
         End
      End
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   480
         ScaleHeight     =   223
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   231
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ListBox lstNPC 
            Height          =   2400
            Left            =   1680
            TabIndex        =   87
            Top             =   360
            Width           =   1575
         End
         Begin Server.jcbutton Command41 
            Height          =   375
            Left            =   1680
            TabIndex        =   172
            Top             =   2880
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
            Caption         =   "Cerrar"
            PictureNormal   =   "frmServer.frx":26D90
            PictureHot      =   "frmServer.frx":276E4
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPCs"
            Height          =   195
            Index           =   13
            Left            =   1680
            TabIndex        =   88
            Top             =   120
            Width           =   375
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dentro:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   86
            Top             =   3000
            Width           =   555
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiendas:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   85
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootY:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   84
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootX:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   83
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootMapa:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   82
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Musica:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   81
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Derecha:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   80
            Top             =   1560
            Width           =   660
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Izquierda:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   79
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abajo:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   78
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arriba:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   77
            Top             =   840
            Width           =   495
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moral:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   600
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revision:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   660
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   390
         End
      End
      Begin VB.PictureBox picExp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   6600
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   67
         Top             =   7320
         Visible         =   0   'False
         Width           =   3255
         Begin VB.HScrollBar scrlExp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   147
            Top             =   360
            Value           =   1
            Width           =   3015
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1560
            TabIndex        =   70
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Execute"
            Height          =   255
            Left            =   1560
            TabIndex        =   68
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblMassExp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experience: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Lista de mapas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2415
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   4095
         Begin VB.ListBox MapList 
            Appearance      =   0  'Flat
            Height          =   1785
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Comandos Masivos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   6840
         TabIndex        =   57
         Top             =   2880
         Width           =   3495
         Begin VB.CommandButton Command33 
            Caption         =   "Mass Experience"
            Height          =   255
            Left            =   1680
            TabIndex        =   59
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Mass Heal"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   2040
            Width           =   1575
         End
         Begin Server.jcbutton Command34 
            Height          =   375
            Left            =   1800
            TabIndex        =   164
            Top             =   840
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
            Caption         =   "Subir 1 Nivel"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command31 
            Height          =   375
            Left            =   120
            TabIndex        =   165
            Top             =   840
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
            Caption         =   "Matar a todos"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command32 
            Height          =   375
            Left            =   1800
            TabIndex        =   166
            Top             =   360
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
            Caption         =   "Mover a todos"
            CaptionEffects  =   4
            ToolTip         =   "Mueve a todos los jugadores a un punto que tu quieras."
            TooltipType     =   1
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command9 
            Height          =   375
            Left            =   120
            TabIndex        =   167
            Top             =   360
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
            Caption         =   "Expulsar a todos"
            CaptionEffects  =   4
            ToolTip         =   "Expulsa a todos los jugadores conectados al servidor."
            TooltipType     =   1
            ColorScheme     =   2
         End
      End
      Begin VB.ListBox lstTopics 
         Height          =   3570
         Left            =   -74760
         TabIndex        =   55
         Top             =   6720
         Width           =   2175
      End
      Begin VB.PictureBox picJail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   -72240
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtJailReason 
            Height          =   285
            Left            =   120
            TabIndex        =   130
            Top             =   2280
            Width           =   3135
         End
         Begin VB.CheckBox chkJailReason 
            Caption         =   "Con una razón"
            Height          =   240
            Left            =   120
            TabIndex        =   129
            Top             =   2040
            Width           =   1815
         End
         Begin VB.HScrollBar scrlJailY 
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlJailX 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlJailMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   49
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin Server.jcbutton cmdServJail 
            Height          =   375
            Left            =   120
            TabIndex        =   191
            Top             =   2640
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
            Caption         =   "Encarcelar"
            PictureNormal   =   "frmServer.frx":28038
            PictureHot      =   "frmServer.frx":2898C
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdJailCancel 
            Height          =   375
            Left            =   1800
            TabIndex        =   192
            Top             =   2640
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":292E0
            PictureHot      =   "frmServer.frx":29C34
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label lblJailY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label lblJailX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblJailMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   585
         End
      End
      Begin VB.PictureBox picStats 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   -72960
         ScaleHeight     =   215
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   4695
         Begin Server.jcbutton Command8 
            Height          =   375
            Left            =   3000
            TabIndex        =   193
            Top             =   2760
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":2A588
            PictureHot      =   "frmServer.frx":2AEDC
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index:"
            Height          =   195
            Index           =   20
            Left            =   2400
            TabIndex        =   47
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   46
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magi:"
            Height          =   195
            Index           =   18
            Left            =   2400
            TabIndex        =   45
            Top             =   1320
            Width           =   390
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed:"
            Height          =   195
            Index           =   17
            Left            =   2400
            TabIndex        =   44
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Def:"
            Height          =   195
            Index           =   16
            Left            =   2400
            TabIndex        =   43
            Top             =   840
            Width           =   315
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str:"
            Height          =   195
            Index           =   15
            Left            =   2400
            TabIndex        =   42
            Top             =   600
            Width           =   270
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Access:"
            Height          =   195
            Index           =   14
            Left            =   2400
            TabIndex        =   41
            Top             =   360
            Width           =   945
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild:"
            Height          =   195
            Index           =   13
            Left            =   2400
            TabIndex        =   40
            Top             =   120
            Width           =   405
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   39
            Top             =   3000
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   38
            Top             =   2760
            Width           =   330
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprite:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   37
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PK:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXP: /"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   33
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SP: /"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP: /"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HP: /"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Character:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   780
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox picBan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   -72360
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox chkBanReason 
            Caption         =   "Con una razón"
            Height          =   240
            Left            =   120
            TabIndex        =   140
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtBanReason 
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
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   3075
         End
         Begin Server.jcbutton cmdServBan 
            Height          =   375
            Left            =   120
            TabIndex        =   194
            Top             =   720
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
            Caption         =   "Banear"
            PictureNormal   =   "frmServer.frx":2B830
            PictureHot      =   "frmServer.frx":2C184
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdBanCancel 
            Height          =   375
            Left            =   1680
            TabIndex        =   195
            Top             =   720
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
            Caption         =   "Cancelar"
            PictureNormal   =   "frmServer.frx":2CAD8
            PictureHot      =   "frmServer.frx":2D42C
            CaptionEffects  =   4
            ColorScheme     =   2
         End
      End
      Begin VB.Frame fraChatOpt 
         Caption         =   "Registros del chat (logs)"
         Height          =   855
         Left            =   -73680
         TabIndex        =   9
         Top             =   3840
         Width           =   8055
         Begin VB.CheckBox chkLogAdmin 
            Caption         =   "Admins"
            Height          =   255
            Left            =   4680
            TabIndex        =   23
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkLogGlobal 
            Caption         =   "Global"
            Height          =   255
            Left            =   3840
            TabIndex        =   22
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkLogPM 
            Caption         =   "Privado"
            Height          =   255
            Left            =   2880
            TabIndex        =   13
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkLogMap 
            Caption         =   "Mapa"
            Height          =   255
            Left            =   2160
            TabIndex        =   12
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkLogEmote 
            Caption         =   "Emoticono"
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkLogBC 
            Caption         =   "Difusion"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin Server.jcbutton cmdSaveLogs 
            Height          =   495
            Left            =   5640
            TabIndex        =   196
            Top             =   240
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
            Caption         =   "Guardar Registros"
            PictureNormal   =   "frmServer.frx":2DD80
            PictureHot      =   "frmServer.frx":2E6D4
            CaptionEffects  =   4
            ColorScheme     =   2
         End
      End
      Begin VB.Timer tmrPlayerSave 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   7440
         Top             =   240
      End
      Begin VB.Timer tmrSpawnMapItems 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   9480
         Top             =   240
      End
      Begin VB.Timer tmrGameAI 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   8880
         Top             =   360
      End
      Begin VB.Timer tmrShutdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   8400
         Top             =   240
      End
      Begin VB.Timer PlayerTimer 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   7920
         Top             =   360
      End
      Begin VB.Frame Frame3 
         Caption         =   "Classes"
         Height          =   1335
         Left            =   6600
         TabIndex        =   6
         Top             =   6240
         Width           =   1935
         Begin VB.CommandButton Command30 
            Caption         =   "Edit"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Reload"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   6255
         Begin VB.CommandButton Command59 
            Caption         =   "Weather"
            Height          =   375
            Left            =   3120
            TabIndex        =   121
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox chkChat 
            Caption         =   "Guardar Logs"
            Height          =   255
            Left            =   4800
            TabIndex        =   100
            Top             =   960
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox mnuServerLog 
            Caption         =   "Logs del server"
            Height          =   255
            Left            =   1560
            TabIndex        =   90
            Top             =   960
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox Closed 
            Caption         =   "Servidor Cerrado"
            Height          =   255
            Left            =   3120
            TabIndex        =   89
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox GMOnly 
            Caption         =   "Solo Admins"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
         Begin Server.jcbutton Command36 
            Height          =   375
            Left            =   120
            TabIndex        =   158
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Ver Información"
            PictureNormal   =   "frmServer.frx":2F028
            PictureHot      =   "frmServer.frx":2F97C
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command35 
            Height          =   375
            Left            =   2160
            TabIndex        =   159
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Lista de mapas"
            PictureNormal   =   "frmServer.frx":302D0
            PictureHot      =   "frmServer.frx":31324
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command1 
            Height          =   375
            Left            =   4200
            TabIndex        =   160
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Apagar Servidor"
            PictureNormal   =   "frmServer.frx":32378
            PictureHot      =   "frmServer.frx":32CCC
            CaptionEffects  =   4
            ColorScheme     =   2
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scripts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2415
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   1815
         Begin Server.jcbutton Command25 
            Height          =   495
            Left            =   120
            TabIndex        =   150
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            ButtonStyle     =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14935011
            Caption         =   "Recargar"
            PictureNormal   =   "frmServer.frx":33620
            PictureHot      =   "frmServer.frx":33F74
            CaptionEffects  =   3
            ToolTip         =   "Si has realizado alguna modificación de scripts, esto los recarga sin necesidad de apagar el servidor."
            TooltipType     =   1
            TooltipIcon     =   1
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command26 
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
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
            Caption         =   "Activar Scripts"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command27 
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
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
            Caption         =   "Apagar Scripts"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command28 
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
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
            Caption         =   "Editar Scripts"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command70 
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
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
            Caption         =   "Ejecutar Script"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label lblScriptOn 
            Alignment       =   2  'Center
            Caption         =   "Scripts: (...)"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Rejilla"
         Height          =   255
         Left            =   -74640
         TabIndex        =   2
         Top             =   4920
         Value           =   1  'Checked
         Width           =   975
      End
      Begin MSWinsockLib.Winsock Socket 
         Index           =   0
         Left            =   9960
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   3855
         Left            =   -74640
         TabIndex        =   1
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6800
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Personaje"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nivel"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sprite"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Privilegio"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Frame News 
         Caption         =   "Noticias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2415
         Left            =   6240
         TabIndex        =   111
         Top             =   360
         Width           =   1935
         Begin Server.jcbutton btnEventEdit 
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
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
            Caption         =   "Editar Evento"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton btnEventSend 
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   1920
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
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
            Caption         =   "Enviar Evento"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command46 
            Height          =   375
            Left            =   120
            TabIndex        =   156
            Top             =   360
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
            Caption         =   "Editar Noticia"
            PictureNormal   =   "frmServer.frx":348C8
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton Command73 
            Height          =   375
            Left            =   120
            TabIndex        =   157
            Top             =   840
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
            Caption         =   "Enviar Noticia"
            PictureNormal   =   "frmServer.frx":3521C
            PictureHot      =   "frmServer.frx":35B70
            CaptionEffects  =   4
            ColorScheme     =   2
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Editores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2415
         Left            =   8280
         TabIndex        =   113
         Top             =   360
         Width           =   2295
         Begin Server.jcbutton editorclases 
            Height          =   375
            Left            =   240
            TabIndex        =   208
            Top             =   360
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
            Caption         =   "Editar Clases"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton editorconfig 
            Height          =   375
            Left            =   240
            TabIndex        =   209
            Top             =   840
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
            Caption         =   "Editar Configuración"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton envanuncio 
            Height          =   375
            Left            =   240
            TabIndex        =   211
            Top             =   1320
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
            Caption         =   "Anuncios In-Game"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton BtnColorPjs 
            Height          =   375
            Left            =   240
            TabIndex        =   213
            Top             =   1800
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
            Caption         =   "Colores Pjs"
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin VB.Label lblVer 
            Caption         =   "Build: (...)"
            Height          =   255
            Left            =   960
            TabIndex        =   115
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblEngine 
            Caption         =   "Eclipse Evolution"
            Height          =   255
            Left            =   1200
            TabIndex        =   114
            Top             =   2520
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Acciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   5055
         Left            =   -66600
         TabIndex        =   155
         Top             =   240
         Width           =   2175
         Begin Server.jcbutton cmdJailPlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   177
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Encarcelar            "
            PictureNormal   =   "frmServer.frx":364C4
            PictureHot      =   "frmServer.frx":36E18
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdViewInfo 
            Height          =   375
            Left            =   120
            TabIndex        =   178
            Top             =   1680
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Ver Información "
            PictureNormal   =   "frmServer.frx":3776C
            PictureHot      =   "frmServer.frx":380C0
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdMsgPlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   179
            Top             =   2160
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Mensaje Privado"
            PictureNormal   =   "frmServer.frx":38A14
            PictureHot      =   "frmServer.frx":39368
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdMutePlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   180
            Top             =   2640
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Silenciar               "
            PictureNormal   =   "frmServer.frx":39CBC
            PictureHot      =   "frmServer.frx":3A610
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdUnmutePlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   181
            Top             =   3120
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Desilenciar           "
            PictureNormal   =   "frmServer.frx":3AF64
            PictureHot      =   "frmServer.frx":3B8B8
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdKillPlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   182
            Top             =   3600
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Matar Jugador    "
            PictureNormal   =   "frmServer.frx":3C20C
            PictureHot      =   "frmServer.frx":3CB60
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdHealPlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   183
            Top             =   4080
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Curar Jugador     "
            PictureNormal   =   "frmServer.frx":3D4B4
            PictureHot      =   "frmServer.frx":3DE08
            CaptionEffects  =   4
            ColorScheme     =   2
         End
         Begin Server.jcbutton cmdWarpPlayer 
            Height          =   375
            Left            =   120
            TabIndex        =   184
            Top             =   4560
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Mover Jugador   "
            PictureNormal   =   "frmServer.frx":3E75C
            PictureHot      =   "frmServer.frx":3F0B0
            CaptionEffects  =   4
            ColorScheme     =   2
         End
      End
      Begin Server.jcbutton Command66 
         Height          =   495
         Left            =   -68400
         TabIndex        =   173
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Refrescar"
         PictureNormal   =   "frmServer.frx":3FA04
         PictureHot      =   "frmServer.frx":40358
         CaptionEffects  =   3
         ColorScheme     =   2
      End
      Begin Server.jcbutton cmdGiveAccess 
         Height          =   495
         Left            =   -68400
         TabIndex        =   174
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "Privilegios"
         PictureNormal   =   "frmServer.frx":40CAC
         PictureHot      =   "frmServer.frx":4193C
         CaptionEffects  =   3
      End
      Begin VB.Label versionnueva 
         BackStyle       =   0  'Transparent
         Caption         =   "Una nueva versión de AE ha sido lanzada."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   6360
         TabIndex        =   212
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   2460
         Left            =   -74880
         Picture         =   "frmServer.frx":425CC
         Top             =   360
         Width           =   10335
      End
      Begin VB.Label lblContent 
         Caption         =   "Contents:"
         Height          =   255
         Left            =   -74040
         TabIndex        =   117
         Top             =   6480
         Width           =   735
      End
      Begin VB.Label lblLogTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El chat se guardara! j3j3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72120
         TabIndex        =   110
         Top             =   4800
         Width           =   5775
      End
      Begin VB.Label lblTopic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   56
         Top             =   6480
         Width           =   510
      End
      Begin VB.Label TPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jugadores Conectados:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   -74640
         TabIndex        =   3
         Top             =   600
         Width           =   4050
      End
   End
   Begin VB.Timer tmrScriptedTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   120
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Base original Eclipse Source v2.7 Patch #4


Option Explicit


' Función por la cual comprueba remotamente que tiene la ultima versión.
' Stream ;)

Private Sub actualizar_aw_Click()

ae_actualizaciones = tolpene.OpenURL("http://www.alterengine.net/internet/version.txt")

If ae_actualizaciones.Text = ae_versionactual.Text Then
tieneslaultima.Visible = True
notienes.Visible = False
Else
tieneslaultima.Visible = False
notienes.Visible = True
End If
End Sub

Private Sub BtnColorPjs_Click()
frmColorPjs.Visible = True
End Sub

Private Sub cmdGiveAccess_Click()
    Dim Access As String
    Dim index As Integer
    
    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If IsPlaying(index) Then
        Access = InputBox("¿Que tipo de acceso le daras?" & vbNewLine & vbNewLine & "0 - Jugador" & vbNewLine & "1 - Moderador" & vbNewLine & "2 - Mapeador" & vbNewLine & "3 - Developer" & vbNewLine & "4 - Administrador" & vbNewLine & "5 - Jefe" & vbNewLine, "Dar Privilegio", CStr(Player(index).Char(Player(index).CharNum).Access))
        
        If IsNumeric(Access) Then
            If Val(Access) < 0 Or Val(Access) > 5 Then
                Call MsgBox("Por favor, introduce un valor de 0 a 5.")
                Exit Sub
            End If

            Call SetPlayerAccess(index, Val(Access))
            Select Case Val(Access)
                Case 0
                    Call SetPlayerColor(index, UserCr(1), UserCr(2), UserCr(3))
                
                Case 1
                    Call SetPlayerColor(index, ModCr(1), ModCr(2), ModCr(3))
                
                Case 2
                    Call SetPlayerColor(index, MapperCr(1), MapperCr(2), MapperCr(3))

                Case 3
                    Call SetPlayerColor(index, DeveloperCr(1), DeveloperCr(2), DeveloperCr(3))
                
                Case 4
                    Call SetPlayerColor(index, AdminCr(1), AdminCr(2), AdminCr(3))
                
                Case 5
                    Call SetPlayerColor(index, OwnerCr(1), OwnerCr(2), OwnerCr(3))
            End Select
            Call SendPlayerData(index)

            If GetPlayerAccess(index) > 0 Then
                Call PlayerMsg(index, "Has recibido privilegios administrativos.", AdminColor)
            End If

            Call ShowPLR(index)
        End If
    End If
End Sub

Private Sub btnEventEdit_Click()
    If FileExists("Editor.exe") Then
        Call Shell(App.Path & "\Editor.exe Events.ini", vbNormalNoFocus)
    Else
        Call MsgBox("No se ha encontrado el editor de AE!", vbOKOnly, "Error")
    End If
End Sub

Private Sub cmdJailCancel_Click()
    picJail.Visible = False
End Sub

Private Sub cmdBanCancel_Click()
    picBan.Visible = False
End Sub

Private Sub cmdKickCancel_Click()
    picKick.Visible = False
End Sub

Private Sub cmdMsgCancel_Click()
    picMessage.Visible = False
End Sub

Private Sub cmdServBan_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If chkBanReason.Value = Checked Then
        If LenB(txtBanReason.Text) = 0 Then
            Call MsgBox("Por favor, introduce una razón para banear a este jugador!")
            Exit Sub
        End If

        If IsPlaying(index) Then
            Call BanByServer(index, txtBanReason.Text)
        End If
    Else
        If IsPlaying(index) Then
            Call BanByServer(index, vbNullString)
        End If
    End If

    picBan.Visible = False
End Sub

Private Sub cmdServKick_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If chkKickReason.Value = Checked Then
        If LenB(txtKickReason.Text) = 0 Then
            Call MsgBox("Por favor, introduce una razón para expulsar a este jugador!")
            Exit Sub
        End If

        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido expulsado por el servidor. - Razón: (" & txtWarpReason.Text & ")", WHITE)
            Call AlertMsg(index, "Has sido expulsado del servidor. - Razón: (" & txtKickReason.Text & ")")
        End If
    Else
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido expulsado del servidor.", WHITE)
            Call AlertMsg(index, "Has sido expulsado del servidor.")
        End If
    End If

    picKick.Visible = False
End Sub

Private Sub cmdServMsg_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)
    
    If IsPlaying(index) Then
        Call PlayerMsg(index, "* Mensaje privado del servidor: " & txtPlayerMsg.Text, BRIGHTGREEN)
    End If

    picMessage.Visible = False
End Sub

Private Sub cmdServWarp_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If chkWarpReason.Value = Checked Then
        If LenB(txtWarpReason.Text) = 0 Then
            Call MsgBox("Por favor, introduce una razón para mover a este jugador.")
            Exit Sub
        End If

        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido movido por el servidor. - Razón: (" & txtWarpReason.Text & ")", WHITE)
            Call PlayerWarp(index, scrlWarpMap.Value, scrlWarpX.Value, scrlWarpY.Value)
        End If
    Else
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido movido por el servidor.", WHITE)
            Call PlayerWarp(index, scrlWarpMap.Value, scrlWarpX.Value, scrlWarpY.Value)
        End If
    End If

    picWarp.Visible = False
End Sub

Private Sub cmdServJail_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If chkJailReason.Value = Checked Then
        If LenB(txtJailReason.Text) = 0 Then
            Call MsgBox("Por favor, introduce una razón para encarcelar al jugador!")
            Exit Sub
        End If

        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido encarcelado por el servidor. Razón: (" & txtJailReason.Text & ")", WHITE)
            Call PlayerWarp(index, scrlJailMap.Value, scrlJailX.Value, scrlJailY.Value)
        End If
    Else
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido encarcelado por el servidor.", WHITE)
            Call PlayerWarp(index, scrlJailMap.Value, scrlJailX.Value, scrlJailY.Value)
        End If
    End If

    picJail.Visible = False
End Sub

Private Sub cmdSetTime_Click()
    Dim TimeH As Integer
    Dim TimeM As Integer
    Dim TimeS As Integer

    TimeH = Val(txtTimeH.Text)
    TimeM = Val(txtTimeM.Text)
    TimeS = Val(txtTimeS.Text)
    
    If TimeH < 1 Or TimeH > 24 Then
        Exit Sub
    End If
    
    If TimeM < 0 Or TimeM > 59 Then
        Exit Sub
    End If
    
    If TimeS < 0 Or TimeS > 59 Then
        Exit Sub
    End If
    
    If TimeH = 24 And (TimeM > 0 Or TimeS > 0) Then
        Exit Sub
    End If

    Hours = TimeH
    Minutes = TimeM
    Seconds = TimeS

    SendGameClockToAll
End Sub

Private Sub cmdWarpCancel_Click()
    picWarp.Visible = False
End Sub

Private Sub Command46_Click()
    frmNews.Visible = True
End Sub

Private Sub Command68_Click()
    Dim TempSpeed As Long

    TempSpeed = Val(GameTimeSpeed.Text)

    If TempSpeed < 0 Or TempSpeed > 59 Then
        Call MsgBox("Por favor introduce un valor positivo menor de 60.")
        Exit Sub
    End If

    Gamespeed = TempSpeed

    SendGameClockToAll
End Sub

Private Sub Command69_Click()
    If Not TimeDisable Then
        Gamespeed = 0
        GameTimeSpeed.Text = 0
        TimeDisable = True
        Timer1.Enabled = False
        frmServer.Command69.Caption = "Activar Tiempo"
    Else
        Gamespeed = 1
        GameTimeSpeed.Text = 1
        TimeDisable = False
        Timer1.Enabled = True
        frmServer.Command69.Caption = "Desactivar Tiempo"
    End If

    Call DisabledTime

    If Not TimeDisable Then
        SendGameClockToAll
    End If
End Sub

Private Sub Command70_Click()
    ServerScript.Text = "Sub Server()" & vbNewLine & vbNewLine & "End Sub"
    Script.Visible = True
End Sub

Private Sub Command71_Click()
    Script.Visible = False
End Sub

Private Sub Command72_Click()
    Dim FileID As Integer
    Dim i As Long

    If scripting = 1 Then
        FileID = FreeFile

        Do
            If FileExists("\Scripts\Server" & i & ".txt") Then
                i = i + 1
            Else
                Open App.Path & "\Scripts\Server" & i & ".txt" For Output As #FileID
                    Print #FileID, ServerScript.Text
                Close #FileID
                
                Exit Do
            End If
        Loop
    
        MyScript.ReadInCode App.Path & "\Scripts\Server" & i & ".txt", "Scripts\Server" & i & ".txt", MyScript.SControl
        MyScript.ExecuteStatement "Scripts\Server" & i & ".txt", "Server "
    Else
        Call MsgBox("Los scripts estan desactivados. Activalos para realizar esta acción.")
    End If

    Script.Visible = False
End Sub

Private Sub Command73_Click()
    Dim i As Integer

    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            Call SendNewsTo(i)
        End If
    Next i
    
    MsgBox "Noticia enviada con exito", vbInformation
End Sub

Private Sub editorclases_Click()
Dim i As Long

For i = 0 To MAX_CLASSES
    If FileExists("\Clases\Class" & i & ".ini") Then
        frmClases.CP.AddItem "Class" & i
    End If
Next
frmClases.Visible = True
End Sub

Private Sub editorconfig_Click()
If frmConfiguracion.customnpc.Text = 1 Then
frmConfiguracion.Label44.Visible = True
frmConfiguracion.Label45.Visible = True
frmConfiguracion.Label46.Visible = True
frmConfiguracion.Text1.Visible = True
frmConfiguracion.Text2.Visible = True
frmConfiguracion.Text3.Visible = True
frmConfiguracion.Text1.Text = MAX_HEAD
frmConfiguracion.Text2.Text = MAX_BODY
frmConfiguracion.Text3.Text = MAX_LEGS
frmConfiguracion.nivelminimopk.Top = 5520
frmConfiguracion.Label24.Top = 5520
frmConfiguracion.Label25.Top = 6000
frmConfiguracion.mostrarnivel.Top = 6000
frmConfiguracion.Label26.Top = 6480
frmConfiguracion.clases.Top = 6480
Else
frmConfiguracion.Label44.Visible = False
frmConfiguracion.Label45.Visible = False
frmConfiguracion.Label46.Visible = False
frmConfiguracion.Text1.Visible = False
frmConfiguracion.Text2.Visible = False
frmConfiguracion.Text3.Visible = False
frmConfiguracion.Label24.Top = 3960
frmConfiguracion.Label25.Top = 4440
frmConfiguracion.Label26.Top = 4920
frmConfiguracion.nivelminimopk.Top = 3960
frmConfiguracion.mostrarnivel.Top = 4440
frmConfiguracion.clases.Top = 4920
End If
frmConfiguracion.Visible = True
End Sub

Private Sub envanuncio_Click()
frmAnuncio.Visible = True
End Sub

Private Sub Form_Load()
    Hours = Rand(1, 24)
    Minutes = Rand(0, 59)
    Seconds = Rand(0, 59)

    Gamespeed = 1

    lblVer.Caption = "Build: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        lvUsers.GridLines = True
    Else
        lvUsers.GridLines = False
    End If
End Sub

Private Sub Command1_Click()
    If Not tmrShutdown.Enabled Then
        tmrShutdown.Enabled = True
    End If
    
    Command1.Enabled = False
    MsgBox "El servidor se apagara en 30 segundos...", vbOKOnly
End Sub

Private Sub Command12_Click()
    Dim index As Long

    For index = 1 To MAX_PLAYERS
        If IsPlaying(index) Then
            If GetPlayerHP(index) < GetPlayerMaxHP(index) Then
                Call SetPlayerHP(index, GetPlayerMaxHP(index))
                Call SendHP(index)
            End If
        End If
    Next index

    Call GlobalMsg("El servidor ha curado a todos.", BRIGHTGREEN)
End Sub

Private Sub cmdKickPlayerReason_Click()
    If picKick.Visible Then
        picKick.Visible = False
    Else
        picKick.Visible = True
    End If
End Sub

Private Sub cmdBanPlayerReason_Click()
    If picBan.Visible Then
        picBan.Visible = False
    Else
        picBan.Visible = True
    End If
End Sub

Private Sub cmdJailPlayer_Click()
    If picJail.Visible Then
        picJail.Visible = False
    Else
        scrlJailMap.Max = MAX_MAPS
        scrlJailX.Max = MAX_MAPX
        scrlJailY.Max = MAX_MAPY

        picJail.Visible = True
    End If
End Sub

Private Sub cmdViewInfo_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If IsPlaying(index) Then
        CharInfo(0).Caption = "Cuenta: " & GetPlayerLogin(index)
        CharInfo(1).Caption = "Personaje: " & GetPlayerName(index)
        CharInfo(2).Caption = "Nivel: " & GetPlayerLevel(index)
        CharInfo(3).Caption = "PV: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index)
        CharInfo(4).Caption = "PM: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index)
        CharInfo(5).Caption = "PS: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index)
        CharInfo(6).Caption = "EXP: " & GetPlayerExp(index) & "/" & GetPlayerNextLevel(index)
        CharInfo(7).Caption = "Privilegio: " & GetPlayerAccess(index)
        CharInfo(8).Caption = "PK: " & GetPlayerPK(index)
        CharInfo(9).Caption = "Clase: " & ClassData(GetPlayerClass(index)).Name
        CharInfo(10).Caption = "Sprite: " & GetPlayerSprite(index)
        CharInfo(11).Caption = "Sexo: " & CStr(Player(index).Char(Player(index).CharNum).Sex)
        CharInfo(12).Caption = "Mapa: " & GetPlayerMap(index)
        CharInfo(13).Caption = "Clan: " & GetPlayerGuild(index)
        CharInfo(14).Caption = "Privilegio en clan: " & GetPlayerGuildAccess(index)
        CharInfo(15).Caption = "FRZ: " & GetPlayerSTR(index)
        CharInfo(16).Caption = "DEF: " & GetPlayerDEF(index)
        CharInfo(17).Caption = "Velocidad: " & GetPlayerSPEED(index)
        CharInfo(18).Caption = "Magia: " & GetPlayerMAGI(index)
        CharInfo(19).Caption = "Puntos: " & GetPlayerPOINTS(index)
        CharInfo(20).Caption = "Indice: " & index
        picStats.Visible = True
    End If
End Sub

Private Sub cmdMsgPlayer_Click()
    If picMessage.Visible Then
        picMessage.Visible = False
    Else
        picMessage.Visible = True
    End If
End Sub

Private Sub cmdMutePlayer_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If IsPlaying(index) Then
        Call PlayerMsg(index, "Has sido silenciado!", WHITE)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " ha sido silenciado!", True)
        Player(index).Mute = True
    End If
End Sub

Private Sub cmdUnmutePlayer_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If IsPlaying(index) Then
        Call PlayerMsg(index, "Has sido desilenciado!", WHITE)
        Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " ha sido desilenciado!", True)
        Player(index).Mute = False
    End If
End Sub

Private Sub cmdKillPlayer_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If IsPlaying(index) Then
        Call SetPlayerHP(index, 0)

        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
        Else
            If Map(GetPlayerMap(index)).BootMap > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).BootMap, Map(GetPlayerMap(index)).BootX, Map(GetPlayerMap(index)).BootY)
            Else
                Call PlayerWarp(index, START_MAP, START_X, START_Y)
            End If
        End If

        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))

        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)

        Call PlayerMsg(index, "Has sido asesinado por el servidor.", BRIGHTRED)
    End If
End Sub

Private Sub Command25_Click()
    If scripting = 1 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing

        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands

        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

        MyScript.ExecuteStatement "Scripts\Main.txt", "OnScriptReload"

        Call TextAdd(frmServer.txtText(0), "* Scripts Recargados *", True)
        Call AdminMsg("Los scripts han sido recargados.", 15)
        MsgBox "Scripts recargados con exito :)", vbInformation
    End If
End Sub

Private Sub Command26_Click()
    If scripting = 0 Then
        ' Check for Main.txt
        If Not FileExists("\Scripts\Main.txt") Then
            Call MsgBox("El archivo 'Scripts\Main.txt' no ha sido encontrado!", vbExclamation)
            Exit Sub
        End If

        scripting = 1

        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Scripting", 1

        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands

        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

        MyScript.ExecuteStatement "Scripts\Main.txt", "OnScriptReload"
        
        lblScriptOn.Caption = "Scripts: ON"
    End If
End Sub

Private Sub Command27_Click()
    If scripting = 1 Then
        scripting = 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Scripting", 0

        Set MyScript = Nothing
        Set clsScriptCommands = Nothing

        lblScriptOn.Caption = "Scripts: OFF"
    End If
End Sub

Private Sub Command28_Click()
    If FileExists("Editor.exe") Then
        Call Shell(App.Path & "\Editor.exe Scripts\Main.txt", vbNormalNoFocus)
    Else
        Call MsgBox("El editor de AE no ha sido encontrado!", vbOKOnly, "Error")
    End If
End Sub

Private Sub Command29_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "Todas las clases han sido recargadas.", True)
End Sub

Private Sub cmdHealPlayer_Click()
    Dim index As Long

    index = Val(lvUsers.ListItems(lvUsers.SelectedItem.index).Text)

    If IsPlaying(index) Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SendHP(index)

        Call PlayerMsg(index, "Has sido curado por el servidor.", BRIGHTGREEN)
    End If
End Sub

Private Sub Command30_Click()
    If FileExists("Editor.exe") Then
        Call Shell(App.Path & "\Editor.exe Classes\Info.ini", vbNormalNoFocus)
    Else
        Call MsgBox("El editor de AE no ha sido encontrado!", vbOKOnly, "Error")
    End If
End Sub

Private Sub Command31_Click()
    Dim index As Long

    For index = 1 To MAX_PLAYERS
        If IsPlaying(index) = True Then
            If GetPlayerAccess(index) <= 0 Then
                Call SetPlayerHP(index, 0)
                Call PlayerMsg(index, "Has sido asesinado por el servidor!", BRIGHTRED)

                ' Warp player away
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
                Else
                    If Map(GetPlayerMap(index)).BootMap > 0 Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).BootMap, Map(GetPlayerMap(index)).BootX, Map(GetPlayerMap(index)).BootY)
                    Else
                        Call PlayerWarp(index, START_MAP, START_X, START_Y)
                    End If
                End If

                Call SetPlayerHP(index, GetPlayerMaxHP(index))
                Call SetPlayerMP(index, GetPlayerMaxMP(index))
                Call SetPlayerSP(index, GetPlayerMaxSP(index))

                Call SendHP(index)
                Call SendMP(index)
                Call SendSP(index)
            End If
        End If
    Next index
End Sub

Private Sub Command32_Click()
    scrlMM.Max = MAX_MAPS
    scrlMX.Max = MAX_MAPX
    scrlMY.Max = MAX_MAPY
    picWarpAll.Visible = True
End Sub

Private Sub Command33_Click()
    picExp.Visible = True
End Sub

Private Sub Command34_Click()
    Dim index As Long
    Dim i As Long

    For index = 1 To MAX_PLAYERS
        If IsPlaying(index) Then
            If GetPlayerLevel(index) >= MAX_LEVEL Then
                Call SetPlayerExp(index, Experience(MAX_LEVEL))
            Else
                Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)

                i = Int(GetPlayerSPEED(index) / 10)

                If i < 1 Then i = 1
                If i > 3 Then i = 3

                Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)

                If GetPlayerLevel(index) >= MAX_LEVEL Then
                    Call SetPlayerExp(index, Experience(MAX_LEVEL))
                End If
            End If

            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Call SendPTS(index)
        End If
    Next index

    Call GlobalMsg("El servidor ha dado 1 nivel mas a todos!", BRIGHTGREEN)
End Sub

Private Sub Command35_Click()
    Dim i As Long

    MapList.Clear

    For i = 1 To MAX_MAPS
        MapList.AddItem i & ": " & Map(i).Name
    Next i

    frmServer.MapList.Selected(0) = True
End Sub

Private Sub Command36_Click()
    Dim mapnum As Long
    Dim i As Long

    mapnum = MapList.ListIndex + 1

    MapInfo(0).Caption = "Mapa " & mapnum & " - " & Map(mapnum).Name
    MapInfo(1).Caption = "Revisión: " & Map(mapnum).Revision
    MapInfo(2).Caption = "Moral: " & Map(mapnum).Moral
    MapInfo(3).Caption = "Arriba: " & Map(mapnum).Up
    MapInfo(4).Caption = "Abajo: " & Map(mapnum).Down
    MapInfo(5).Caption = "Izquierda: " & Map(mapnum).Left
    MapInfo(6).Caption = "Derecha: " & Map(mapnum).Right
    MapInfo(7).Caption = "Musica: " & Map(mapnum).music
    MapInfo(8).Caption = "BootMap: " & Map(mapnum).BootMap
    MapInfo(9).Caption = "BootX: " & Map(mapnum).BootX
    MapInfo(10).Caption = "BootY: " & Map(mapnum).BootY
    MapInfo(11).Caption = "Tiendas: " & Map(mapnum).Shop
    MapInfo(12).Caption = "Interior: " & Map(mapnum).Indoors

    lstNPC.Clear

    For i = 1 To MAX_MAP_NPCS
        lstNPC.AddItem i & ": " & NPC(Map(mapnum).NPC(i)).Name
    Next i

    picMap.Visible = True
End Sub

' Sistema de conexion Winsock
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' Fin de winsock
Private Sub Command37_Click()
    Dim index As Long
    Dim mapnum As Long
    Dim MapX As Long
    Dim MapY As Long

    mapnum = Int(scrlMM.Value)
    MapX = Int(scrlMX.Value)
    MapY = Int(scrlMY.Value)

    For index = 1 To MAX_PLAYERS
        If IsPlaying(index) Then
            If GetPlayerAccess(index) = 0 Then
                Call PlayerWarp(index, mapnum, MapX, MapY)
            End If
        End If
    Next index

    Call GlobalMsg("El servidor ha movido a todos los jugadores al mapa" & mapnum & ".", YELLOW)

    picWarpAll.Visible = False
End Sub

Private Sub Command38_Click()
    picWarpAll.Visible = False
End Sub

Private Sub Command39_Click()
    picExp.Visible = False
End Sub

Private Sub Command40_Click()
    Dim index As Long
    Dim TotalExp As Long

    TotalExp = CLng(scrlExp.Value)

    If TotalExp > 0 Then
        For index = 1 To MAX_PLAYERS
            If IsPlaying(index) Then
                Call SetPlayerExp(index, GetPlayerExp(index) + TotalExp)
                Call CheckPlayerLevelUp(index)
            End If
        Next index

        Call GlobalMsg("El servidor ha dado a todos " & TotalExp & " de experiencia!", BRIGHTGREEN)
    End If

    picExp.Visible = False
End Sub

Private Sub Command41_Click()
    picMap.Visible = False
End Sub

Private Sub cmdWarpPlayer_Click()
    If picWarp.Visible Then
        picWarp.Visible = False
    Else
        scrlWarpMap.Max = MAX_MAPS
        scrlWarpX.Max = MAX_MAPX
        scrlWarpY.Max = MAX_MAPY

        picWarp.Visible = True
    End If
End Sub

Private Sub Command5_Click()
    picCMsg.Visible = False
End Sub

Private Sub Command59_Click()
    picWeather.Visible = True
End Sub

Private Sub cmdSaveLogs_Click()
    Call SaveLogs
End Sub

Private Sub Command61_Click()
    picWeather.Visible = False
End Sub

Private Sub Command62_Click()
    WeatherType = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Private Sub Command63_Click()
    WeatherType = WEATHER_THUNDER
    Call SendWeatherToAll
End Sub

Private Sub Command64_Click()
    WeatherType = WEATHER_RAINING
    Call SendWeatherToAll
End Sub

Private Sub Command65_Click()
    WeatherType = WEATHER_SNOWING
    Call SendWeatherToAll
End Sub

Private Sub Command66_Click()
    Dim i As Long

    lvUsers.ListItems.Clear

    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
End Sub

Private Sub Command8_Click()
    picStats.Visible = False
End Sub

Private Sub Command9_Click()
    Dim index As Long

    For index = 1 To MAX_PLAYERS
        If IsPlaying(index) Then
            If GetPlayerAccess(index) = 0 Then
                Call GlobalMsg(GetPlayerName(index) & " ha sido expulsado por el servidor!", WHITE)
                Call AlertMsg(index, "Has sido expulsado por el servidor!")
            End If
        End If
    Next index
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide

        With nid
            .cbSize = Len(nid)
            .hWnd = Me.hWnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon
            .szTip = Chr$(0)
            .uTimeout = 3000
            .dwState = NIS_SHAREDICON
            .dwInfoFlags = vbInformation
        End With
        
        Call Shell_NotifyIcon(NIM_ADD, nid)
    Else
        Call Shell_NotifyIcon(NIM_DELETE, nid)
    End If
End Sub

Private Sub Form_Terminate()
    Call SaveAllPlayersOnline
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveAllPlayersOnline
    Call DestroyServer
End Sub
Private Sub lstTopics_Click()
    Dim filename As String
    Dim hfile As Long

    txtTopic.Text = vbNullString

    filename = lstTopics.ListIndex + 1 & ".txt"

    If FileExists("Guides\" & filename) = True Then
        hfile = FreeFile

        Open App.Path & "\Guides\" & filename For Input As #hfile
            txtTopic.Text = Input$(LOF(hfile), hfile)
        Close #hfile
    End If
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Value = Checked Then
        ServerLog = False
    Else
        ServerLog = True
    End If
End Sub
Private Sub PlayerTimer_Timer()
    If PlayerI <= MAX_PLAYERS Then
        If IsPlaying(PlayerI) Then
            Call SavePlayer(PlayerI)
        End If

        PlayerI = PlayerI + 1
    End If

    If PlayerI >= MAX_PLAYERS Then
        PlayerI = 1
        PlayerTimer.Enabled = False
        tmrPlayerSave.Enabled = True
    End If
End Sub

Private Sub scrlExp_Change()
    lblMassExp.Caption = "Experiencia: " & scrlExp.Value
End Sub

Private Sub scrlJailMap_Change()
    lblJailMap.Caption = "Mapa: " & scrlJailMap.Value
End Sub

Private Sub scrlMM_Change()
    lblMM.Caption = "Mapa: " & scrlMM.Value
End Sub

Private Sub scrlMX_Change()
    lblMX.Caption = "X: " & scrlMX.Value
End Sub

Private Sub scrlMY_Change()
    lblMY.Caption = "Y: " & scrlMY.Value
End Sub

Private Sub scrlRainIntensity_Change()
    lblRainIntensity.Caption = "Intensity: " & scrlRainIntensity.Value
    WeatherLevel = scrlRainIntensity.Value

    Call SendWeatherToAll
End Sub

Private Sub scrlJailX_Change()
    lblJailX.Caption = "X: " & scrlJailX.Value
End Sub

Private Sub scrlJailY_Change()
    lblJailY.Caption = "Y: " & scrlJailY.Value
End Sub
Private Sub Timer1_Timer()
    Dim AMorPM As String
    Dim TempSeconds As Integer
    Dim PrintSeconds As String
    Dim PrintSeconds2 As String
    Dim PrintMinutes As String
    Dim PrintMinutes2 As String
    Dim PrintHours As Integer

    Seconds = Seconds + Gamespeed

    If Seconds > 59 Then
        Minutes = Minutes + 1
        Seconds = Seconds - 60
    End If

    If Minutes > 59 Then
        Hours = Hours + 1
        Minutes = 0
    End If
    If Hours > 24 Then
        Hours = 1
    End If

    If Hours > 12 Then
        AMorPM = "PM"
        PrintHours = Hours - 12
    Else
        AMorPM = "AM"
        PrintHours = Hours
    End If

    If Hours = 24 Then
        AMorPM = "AM"
    End If

    TempSeconds = Seconds

    If Seconds > 9 Then
        PrintSeconds = TempSeconds
    Else
        PrintSeconds = "0" & Seconds
    End If

    If Seconds > 50 Then
        PrintSeconds2 = "0" & 60 - TempSeconds
    Else
        PrintSeconds2 = 60 - TempSeconds
    End If

    If Minutes > 9 Then
        PrintMinutes = Minutes
    Else
        PrintMinutes = "0" & Minutes
    End If

    If Minutes > 50 Then
        PrintMinutes2 = "0" & 60 - Minutes
    Else
        PrintMinutes2 = 60 - Minutes
    End If

    Label8.Caption = "Tiempo: " & PrintHours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM

    If Hours > 20 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
            Call SendTimeToAll
        End If
    ElseIf Hours < 21 Then
        If Hours > 6 Then
            If GameTime = TIME_NIGHT Then
                GameTime = TIME_DAY
                Call SendTimeToAll
            End If
        End If
    ElseIf Hours < 7 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
            Call SendTimeToAll
        End If
    End If

    If Hours > 11 Then
        GameClock = Hours - 12 & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
    Else
        GameClock = Hours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
    End If

    ' Sync game clock every 10 minutes
    If Minutes Mod 10 = 0 Then
        Call SendGameClockToAll
    End If

    If scripting = 1 Then
       ' MyScript.ExecuteStatement "Scripts\Main.txt", "TimedEvent " & Hours & "," & Minutes & "," & Seconds
    End If
End Sub

Private Sub tmrChatLogs_Timer()
    If frmServer.chkChat.Value = Unchecked Then
        CHATLOG_TIMER = 3600
        lblLogTime.Caption = "Guardar registro de chat desactivado!"
        Exit Sub
    End If

    If CHATLOG_TIMER < 1 Then
        CHATLOG_TIMER = 3600
    End If

    If CHATLOG_TIMER > 60 Then
        lblLogTime.Caption = "El registro del chat se guardara en " & Int(CHATLOG_TIMER / 60) & " Minuto(s)"
    Else
        lblLogTime.Caption = "El registro del chat se guardara en " & Int(CHATLOG_TIMER) & " Segundo(s)"
    End If

    CHATLOG_TIMER = CHATLOG_TIMER - 1

    If CHATLOG_TIMER <= 0 Then
        Call TextAdd(txtText(0), "Los registros del chat han sido guardados!", True)
        Call SaveLogs
    End If
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrScriptedTimer_Timer()
    Call ScriptedTimer
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) <> 0 Then
            Call GlobalMsg(txtChat.Text, WHITE)
            Call TextAdd(frmServer.txtText(0), "Servidor: " & txtChat.Text, True)
            txtChat.Text = vbNullString
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub tmrShutdown_Timer()
    If SHUTDOWN_TIMER < 1 Then
        SHUTDOWN_TIMER = 30
    End If

    If SHUTDOWN_TIMER Mod 5 = 0 Or SHUTDOWN_TIMER <= 10 Then
        Call GlobalMsg("El servidor se apagara en " & SHUTDOWN_TIMER & " segundo(s).", BRIGHTBLUE)
        Call TextAdd(frmServer.txtText(0), "El servidor se apagara en " & SHUTDOWN_TIMER & " segundo(s).", True)
    End If
    
    SHUTDOWN_TIMER = SHUTDOWN_TIMER - 1
    
    If SHUTDOWN_TIMER < 1 Then
        Call GlobalMsg("El servidor ha sido apagado.", BRIGHTRED)
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub txtText_GotFocus(index As Integer)
    txtChat.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
