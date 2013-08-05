VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Hechizos"
   ClientHeight    =   5910
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   6375
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpellEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   397
      TabMaxWidth     =   3545
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Información"
      TabPicture(0)   =   "frmSpellEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVitalMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRange"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblElement"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "info"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlSound"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlVitalMod"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "scrlRange"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "scrlElement"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkArea"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Animación"
      TabPicture(1)   =   "frmSpellEditor.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "scrlSpellAnim"
      Tab(1).Control(1)=   "scrlSpellTime"
      Tab(1).Control(2)=   "scrlSpellDone"
      Tab(1).Control(3)=   "picSpell"
      Tab(1).Control(4)=   "chkBig"
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(6)=   "Picture1"
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(8)=   "lblSpellAnim"
      Tab(1).Control(9)=   "lblSpellTime"
      Tab(1).Control(10)=   "lblSpellDone"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Requerimientos"
      TabPicture(2)   =   "frmSpellEditor.frx":0FFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.HScrollBar scrlSpellAnim 
         Height          =   270
         Left            =   -74760
         Max             =   2000
         TabIndex        =   39
         Top             =   720
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSpellTime 
         Height          =   270
         Left            =   -74760
         Max             =   500
         Min             =   40
         TabIndex        =   38
         Top             =   3120
         Value           =   40
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSpellDone 
         Height          =   270
         Left            =   -74760
         Max             =   10
         Min             =   1
         TabIndex        =   37
         Top             =   3840
         Value           =   1
         Width           =   5655
      End
      Begin VB.PictureBox picSpell 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -70120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   35
         Top             =   1275
         Width           =   480
      End
      Begin VB.CheckBox chkBig 
         Caption         =   "Hechizo Grande"
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
         Left            =   -74760
         TabIndex        =   34
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Icono"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   27
         Top             =   1440
         Width           =   3735
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   840
            Max             =   100
            TabIndex        =   31
            Top             =   600
            Width           =   2775
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   28
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
               TabIndex        =   29
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   30
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ID Hechizo:"
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
            Left            =   840
            TabIndex        =   33
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label13 
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
            Left            =   1680
            TabIndex        =   32
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Efecto de Area"
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
         Left            =   1320
         TabIndex        =   26
         Top             =   3550
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Requerimientos del Hechizo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74760
         TabIndex        =   16
         Top             =   600
         Width           =   5175
         Begin VB.HScrollBar CTHScroll 
            Height          =   270
            Left            =   120
            Max             =   10000
            TabIndex        =   46
            Top             =   2880
            Width           =   4935
         End
         Begin VB.HScrollBar TTCHScroll 
            Height          =   270
            Left            =   120
            Max             =   1200
            TabIndex        =   45
            Top             =   2280
            Width           =   4935
         End
         Begin VB.ComboBox cmbClassReq 
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
            ItemData        =   "frmSpellEditor.frx":1016
            Left            =   120
            List            =   "frmSpellEditor.frx":1018
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   480
            Width           =   4905
         End
         Begin VB.HScrollBar scrlCost 
            Height          =   270
            Left            =   120
            Max             =   1000
            TabIndex        =   18
            Top             =   1680
            Width           =   4935
         End
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   270
            Left            =   120
            Max             =   500
            TabIndex        =   17
            Top             =   1080
            Value           =   1
            Width           =   4935
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "En Segundos"
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
            Left            =   4200
            TabIndex        =   50
            Top             =   2040
            Width           =   810
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "En Milisegundos"
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
            Left            =   4080
            TabIndex        =   49
            Top             =   2640
            Width           =   990
         End
         Begin VB.Label Label14 
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
            Height          =   165
            Left            =   1320
            TabIndex        =   48
            Top             =   2640
            Width           =   75
         End
         Begin VB.Label CTlbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo a Castear:"
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
            Left            =   75
            TabIndex        =   47
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label Label12 
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
            Height          =   165
            Left            =   1080
            TabIndex        =   44
            Top             =   2040
            Width           =   75
         End
         Begin VB.Label TTCLbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CoolDown:"
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
            Left            =   75
            TabIndex        =   43
            Top             =   2040
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase Requerida"
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
            TabIndex        =   24
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblCost 
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
            Height          =   165
            Left            =   1080
            TabIndex        =   22
            Top             =   1440
            Width           =   75
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Solo Admins"
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
            Left            =   1320
            TabIndex        =   21
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Coste de PM:"
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
            TabIndex        =   20
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Requerido:"
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
            TabIndex        =   19
            Top             =   840
            Width           =   1035
         End
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   13
         Top             =   4560
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   270
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   12
         Top             =   3840
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   270
         Left            =   240
         Max             =   1000
         TabIndex        =   5
         Top             =   2400
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSound 
         Height          =   270
         Left            =   240
         Max             =   100
         TabIndex        =   4
         Top             =   3120
         Width           =   5655
      End
      Begin VB.Frame info 
         Caption         =   "Información"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5655
         Begin VB.ComboBox cmbType 
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
            ItemData        =   "frmSpellEditor.frx":101A
            Left            =   120
            List            =   "frmSpellEditor.frx":1033
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1080
            Width           =   5355
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
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   5355
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Hechizo"
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
            TabIndex        =   25
            Top             =   840
            Width           =   1320
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Hechizo"
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
            TabIndex        =   3
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   -70680
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   36
         Top             =   720
         Width           =   1605
      End
      Begin Eclipse.jcbutton Command1 
         Height          =   495
         Left            =   -70680
         TabIndex        =   54
         Top             =   2400
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
         Caption         =   "Refrescar"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblSpellAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Animación: 0"
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   480
         Width           =   810
      End
      Begin VB.Label lblSpellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo: 40"
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
         Left            =   -74760
         TabIndex        =   41
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label lblSpellDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclo de Animación 1 vez"
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
         Left            =   -74760
         TabIndex        =   40
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label9 
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
         Left            =   180
         TabIndex        =   15
         Top             =   4320
         Width           =   630
      End
      Begin VB.Label lblElement 
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
         Left            =   960
         TabIndex        =   14
         Top             =   4320
         Width           =   510
      End
      Begin VB.Label lblRange 
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
         Height          =   165
         Left            =   720
         TabIndex        =   11
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label lblVitalMod 
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
         Height          =   165
         Left            =   1440
         TabIndex        =   9
         Top             =   2160
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modificación Vital:"
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
         TabIndex        =   8
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido:"
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
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sin sonido"
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
         Left            =   840
         TabIndex        =   6
         Top             =   2880
         Width           =   630
      End
   End
   Begin Eclipse.jcbutton cmdOk 
      Height          =   495
      Left            =   960
      TabIndex        =   52
      Top             =   5280
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
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Eclipse.jcbutton cmdCancel 
      Height          =   495
      Left            =   3240
      TabIndex        =   53
      Top             =   5280
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
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Done As Long
Private time As Long
Private SpellVar As Long

Private Sub chkBig_Click()
    frmSpellEditor.ScaleMode = 3
    Done = 0
    SpellVar = 0
    picSpell.Refresh
    If chkBig.value = Checked Then
        picSpell.Width = 1440
        picSpell.Height = 1440
        picSpell.top = 800
        picSpell.Left = 4400
    Else
        picSpell.Width = 480
        picSpell.Height = 480
        picSpell.top = 1275
        picSpell.Left = 4880
    End If
End Sub

Private Sub Command1_Click()
    Done = 0
End Sub

Private Sub CTHScroll_Change()
Label14.Caption = CTHScroll.value
End Sub

Private Sub Form_Load()
    scrlElement.max = MAX_ELEMENTS
End Sub

Private Sub HScroll1_Change()
    Label13.Caption = STR(HScroll1.value)
    frmSpellEditor.iconn.top = (HScroll1.value * 32) * -1
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = STR(scrlCost.value)
End Sub

Private Sub scrlElement_Change()
    lblElement.Caption = Element(scrlElement.value).name
End Sub

Private Sub scrlLevelReq_Change()
    If STR(scrlLevelReq.value) = 0 Then
        lblLevelReq.Caption = "Solo Staff"
    Else
        lblLevelReq.Caption = STR(scrlLevelReq.value)
    End If
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.value)
End Sub

Private Sub scrlSound_Change()
    If STR(scrlSound.value) = 0 Then
        lblSound.Caption = "Sin Sonido"
    Else
        lblSound.Caption = STR(scrlSound.value)
        Call PlaySound("magic" & scrlSound.value & ".wav")
    End If
End Sub

Private Sub scrlSpellAnim_Change()
    lblSpellAnim.Caption = "Animación: " & scrlSpellAnim.value
    Done = 0
End Sub

Private Sub scrlSpellDone_Change()
    Dim String2 As String
    String2 = "Veces"
    If scrlSpellDone.value = 1 Then
        String2 = "Vez"
    End If
    lblSpellDone.Caption = "Ciclo de animación " & scrlSpellDone.value & " " & String2
    Done = 0
End Sub

Private Sub scrlSpellTime_Change()
    lblSpellTime.Caption = "Tiempo: " & scrlSpellTime.value
    Done = 0
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.value)
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub Timer1_Timer()
    Dim sRECT As RECT
    Dim dRECT As RECT
    Dim SpellDone As Long
    Dim SpellAnim As Long
    Dim SpellTime As Long

    SpellDone = scrlSpellDone.value
    SpellAnim = scrlSpellAnim.value
    SpellTime = scrlSpellTime.value

    If chkBig.value = Checked Then
        SpellAnim = scrlSpellAnim.value * 3
    End If

    If SpellAnim <= 0 Then
        Exit Sub
    End If
    If Done = SpellDone Then
        Exit Sub
    End If
    If chkBig = Checked Then
        With dRECT
            .top = 0
            .Bottom = PIC_Y + 64
            .Left = 0
            .Right = PIC_X + 64
        End With
    Else
        With dRECT
            .top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    If chkBig.value = Checked Then
        If SpellVar > 32 Then
            Done = Done + 1
            SpellVar = 0
        End If
        If GetTickCount > time + SpellTime Then
            time = GetTickCount
            SpellVar = SpellVar + 3
        End If
    Else
        If SpellVar > 10 Then
            Done = Done + 1
            SpellVar = 0
        End If
        If GetTickCount > time + SpellTime Then
            time = GetTickCount
            SpellVar = SpellVar + 1
        End If
    End If
    If chkBig = Checked Then
        If DD_BigSpellAnim Is Nothing Then
        Else
            With sRECT
                .top = SpellAnim * PIC_Y
                .Bottom = .top + (PIC_Y * 3)
                .Left = SpellVar * PIC_X
                .Right = .Left + (PIC_X * 3)
            End With

            Call DD_BigSpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    Else
        If DD_SpellAnim Is Nothing Then
        Else
            With sRECT
                .top = SpellAnim * PIC_Y
                .Bottom = .top + PIC_Y
                .Left = SpellVar * PIC_X
                .Right = .Left + PIC_X
            End With

            Call DD_SpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    End If
End Sub
Private Sub cmbType_Click()
    If (cmbType.ListIndex = SPELL_TYPE_SCRIPTED) Then
        Label4.Caption = "Script"
        Label8.Visible = False
        lblSound.Visible = False
        scrlSound.Visible = False
        Label2.Visible = False
        lblRange.Visible = False
        scrlRange.Visible = False
        lblSpellAnim.Visible = False
        scrlSpellAnim.Visible = False
        lblSpellTime.Visible = False
        scrlSpellTime.Visible = False
        lblSpellDone.Visible = False
        scrlSpellDone.Visible = False
        chkArea.Visible = False
        Command1.Visible = False
        picSpell.Visible = False

    Else
        Label4.Caption = "Modificación Vital"
        Label8.Visible = True
        lblSound.Visible = True
        scrlSound.Visible = True
        Label2.Visible = True
        lblRange.Visible = True
        scrlRange.Visible = True
        lblSpellAnim.Visible = True
        scrlSpellAnim.Visible = True
        lblSpellTime.Visible = True
        scrlSpellTime.Visible = True
        lblSpellDone.Visible = True
        scrlSpellDone.Visible = True
        chkArea.Visible = True
        Command1.Visible = True
        picSpell.Visible = True
    End If
End Sub

Private Sub TTCHScroll_Change()
Label12.Caption = TTCHScroll.value
End Sub

