VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades de Mapa"
   ClientHeight    =   6600
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   9675
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
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11298
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   344
      TabMaxWidth     =   1773
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmMapProperties.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMapName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOk"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtMapName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraSwitch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraSettings"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraRespawn"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraBGM"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "NPC"
      TabPicture(1)   =   "frmMapProperties.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCoordX"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblMonster"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblCoordY"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmbNpc(14)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbNpc(13)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbNpc(12)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmbNpc(11)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmbNpc(10)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmbNpc(9)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmbNpc(8)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmbNpc(7)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmbNpc(6)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmbNpc(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmbNpc(4)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmbNpc(3)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmbNpc(2)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmbNpc(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmbNpc(0)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "cmdClear"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdCopy(0)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmdCopy(1)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmdCopy(2)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmdCopy(3)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmdCopy(4)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "cmdCopy(5)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cmdCopy(6)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cmdCopy(7)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "cmdCopy(8)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmdCopy(10)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmdCopy(11)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "cmdCopy(12)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "cmdCopy(13)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "cmdCopy(9)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "cmbNpcY(0)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "cmbNpcX(1)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cmbNpcY(1)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "cmbNpcX(2)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "cmbNpcY(2)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "cmbNpcX(3)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "cmbNpcY(3)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "cmbNpcX(4)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "cmbNpcY(4)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "cmbNpcX(5)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "cmbNpcY(5)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "cmbNpcX(6)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "cmbNpcY(6)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "cmbNpcX(7)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "cmbNpcY(7)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "cmbNpcX(8)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "cmbNpcY(8)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "cmbNpcX(9)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "cmbNpcY(9)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "cmbNpcX(10)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "cmbNpcY(10)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "cmbNpcX(11)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "cmbNpcY(11)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "cmbNpcX(12)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "cmbNpcY(12)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "cmbNpcX(13)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "cmbNpcY(13)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "cmbNpcX(14)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "cmbNpcY(14)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "cmbNpcX(0)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "cmdSetRand"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "cmbNpcY(15)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "cmbNpcX(15)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "cmbNpc(15)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "cmbNpcY(16)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "cmbNpcX(16)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "cmbNpc(16)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "cmbNpcY(17)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "cmbNpcX(17)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "cmbNpc(17)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "cmbNpcY(18)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "cmbNpcX(18)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "cmbNpc(18)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "cmbNpcY(19)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "cmbNpcX(19)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "cmbNpc(19)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "cmbNpcY(20)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "cmbNpcX(20)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "cmbNpc(20)"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "cmbNpcY(21)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "cmbNpcX(21)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "cmbNpc(21)"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "cmbNpcY(22)"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "cmbNpcX(22)"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "cmbNpc(22)"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "cmbNpcY(23)"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "cmbNpcX(23)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "cmbNpc(23)"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "cmbNpcY(24)"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "cmbNpcX(24)"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "cmbNpc(24)"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "cmdCopy(14)"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "cmdCopy(15)"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "cmdCopy(16)"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "cmdCopy(17)"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "cmdCopy(18)"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "cmdCopy(19)"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "cmdCopy(20)"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "cmdCopy(21)"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "cmdCopy(22)"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).Control(106)=   "cmdCopy(23)"
      Tab(1).Control(106).Enabled=   0   'False
      Tab(1).ControlCount=   107
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   23
         Left            =   -66360
         TabIndex        =   135
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   22
         Left            =   -66360
         TabIndex        =   134
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   21
         Left            =   -66360
         TabIndex        =   133
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   20
         Left            =   -66360
         TabIndex        =   132
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   19
         Left            =   -66360
         TabIndex        =   131
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   18
         Left            =   -66360
         TabIndex        =   130
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   17
         Left            =   -66360
         TabIndex        =   129
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   16
         Left            =   -66360
         TabIndex        =   128
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   15
         Left            =   -66360
         TabIndex        =   127
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   14
         Left            =   -71160
         TabIndex        =   126
         Top             =   5760
         Width           =   615
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   3960
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   118
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   -70080
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdSetRand 
         Caption         =   "Resetear todas las coordenadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69600
         TabIndex        =   95
         Top             =   5280
         Width           =   3135
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   5760
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   5760
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   5400
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   5400
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   5040
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   5040
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   4680
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   4680
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -72960
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   9
         Left            =   -71160
         TabIndex        =   51
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   13
         Left            =   -71160
         TabIndex        =   50
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   12
         Left            =   -71160
         TabIndex        =   49
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   11
         Left            =   -71160
         TabIndex        =   48
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   10
         Left            =   -71160
         TabIndex        =   47
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   8
         Left            =   -71160
         TabIndex        =   46
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   7
         Left            =   -71160
         TabIndex        =   45
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   6
         Left            =   -71160
         TabIndex        =   44
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   5
         Left            =   -71160
         TabIndex        =   43
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   4
         Left            =   -71160
         TabIndex        =   42
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   3
         Left            =   -71160
         TabIndex        =   41
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   2
         Left            =   -71160
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   1
         Left            =   -71160
         TabIndex        =   39
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar"
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
         Index           =   0
         Left            =   -71160
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Vaciar mapa de NPCs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69600
         TabIndex        =   37
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Frame fraBGM 
         Caption         =   "Musica de Fondo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3375
         Left            =   3000
         TabIndex        =   36
         Top             =   2640
         Width           =   5175
         Begin VB.CheckBox chkURL 
            Caption         =   "Usar URL"
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
            TabIndex        =   61
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtURL 
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
            Left            =   240
            MaxLength       =   100
            TabIndex        =   60
            Top             =   2160
            Width           =   4695
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "Reproducir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   58
            Top             =   2880
            Width           =   2280
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Detener"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   57
            Top             =   2880
            Width           =   2280
         End
         Begin VB.ListBox lstMusic 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            ItemData        =   "frmMapProperties.frx":0FFA
            Left            =   240
            List            =   "frmMapProperties.frx":0FFC
            TabIndex        =   56
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label lblURL 
            Caption         =   "URL"
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
            TabIndex        =   59
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.Frame fraRespawn 
         Caption         =   "Respawneo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   1080
         TabIndex        =   29
         Top             =   3240
         Width           =   1815
         Begin VB.TextBox txtBootMap 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   32
            Text            =   "0"
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtBootX 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   31
            Text            =   "0"
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox txtBootY 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   30
            Text            =   "0"
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label lblMap 
            Caption         =   "Mapa"
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
            Left            =   240
            TabIndex        =   35
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label lblX 
            Caption         =   "X"
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
            TabIndex        =   34
            Top             =   660
            Width           =   135
         End
         Begin VB.Label lblY 
            Caption         =   "Y"
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
            TabIndex        =   33
            Top             =   1020
            Width           =   120
         End
      End
      Begin VB.Frame fraSettings 
         Caption         =   "Configuración"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1620
         Left            =   3000
         TabIndex        =   26
         Top             =   960
         Width           =   5205
         Begin VB.ComboBox cmbWeather 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMapProperties.frx":0FFE
            Left            =   240
            List            =   "frmMapProperties.frx":100E
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1080
            Width           =   4695
         End
         Begin VB.ComboBox cmbMoral 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMapProperties.frx":1033
            Left            =   240
            List            =   "frmMapProperties.frx":1043
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label lblWeather 
            Caption         =   "Climatologia"
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
            Left            =   240
            TabIndex        =   54
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label lblMorality 
            Caption         =   "Moralidad"
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
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.Frame fraSwitch 
         Caption         =   "Cambio de Mapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   1815
         Begin VB.TextBox txtUp 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   53
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkIndoors 
            Caption         =   "Interiores"
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
            Left            =   360
            TabIndex        =   52
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtLeft 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   24
            Text            =   "0"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtDown 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   22
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtRight 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   20
            Text            =   "0"
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblLeft 
            Caption         =   "Izq"
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
            TabIndex        =   25
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label lblDown 
            Caption         =   "Abajo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   435
         End
         Begin VB.Label lblRight 
            Caption         =   "Der"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblUp 
            Caption         =   "Arriba"
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
            TabIndex        =   19
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.TextBox txtMapName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   16
         Top             =   600
         Width           =   6225
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmMapProperties.frx":107F
         Left            =   -74880
         List            =   "frmMapProperties.frx":1081
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         ItemData        =   "frmMapProperties.frx":1083
         Left            =   -74880
         List            =   "frmMapProperties.frx":1085
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         ItemData        =   "frmMapProperties.frx":1087
         Left            =   -74880
         List            =   "frmMapProperties.frx":1089
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3960
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4320
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4680
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   5040
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         ItemData        =   "frmMapProperties.frx":108B
         Left            =   -74880
         List            =   "frmMapProperties.frx":108D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   5400
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5760
         Width           =   1815
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   495
         Left            =   960
         TabIndex        =   139
         Top             =   4920
         Width           =   1935
         _ExtentX        =   3413
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
         PictureNormal   =   "frmMapProperties.frx":108F
         PictureHot      =   "frmMapProperties.frx":1873
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cmdCancel 
         Height          =   495
         Left            =   960
         TabIndex        =   140
         Top             =   5520
         Width           =   1935
         _ExtentX        =   3413
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
         PictureNormal   =   "frmMapProperties.frx":2057
         PictureHot      =   "frmMapProperties.frx":29AB
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label Label3 
         Caption         =   "X Coord"
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
         Left            =   -68160
         TabIndex        =   138
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre del NPC"
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
         Left            =   -70080
         TabIndex        =   137
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Y Coord"
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
         Left            =   -67200
         TabIndex        =   136
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCoordY 
         Caption         =   "Y Coord"
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
         Left            =   -72000
         TabIndex        =   65
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMonster 
         Caption         =   "Nombre del NPC"
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
         Left            =   -74880
         TabIndex        =   63
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCoordX 
         Caption         =   "X Coord"
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
         TabIndex        =   62
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMapName 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Dim I As Integer

    For I = 1 To MAX_MAP_NPCS
        cmbNpc(I - 1).ListIndex = 0
        cmbNpcX(I - 1).ListIndex = 0
        cmbNpcY(I - 1).ListIndex = 0
    Next I
End Sub

Private Sub cmdPlay_Click()
    Call StopBGM
    If chkURL.value = 0 Then
        MapSound = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\Musica\" & lstMusic.List(lstMusic.ListIndex)), 0, 0, 0)
        Call BASS_ChannelPlay(MapSound, BASSFALSE)
    Else
        MapSound = BASS_StreamCreateURL(txtURL.Text, 0, BASS_SAMPLE_LOOP, 0, 0)
        Call BASS_ChannelPlay(MapSound, BASSFALSE)
    End If
End Sub
Private Sub cmdSetRand_Click()
    Dim X As Long
    
    For X = 1 To 25
        cmbNpcX(X - 1).ListIndex = 0
        cmbNpcY(X - 1).ListIndex = 0
    Next X
End Sub

Private Sub cmdStop_Click()
    Call StopBGM
End Sub

Private Sub cmdCopy_Click(index As Integer)
    cmbNpc(index + 1).ListIndex = cmbNpc(index).ListIndex
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim Y As Integer

    ListMusic (App.Path & "\Musica\")
    ListBGS (App.Path & "\BGS\")

    txtMapName.Text = Trim$(Map(GetPlayerMap(MyIndex)).name)
    txtUp.Text = STR$(Map(GetPlayerMap(MyIndex)).Up)
    txtDown.Text = STR$(Map(GetPlayerMap(MyIndex)).Down)
    txtLeft.Text = STR$(Map(GetPlayerMap(MyIndex)).Left)
    txtRight.Text = STR$(Map(GetPlayerMap(MyIndex)).Right)
    cmbMoral.ListIndex = Map(GetPlayerMap(MyIndex)).Moral
    txtBootMap.Text = STR$(Map(GetPlayerMap(MyIndex)).BootMap)
    txtBootX.Text = STR$(Map(GetPlayerMap(MyIndex)).BootX)
    txtBootY.Text = STR$(Map(GetPlayerMap(MyIndex)).BootY)
    lstMusic = Trim$(Map(GetPlayerMap(MyIndex)).music)
    lstMusic.Text = Trim$(Map(GetPlayerMap(MyIndex)).music)
    chkIndoors.value = STR$(Map(GetPlayerMap(MyIndex)).Indoors)
    cmbWeather.ListIndex = Map(GetPlayerMap(MyIndex)).Weather

    For X = 1 To 25
        cmbNpc(X - 1).addItem "No NPC"
        cmbNpcX(X - 1).addItem "Rand"
        cmbNpcY(X - 1).addItem "Rand"
    Next X

    For Y = 1 To MAX_NPCS
        For X = 1 To 25
            cmbNpc(X - 1).addItem Y & ": " & Trim$(Npc(Y).name)
        Next X
    Next Y

    For X = 1 To 25
        cmbNpc(X - 1).ListIndex = Map(GetPlayerMap(MyIndex)).Npc(X)
    Next X

    For X = 1 To 25
        For Y = 0 To MAX_MAPX
            cmbNpcX(X - 1).addItem Y
        Next Y
        cmbNpcX(X - 1).ListIndex = Map(GetPlayerMap(MyIndex)).SpawnX(X)
    Next X
    
    For X = 1 To 25
        For Y = 0 To MAX_MAPY
            cmbNpcY(X - 1).addItem Y
        Next Y
        cmbNpcY(X - 1).ListIndex = Map(GetPlayerMap(MyIndex)).SpawnY(X)
    Next X

    Call StopBGM
End Sub

Private Sub cmdOk_Click()
    Dim I As Integer

    Call StopBGM

    Map(GetPlayerMap(MyIndex)).name = txtMapName.Text
    Map(GetPlayerMap(MyIndex)).Up = Val(txtUp.Text)
    Map(GetPlayerMap(MyIndex)).Down = Val(txtDown.Text)
    Map(GetPlayerMap(MyIndex)).Left = Val(txtLeft.Text)
    Map(GetPlayerMap(MyIndex)).Right = Val(txtRight.Text)
    Map(GetPlayerMap(MyIndex)).Moral = cmbMoral.ListIndex
    Map(GetPlayerMap(MyIndex)).BootMap = Val(txtBootMap.Text)
    Map(GetPlayerMap(MyIndex)).BootX = Val(txtBootX.Text)
    Map(GetPlayerMap(MyIndex)).BootY = Val(txtBootY.Text)
    Map(GetPlayerMap(MyIndex)).Indoors = Val(chkIndoors.value)
    Map(GetPlayerMap(MyIndex)).Weather = cmbWeather.ListIndex

    For I = 1 To 25
        Map(GetPlayerMap(MyIndex)).Npc(I) = cmbNpc(I - 1).ListIndex
        Map(GetPlayerMap(MyIndex)).SpawnX(I) = cmbNpcX(I - 1).ListIndex
        Map(GetPlayerMap(MyIndex)).SpawnY(I) = cmbNpcY(I - 1).ListIndex
    Next I

    If chkURL.value = 0 Then
        Map(GetPlayerMap(MyIndex)).music = lstMusic.Text
    Else
        If Not Left$(txtURL.Text, 7) = "http://" Then
            txtURL.Text = "http://" & txtURL.Text
        End If

        Map(GetPlayerMap(MyIndex)).music = txtURL.Text
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call StopBGM

    Unload Me
End Sub

