VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo de Clase"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmBClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
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
      TabCaption(0)   =   "Establecer"
      TabPicture(0)   =   "frmBClass.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNum1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblNum2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblNum3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancel"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlNum1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlNum2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlNum3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.HScrollBar scrlNum3 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   7
         Top             =   1800
         Width           =   4335
      End
      Begin VB.HScrollBar scrlNum2 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   6
         Top             =   1200
         Width           =   4335
      End
      Begin VB.HScrollBar scrlNum1 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   2160
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
         Caption         =   "Aceptar"
         PictureNormal   =   "frmBClass.frx":0FDE
         PictureHot      =   "frmBClass.frx":17C2
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
         TabIndex        =   11
         Top             =   2160
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
         Caption         =   "Cancelar"
         PictureNormal   =   "frmBClass.frx":1FA6
         PictureHot      =   "frmBClass.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblNum3 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label lblNum2 
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
         Left            =   1200
         TabIndex        =   8
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Permitir clase 3:"
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
         Left            =   150
         TabIndex        =   5
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Permitir clase 2:"
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
         Left            =   150
         TabIndex        =   4
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Permitir clase 1:"
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
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblNum1 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlNum1.value = 0
    scrlNum2.value = 0
    scrlNum3.value = 0

    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblNum1.Caption = scrlNum1.value & " - " & Trim$(Class(scrlNum1.value).name)
    lblNum2.Caption = scrlNum2.value & " - " & Trim$(Class(scrlNum2.value).name)
    lblNum3.Caption = scrlNum3.value & " - " & Trim$(Class(scrlNum3.value).name)

    If EditorItemNum1 < scrlNum1.min Then
        EditorItemNum1 = scrlNum1.min
    End If

    scrlNum1.value = EditorItemNum1

    If EditorItemNum2 < scrlNum2.min Then
        EditorItemNum2 = scrlNum2.min
    End If

    scrlNum2.value = EditorItemNum2

    If EditorItemNum3 < scrlNum3.min Then
        EditorItemNum3 = scrlNum3.min
    End If

    scrlNum3.value = EditorItemNum3
End Sub

Private Sub scrlNum1_Change()
    lblNum1.Caption = scrlNum1.value & " - " & Trim$(Class(scrlNum1.value).name)
    EditorItemNum1 = scrlNum1.value
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = scrlNum2.value & " - " & Trim$(Class(scrlNum2.value).name)
    EditorItemNum2 = scrlNum2.value
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = scrlNum3.value & " - " & Trim$(Class(scrlNum3.value).name)
    EditorItemNum3 = scrlNum3.value
End Sub
