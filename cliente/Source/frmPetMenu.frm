VERSION 5.00
Begin VB.Form frmPetMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Pet Menu"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPetMenu.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPet 
      Caption         =   "Elección de Mascota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   5175
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   120
         Max             =   130
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1680
         ScaleHeight     =   62
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   18
         Top             =   1560
         Width           =   2415
      End
      Begin Eclipse.jcbutton Command1 
         Height          =   495
         Left            =   960
         TabIndex        =   35
         Top             =   3000
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
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
         BackColor       =   16576
         Caption         =   "Elegir Mascota"
         ForeColor       =   16777215
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Selecciona la apariencia de tu mascota:"
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
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblSprite 
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
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre de tu Mascota :"
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
         Left            =   960
         TabIndex        =   21
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4680
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4680
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4200
      Top             =   0
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2280
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   14
      Top             =   6960
      Width           =   3375
      Begin VB.Label lblExecute 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   1560
         Width           =   255
      End
   End
   Begin Eclipse.jcbutton KillPet 
      Height          =   495
      Left            =   1560
      TabIndex        =   36
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      BackColor       =   16512
      Caption         =   "Matar Mascota"
      ForeColor       =   16777215
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1200
      TabIndex        =   34
      Top             =   4680
      Width           =   2385
   End
   Begin VB.Label lblSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1200
      TabIndex        =   33
      Top             =   4440
      Width           =   2385
   End
   Begin VB.Label lblFP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1200
      TabIndex        =   32
      Top             =   4920
      Width           =   2385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comida:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   360
      TabIndex        =   31
      Top             =   4920
      Width           =   780
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   30
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EXP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   29
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   840
      TabIndex        =   28
      Top             =   4200
      Width           =   300
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   27
      Top             =   3960
      Width           =   270
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB884B&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1200
      TabIndex        =   26
      Top             =   4200
      Width           =   2385
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1200
      TabIndex        =   25
      Top             =   3960
      Width           =   2385
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblAlive 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vivo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label AddMagi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3360
      TabIndex        =   12
      Top             =   2760
      Width           =   105
   End
   Begin VB.Label AddSpeed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3360
      TabIndex        =   11
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label AddDef 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3360
      TabIndex        =   10
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label AddStr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3360
      TabIndex        =   9
      Top             =   2280
      Width           =   105
   End
   Begin VB.Label lblMAGI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Magia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label lblSPEED 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Label lblDEF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   1785
   End
   Begin VB.Label lblSTR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuerza"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lblLevel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   " Pet Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   6600
      Width           =   1830
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   165
      Left            =   1230
      Top             =   3975
      Width           =   2370
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00CB884B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   165
      Left            =   1215
      Top             =   4215
      Width           =   2370
   End
   Begin VB.Shape shpSP 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF80&
      Height          =   165
      Left            =   1215
      Top             =   4455
      Width           =   2370
   End
   Begin VB.Shape shpFP 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   165
      Left            =   1200
      Top             =   4935
      Width           =   2385
   End
   Begin VB.Shape shpTNL 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   165
      Left            =   1215
      Top             =   4695
      Width           =   2370
   End
End
Attribute VB_Name = "frmPetMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lparam As Any) As Long

      Private Declare Sub ReleaseCapture Lib "user32" ()

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2

Dim PetChoice As Long
Dim animi As Long

Private Sub AddDef_Click()
Call SendData("usepetstatpoint" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMagi_Click()
Call SendData("usepetstatpoint" & SEP_CHAR & 3 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
Call SendData("usepetstatpoint" & SEP_CHAR & 2 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddStr_Click()
Call SendData("usepetstatpoint" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command1_Click()
Call SendData("CHOOSEPET" & SEP_CHAR & PetChoice & SEP_CHAR & Trim$(txtName.Text) & SEP_CHAR & END_CHAR)
fraPet.Visible = False
End Sub

Private Sub KillPet_Click()
Call SendData("KILLPET" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label38_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
         Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(frmPetMenu.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
      End Sub

Private Sub scrlSprite_Change()
lblSprite.Caption = STR(scrlSprite.value)
PetChoice = scrlSprite.value
End Sub

Private Sub Timer2_Timer()
Dim sDc As Long
sDc = DD_PetSpriteSurf.GetDC
'Call BitBlt(Picpic.hdc, 0, 0, 32, 32, sDc, animi * SIZE_X, PetChoice * SIZE_Y, SRCCOPY)
Call BitBlt(picPic.hDC, 0, 0, 32, 32, sDc, animi * 32, PetChoice * 32, SRCCOPY)
Call DD_PetSpriteSurf.ReleaseDC(sDc)
End Sub


Private Sub Timer1_Timer()
If PetAlive > 0 Then

If PetPoints > 0 Then
AddStr.Visible = True
AddDef.Visible = True
AddSpeed.Visible = True
AddMagi.Visible = True
Else
AddStr.Visible = False
AddDef.Visible = False
AddSpeed.Visible = False
AddMagi.Visible = False
End If

KillPet.Visible = True
lblAlive.Caption = "Vivo: Si"
frmPetMenu.lblHP.Caption = PetHP & " / " & PetMaxHP
If PetHP <= 0 Then
    frmPetMenu.shpHP.Width = 0
Else
    frmPetMenu.shpHP.Width = (((PetHP / frmPetMenu.lblHP.Width) / (PetMaxHP / frmPetMenu.lblHP.Width)) * frmPetMenu.lblHP.Width)
End If

frmPetMenu.lblSP.Caption = PetSP & " / " & PetMaxSP

If PetSP <= 0 Then
    frmPetMenu.shpSP.Width = 0
Else
    frmPetMenu.shpSP.Width = (((PetSP / frmPetMenu.lblSP.Width) / (PetMaxSP / frmPetMenu.lblSP.Width)) * frmPetMenu.lblSP.Width)
End If

frmPetMenu.lblMP.Caption = PetMP & " / " & PetMaxMP

If PetMP <= 0 Then
    frmPetMenu.shpMP.Width = 0
Else
    frmPetMenu.shpMP.Width = (((PetMP / frmPetMenu.lblMP.Width) / (PetMaxMP / frmPetMenu.lblMP.Width)) * frmPetMenu.lblMP.Width)
End If

frmPetMenu.lblFP.Caption = PetFP & " / " & PetMaxFP

If PetFP <= 0 Then
    frmPetMenu.shpFP.Width = 0
Else
    frmPetMenu.shpFP.Width = (((PetFP / frmPetMenu.lblFP.Width) / (PetMaxFP / frmPetMenu.lblFP.Width)) * frmPetMenu.lblFP.Width)
End If

frmPetMenu.lblExp.Caption = PetExp & " / " & PetNextLevel
If PetExp <= 0 Then
    frmPetMenu.shpTNL.Width = 0
Else
    frmPetMenu.shpTNL.Width = (((PetExp / frmPetMenu.lblExp.Width) / (PetNextLevel / frmPetMenu.lblExp.Width)) * frmPetMenu.lblExp.Width)
End If

lblName.Caption = "Nombre: " & PetName
lblSTR.Caption = "Fuerza: " & PetSTR
lblDEF.Caption = "Defensa: " & PetDEF
lblSPEED.Caption = "Velocidad: " & PetSPEED
lblMAGI.Caption = "Magia: " & PetMAGI
lblLevel.Caption = "Nivel: " & PetLevel
lblPoints.Caption = "Puntos: " & PetPoints
Else
lblAlive.Caption = "Vivo: No"
lblName.Caption = "Nombre: Ninguno"
lblHP.Caption = "PV: Ninguno"
lblSTR.Caption = "Fuerza: Ninguno"
lblDEF.Caption = "Defensa: Ninguno"
lblSPEED.Caption = "Velocidad: Ninguno"
lblMAGI.Caption = "Magia: Ninguno"
lblLevel.Caption = "Nivel: Ninguno"
KillPet.Visible = False
End If
End Sub

Private Sub Form_Load()
Dim sDc As Long

    Dim I As Long
    Dim Ending As String
    For I = 1 To 3
        If I = 1 Then
            Ending = ".gif"
        End If
        If I = 2 Then
            Ending = ".jpg"
        End If
        If I = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Mascota" & Ending) Then
            frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\Mascota" & Ending)
        End If
    Next I

PetChoice = 0
    
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub
