VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de color"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin Server.jcbutton Command1 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Aceptar"
      CaptionEffects  =   0
   End
   Begin Server.jcbutton Command2 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Cancelar"
      CaptionEffects  =   0
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   4920
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R As Integer
Dim G As Integer
Dim B As Integer

Private Sub Command1_Click()
    frmNews.red = Slider1.Value
    frmNews.Green = Slider2.Value
    frmNews.Blue = Slider3.Value
    frmNews.Text1.ForeColor = RGB(Slider1.Value, Slider2.Value, Slider3.Value)
    frmNews.Text2.ForeColor = RGB(Slider1.Value, Slider2.Value, Slider3.Value)
    Call PutVar(App.Path & "\Noticias.ini", "Color", "Red", "" & R)
    Call PutVar(App.Path & "\Noticias.ini", "Color", "Green", "" & G)
    Call PutVar(App.Path & "\Noticias.ini", "Color", "Blue", "" & B)
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Slider1.Value = frmNews.red
    Slider2.Value = frmNews.Green
    Slider3.Value = frmNews.Blue
End Sub

Private Sub Slider1_Change()
    R = Slider1.Value
    Shape1.BackColor = RGB(R, G, B)
End Sub
Private Sub Slider2_Change()
    G = Slider2.Value
    Shape1.BackColor = RGB(R, G, B)
End Sub
Private Sub Slider3_Change()
    B = Slider3.Value
    Shape1.BackColor = RGB(R, G, B)
End Sub

