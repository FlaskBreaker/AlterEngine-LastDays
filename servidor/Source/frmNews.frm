VERSION 5.00
Begin VB.Form frmNews 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de Noticias"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin Server.jcbutton cmdColor 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
      _extentx        =   2566
      _extenty        =   661
      buttonstyle     =   3
      font            =   "frmNews.frx":0000
      backcolor       =   14935011
      caption         =   "Cambiar Color"
      captioneffects  =   4
      tooltip         =   "Expulsa a todos los jugadores conectados al servidor."
      tooltiptype     =   1
      colorscheme     =   2
   End
   Begin Server.jcbutton cmdOK 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   3
      font            =   "frmNews.frx":0028
      backcolor       =   14935011
      caption         =   "Aceptar"
      picturenormal   =   "frmNews.frx":0050
      picturehot      =   "frmNews.frx":09A6
      captioneffects  =   4
      colorscheme     =   2
   End
   Begin Server.jcbutton cmdCancel 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      buttonstyle     =   3
      font            =   "frmNews.frx":12FA
      backcolor       =   14935011
      caption         =   "Cancelar"
      picturenormal   =   "frmNews.frx":1322
      picturehot      =   "frmNews.frx":1C78
      captioneffects  =   4
      colorscheme     =   2
   End
   Begin VB.Label lblNews 
      Caption         =   "Contenido:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblCaption 
      Caption         =   "Titulo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public red
Public Green
Public Blue

Private Sub cmdOK_Click()
    Call PutVar(App.Path & "\Noticias.ini", "Data", "NewsTitle", Text1.Text)
    Call PutVar(App.Path & "\Noticias.ini", "Data", "NewsBody", Text2.Text)

MsgBox "Recuerda que debes darle a Enviar Noticia para que aparezca", vbOKOnly
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click()
    frmColor.Visible = True
End Sub

Private Sub Form_Load()
    Text1.Text = GetVar(App.Path & "\Noticias.ini", "Data", "NewsTitle")
    Text2.Text = GetVar(App.Path & "\Noticias.ini", "Data", "NewsBody")

    red = GetVar(App.Path & "\Noticias.ini", "Color", "Red")
    Green = GetVar(App.Path & "\Noticias.ini", "Color", "Green")
    Blue = GetVar(App.Path & "\Noticias.ini", "Color", "Blue")

    Text1.ForeColor = RGB(red, Green, Blue)
    Text2.ForeColor = RGB(red, Green, Blue)
End Sub
