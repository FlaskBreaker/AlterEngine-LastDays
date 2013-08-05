VERSION 5.00
Begin VB.Form frmAnuncio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar anuncio a los jugadores."
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "El anuncio será visible durante 10 segundos en pantalla."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox anunciotxt 
         Appearance      =   0  'Flat
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   6375
      End
   End
   Begin Server.jcbutton enviaranuncio 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   6615
      _ExtentX        =   11668
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
      Caption         =   "Enviar anuncio a los jugadores ->"
      PictureNormal   =   "frmAnuncio.frx":0000
      CaptionEffects  =   4
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmAnuncio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub enviaranuncio_Click()
Call Anuncio(anunciotxt.Text)
anunciotxt.Text = ""
MsgBox "Anuncio enviado correctamente", vbInformation
End Sub
