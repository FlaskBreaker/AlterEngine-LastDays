VERSION 5.00
Begin VB.Form frmMainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Principal"
   ClientHeight    =   7245
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   10845
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Status 
      Interval        =   2000
      Left            =   10320
      Top             =   120
   End
   Begin VB.Label tips 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   45
      Width           =   7215
   End
   Begin VB.Label lblOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando.."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label picNews 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Noticias..."
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
      Height          =   3495
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label picAutoLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   975
      Left            =   1920
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado del server:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label picIpConfig 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   6480
      Width           =   2580
   End
   Begin VB.Label picLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7560
      TabIndex        =   4
      Top             =   2640
      Width           =   2700
   End
   Begin VB.Label picNewAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7560
      TabIndex        =   3
      Top             =   3480
      Width           =   2625
   End
   Begin VB.Label picDeleteAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7560
      TabIndex        =   2
      Top             =   4200
      Width           =   2580
   End
   Begin VB.Label picCredits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   2625
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   5040
      Width           =   2700
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
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

        If FileExists("GUI\Menu_Principal" & Ending) Then
            frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\Menu_Principal" & Ending)
        End If
    Next I
    
    'Musica en inicio
    Ending = ReadINI("CONFIG", "MenuMusic", App.Path & "\Config.ini")
    If LenB(Ending) <> 0 Then
        MapSound = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\Musica\" & Ending), 0, 0, 0)
        Call BASS_ChannelPlay(MapSound, BASSFALSE)
    End If
    
    'Sistema de consejos
    Dim randos
    randos = Random(1, 15)
    tips.Caption = ReadINI("CONSEJOS", randos & "", App.Path & "\Consejos.ini")

    Call MainMenuInit
End Sub

Function Random(Lowerbound As Long, Upperbound As Long)
Randomize
Random = Int(Rnd * Upperbound) + Lowerbound
End Function

Private Sub Form_GotFocus()
    If frmMirage.Socket.State = 0 Then
        frmMirage.Socket.Connect
    End If
End Sub

Private Sub picAutoLogin_Click()
    If ConnectToServer = False Or (ConnectToServer = True And AutoLogin = 1 And AllDataReceived) Then
        Call MenuState(MENU_STATE_AUTO_LOGIN)
    End If
End Sub

Private Sub picIpConfig_Click()
    Me.Visible = False
    frmIpconfig.Visible = True
End Sub

Private Sub picNewAccount_Click()
    Me.Visible = False
    frmNewAccount.Visible = True
End Sub

Private Sub picDeleteAccount_Click()
    frmDeleteAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picLogin_Click()
    If LenB(frmLogin.txtPassword.Text) <> 0 Then
        frmLogin.Check1.value = Checked
    Else
        frmLogin.Check1.value = Unchecked
    End If

    Me.Visible = False
    frmLogin.Visible = True
End Sub

Private Sub picCredits_Click()
    Me.Visible = False
    frmCredits.Visible = True
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub Status_Timer()
    If ConnectToServer = True Then
        If Not AllDataReceived Then
            Call SendData("givemethemax" & END_CHAR)
        Else
            lblOnline.Caption = "Encendido"
            lblOnline.ForeColor = vbBlue
        End If
    Else
        picNews.Caption = "No se ha podido conectar. El servidor puede que este apagado."

        lblOnline.Caption = "Apagado"
        lblOnline.ForeColor = vbRed
    End If
End Sub
