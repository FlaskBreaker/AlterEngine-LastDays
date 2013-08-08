VERSION 5.00
Begin VB.Form frmMainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Principal"
   ClientHeight    =   7200
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   9600
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Status 
      Interval        =   2000
      Left            =   0
      Top             =   1080
   End
   Begin VB.Label tips 
      Alignment       =   2  'Center
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
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   7215
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
      Height          =   4095
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   6015
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
      Height          =   1095
      Left            =   3480
      TabIndex        =   2
      Top             =   6120
      Width           =   2580
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
      Height          =   1095
      Left            =   6120
      TabIndex        =   1
      Top             =   6120
      Width           =   2985
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
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   6120
      Width           =   2625
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

        If FileExists("GUI\Menu_Principal-on" & Ending) Then
            frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\Menu_Principal-on" & Ending)
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

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub


Private Sub picNewAccount_Click()
    Me.Visible = False
    frmNewAccount.Visible = True
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

Private Sub Status_Timer()
    If ConnectToServer = True Then
        Dim I As Long
        Dim Ending As String
        If Not AllDataReceived Then
            Call SendData("givemethemax" & END_CHAR)
        Else
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
        
                If FileExists("GUI\Menu_Principal-on" & Ending) Then
                    frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\Menu_Principal-on" & Ending)
                End If
            Next I
        End If
    Else
        picNews.Caption = "No se ha podido conectar. El servidor puede que este apagado."
        
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
        
            If FileExists("GUI\Menu_Principal-off" & Ending) Then
                frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\Menu_Principal-off" & Ending)
            End If
        Next I
    End If
End Sub
