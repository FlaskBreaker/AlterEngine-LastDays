VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrar"
   ClientHeight    =   5715
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   8070
   ControlBox      =   0   'False
   ForeColor       =   &H00C000C0&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0FC2
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   6120
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   3240
      Width           =   195
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar contraseña."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   3000
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   4440
      Width           =   2370
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Check2.Value = 0
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1.Value = 1
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim Ending As String

    For i = 1 To 3
        If i = 1 Then
            Ending = ".gif"
        End If
        If i = 2 Then
            Ending = ".jpg"
        End If
        If i = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Entrar" & Ending) Then
            frmLogin.Picture = LoadPicture(App.Path & "\GUI\Entrar" & Ending)
        End If
    Next i

    frmLogin.txtName.Text = Trim$(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim$(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))

    If AutoLogin = 1 Then
        Check2.Value = Checked
        Check1.Value = Checked
    End If

    If LenB(frmLogin.txtPassword.Text) <> 0 Then
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If AllDataReceived Then
        If LenB(txtName.Text) < 6 Then
            Call MsgBox("Tu usuario debe tener más de 3 caracteres de largo.")
            Exit Sub
        End If
    
        If LenB(txtPassword.Text) < 6 Then
            Call MsgBox("Tu contraseña debe tener más de 3 caracteres de largo.")
            Exit Sub
        End If

        Call WriteINI("CONFIG", "Account", txtName.Text, (App.Path & "\config.ini"))

        If Check1.Value = Checked Then
            Call WriteINI("CONFIG", "Password", txtPassword.Text, (App.Path & "\config.ini"))
        Else
            Call WriteINI("CONFIG", "Password", vbNullString, (App.Path & "\config.ini"))
        End If

        If Check2.Value = Checked Then
            Call WriteINI("CONFIG", "Auto", 1, (App.Path & "\config.ini"))
        Else
            Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))
        End If

        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub
