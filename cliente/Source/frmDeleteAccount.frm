VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar Cuenta"
   ClientHeight    =   6210
   ClientLeft      =   195
   ClientTop       =   345
   ClientWidth     =   9315
   ControlBox      =   0   'False
   Icon            =   "frmDeleteAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":0FC2
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   2805
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2280
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña de cuenta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de cuenta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   5280
      Width           =   4245
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Width           =   2685
   End
End
Attribute VB_Name = "frmDeleteAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

        If FileExists("GUI\Borrar_Cuenta" & Ending) Then
            frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\Borrar_Cuenta" & Ending)
        End If
    Next i
End Sub

Private Sub picCancel_Click()
    frmDeleteAccount.Visible = False
    frmMainMenu.Visible = True
End Sub

Private Sub picConnect_Click()
    Dim Answer As Long

    If LenB(txtName.Text) < 6 Then
        MsgBox ("Tu usuario debe tener mas de tres caracteres.")
        Exit Sub
    End If

    If LenB(txtPassword.Text) < 6 Then
        MsgBox ("Tu contraseña debe ser mayor de 3 caracteres.")
        Exit Sub
    End If

    Answer = MsgBox("Estas seguro de querer eliminar tu cuenta?", vbYesNo, GAME_NAME)
    If Answer = vbYes Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub
