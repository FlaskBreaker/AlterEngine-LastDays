VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Personaje"
   ClientHeight    =   6930
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   9315
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0FC2
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
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
      Height          =   2010
      ItemData        =   "frmChars.frx":18CD1
      Left            =   1920
      List            =   "frmChars.frx":18CD3
      TabIndex        =   0
      Top             =   1680
      Width           =   5565
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1320
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   2640
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   4560
      Width           =   2640
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2640
      TabIndex        =   2
      Top             =   5880
      Width           =   4215
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   4560
      Width           =   2640
   End
End
Attribute VB_Name = "frmChars"
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

        If FileExists("GUI\Seleccion_Personaje" & Ending) Then
            frmChars.Picture = LoadPicture(App.Path & "\GUI\Seleccion_Personaje" & Ending)
        End If
    Next I
End Sub

Private Sub Label1_Click()
    If AutoLogin = 1 Then
        Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))

        Me.Visible = False
        frmMainMenu.Visible = True
    End If
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy

    Me.Visible = False
    frmLogin.Visible = True
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Hueco Libre" Then
        MsgBox "Ya hay un personaje en este hueco!"
        Exit Sub
    End If

    frmNewChar.picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Hueco Libre" Then
        MsgBox "No hay un personaje en este hueco!"
        Exit Sub
    End If

    frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\Items.bmp")
    frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    Call MenuState(MENU_STATE_USECHAR)
    picUseChar.Enabled = False
End Sub

Private Sub picDelChar_Click()
    Dim value As Integer

    If lstChars.List(lstChars.ListIndex) = "Hueco Libre" Then
        MsgBox "No hay un personaje en este hueco!"
        Exit Sub
    End If

    value = MsgBox("Estas seguro de querer eliminar el personaje?", vbYesNo, GAME_NAME)
    If value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub
