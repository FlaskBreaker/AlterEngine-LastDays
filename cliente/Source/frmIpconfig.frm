VERSION 5.00
Begin VB.Form frmIpconfig 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar IP del servidor"
   ClientHeight    =   5715
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   8070
   ControlBox      =   0   'False
   Icon            =   "frmIpconfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIpconfig.frx":0FC2
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      TabIndex        =   0
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Introduce el puerto del servidor:"
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
      TabIndex        =   5
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Introduce la IP del servidor:"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label PicCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label PicConfirm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   4440
      Width           =   2295
   End
End
Attribute VB_Name = "frmIpconfig"
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

        If FileExists("GUI\Configurar_IP" & Ending) Then
            frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\Configurar_IP" & Ending)
        End If
    Next i

    TxtIP = ReadINI("IPCONFIG", "IP", App.Path & "\config.ini")
    TxtPort = ReadINI("IPCONFIG", "PORT", App.Path & "\config.ini")
    TxtIP.Text = ReadINI("IPCONFIG", "IP", App.Path & "\config.ini")
    TxtPort.Text = ReadINI("IPCONFIG", "PORT", App.Path & "\config.ini")
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False
End Sub

Private Sub picConfirm_Click()
    Dim IP As String, Port As String
    Dim fErr As Integer

    IP = TxtIP
    Port = Val(TxtPort)

    fErr = 0
    If fErr = 0 And Len(Trim$(IP)) = 0 Then
        fErr = 1
        Call MsgBox("Por favor, introduce una IP valida!", vbCritical, GAME_NAME)
        Exit Sub
    End If
    If fErr = 0 And Port <= 0 Then
        fErr = 1
        Call MsgBox("Por favor, introduce un puerto valido!", vbCritical, GAME_NAME)
        Exit Sub
    End If
    If fErr = 0 Then
        Call WriteINI("IPCONFIG", "IP", TxtIP.Text, (App.Path & "\config.ini"))
        Call WriteINI("IPCONFIG", "PORT", TxtPort.Text, (App.Path & "\config.ini"))
    ' Call MenuState(MENU_STATE_IPCONFIG)
    End If
    Call TcpDestroy
    frmMirage.Socket.RemoteHost = TxtIP.Text
    frmMirage.Socket.RemotePort = TxtPort.Text
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False
End Sub
