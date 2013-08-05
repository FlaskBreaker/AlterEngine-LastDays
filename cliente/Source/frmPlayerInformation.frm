VERSION 5.00
Begin VB.Form frmPlayerInformation 
   BorderStyle     =   0  'None
   Caption         =   "Player Information"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerInformation.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Display 
      Interval        =   1
      Left            =   6840
      Top             =   960
   End
   Begin VB.Label picClose 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.Label picReport 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label picParty 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label picChat 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label picTrade 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblJob 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Job"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A5015&
      Height          =   255
      Left            =   4995
      TabIndex        =   2
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name [Guild]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A5015&
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmPlayerInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Single
Dim OldY As Single
 
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        OldX = X
        OldY = Y
    End If
End Sub
 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (X - OldX), Top + (Y - OldY)
End Sub


Private Sub picChat_Click()
    Call SendPlayerChat(lblName.Caption)
End Sub

Private Sub picClose_Click()
    Me.Hide
End Sub

Private Sub picParty_Click()
    MsgBox "This feature is coming soon."
End Sub

Private Sub picReport_Click()
    MsgBox "This feature is coming soon. However, if you need to report this player to a moderator or a Game Master, please do so at our website.", vbInformation, "Sylerean Online"
End Sub

Private Sub picTrade_Click()
    Call SendTradeRequest(lblName.Caption)
End Sub

