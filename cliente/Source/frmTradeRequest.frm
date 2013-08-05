VERSION 5.00
Begin VB.Form frmTradeRequest 
   BorderStyle     =   0  'None
   Caption         =   "Trade Request from [PLAYER]!"
   ClientHeight    =   2475
   ClientLeft      =   6450
   ClientTop       =   4035
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTradeRequest.frx":0000
   ScaleHeight     =   2475
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Display 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label picAccept 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label picDecline 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "[PLAYER] has requested a trade."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A5015&
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   575
      Width           =   4215
   End
   Begin VB.Label Accept 
      BackStyle       =   0  'Transparent
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A5015&
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1245
      Width           =   1935
   End
   Begin VB.Label lblDecline 
      BackStyle       =   0  'Transparent
      Caption         =   "Decline"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A5015&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1365
      Width           =   1815
   End
End
Attribute VB_Name = "frmTradeRequest"
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
Private Sub picAccept_Click()
    Call SendAcceptTrade
    Me.Hide
End Sub

Private Sub picDecline_Click()
    Call SendDeclineTrade
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
