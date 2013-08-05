VERSION 5.00
Begin VB.Form frmChatRequest 
   BorderStyle     =   0  'None
   Caption         =   "Chat Request from [PLAYER]!"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChatRequest.frx":0000
   ScaleHeight     =   2475
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   960
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
      Caption         =   "[PLAYER] has requested a chat session."
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
      Top             =   580
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
Attribute VB_Name = "frmChatRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub picAccept_Click()
    Call SendAcceptPlayerChat
    Me.Hide
End Sub

Private Sub picDecline_Click()
    Call SendDeclinePlayerChat
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

