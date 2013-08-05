VERSION 5.00
Begin VB.Form frmQuestPrompt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Quest Menu"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNpcQuests 
      Height          =   4215
      Left            =   0
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      Begin VB.CommandButton cmdAcceptQuest 
         Caption         =   "Accept"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelQuest 
         Caption         =   "Deny"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtNpcQuestDesc 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmQuestPrompt.frx":0000
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdAbandonQuest 
         Caption         =   "Abandon"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Chaos Knight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblNpcName 
         Alignment       =   2  'Center
         Caption         =   "Npc Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label lblNpcQuestReward 
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "Amount :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblNpcAmount 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1920
         TabIndex        =   5
         Top             =   2160
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmQuestPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcceptQuest_Click()
Call SendData("acceptkillquest" & SEP_CHAR & NpcKillAmount & SEP_CHAR & NpcKillName & SEP_CHAR & NpcKillFinal2 & SEP_CHAR & END_CHAR)
Unload Me
End Sub

Private Sub cmdCancelQuest_Click()
    NpcKillAmount = 0
    NpcKillName = ""
    npckillfinal = ""
    Unload Me
End Sub

Private Sub Form_Load()
lblNpcAmount.Caption = NpcKillAmount
lblNpcName.Caption = NpcKillName
End Sub

