VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmQuest 
   BorderStyle     =   0  'None
   Caption         =   "NPC Speech"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuest.frx":0000
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNpcQuests 
      Height          =   4455
      Left            =   4320
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   447
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   6765
      Begin VB.CommandButton cmdAbandonQuest 
         Caption         =   "Abandon"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   3720
         Visible         =   0   'False
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
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmQuest.frx":2401A
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancelQuest 
         Caption         =   "Deny"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAcceptQuest 
         Caption         =   "Accept"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
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
         Left            =   3600
         TabIndex        =   16
         Top             =   2160
         Width           =   135
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
         Left            =   2040
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblNpcQuestReward 
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   3360
         Width           =   2295
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
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label2 
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
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Abandon Quest"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   8040
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8040
      Top             =   120
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   7680
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   2
      Top             =   3000
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   15
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   3
         Top             =   15
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2700
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4763
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmQuest.frx":240A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Eclipse.jcbutton cmdYes 
      Height          =   495
      Left            =   720
      TabIndex        =   17
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      Caption         =   "Aceptar Quest"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Eclipse.jcbutton lblChoice 
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8454016
      Caption         =   "Continuar Quest"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Eclipse.jcbutton cmdNo 
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "No Aceptar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label lblQuit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   7560
      TabIndex        =   0
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long

'Private Sub cmdAcceptQuest_Click()
'Call SendData("acceptkillquest" & SEP_CHAR & NpcKillAmount & SEP_CHAR & NpcKillName & SEP_CHAR & NpcKillFinal2 & SEP_CHAR & END_CHAR)
'Unload Me
'End Sub

Private Sub cmdCancelQuest_Click()
NpcKillAmount = 0
NpcKillName = ""
NpcKillFinal = ""
    Unload Me
End Sub

Private Sub cmdNo_Click()
Unload Me
End Sub

Private Sub cmdQuit_Click()
Call SendData("STOPKILLQUEST" & SEP_CHAR & END_CHAR)
cmdQuit.Visible = False
lblChoice.Visible = True
End Sub

Private Sub cmdYes_Click()
Call SendData("ACCEPTQUEST" & SEP_CHAR & CurrentQuestNum & SEP_CHAR & CurrentQuestNpcNum & SEP_CHAR & END_CHAR)
cmdYes.Visible = False
cmdNo.Visible = False
lblChoice.Visible = True
CurrentQuestNum = 0
CurrentQuestNpcNum = 0
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
Dim sDc As Long

Dim result As Long
    result = SetWindowLong(txtChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
        
            If FileExists("GUI\Quest" & Ending) Then
            frmQuest.Picture = LoadPicture(App.Path & "\GUI\Quest" & Ending)
        End If
 
    Next I
    
End Sub

Private Sub lblQuit_Click()
    Unload frmQuest
End Sub

Private Sub lblChoice_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Picpic.Width = 32
Picpic.Height = 64
Picture4.Width = 32 + 4
Picture4.Height = 64 + 4
Call BitBlt(Picpic.hDC, 0, 0, 32, 64, picSprites.hDC, animi * 32, Int(Player(MyIndex).Sprite) * 64, SRCCOPY)
End Sub

Private Sub Timer2_Timer()
animi = animi + 1
    If animi > 4 Then
        animi = 3
    End If
End Sub
