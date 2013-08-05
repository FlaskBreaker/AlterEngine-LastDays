VERSION 5.00
Begin VB.Form frmParty 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Grupo"
   ClientHeight    =   1410
   ClientLeft      =   270
   ClientTop       =   75
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmParty.frx":0000
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRedCheck 
      Height          =   495
      Left            =   4320
      Picture         =   "frmParty.frx":8E3D
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picGreenCheck 
      Height          =   495
      Left            =   3120
      Picture         =   "frmParty.frx":8F7E
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   2400
      Picture         =   "frmParty.frx":90AE
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   20
      Top             =   840
      Width           =   630
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   1320
      Picture         =   "frmParty.frx":922B
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   19
      Top             =   840
      Width           =   630
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   240
      Picture         =   "frmParty.frx":93A8
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   18
      Top             =   840
      Width           =   630
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   3480
      Picture         =   "frmParty.frx":9525
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   17
      Top             =   840
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   3480
      Picture         =   "frmParty.frx":96A2
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   16
      Top             =   720
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   2400
      Picture         =   "frmParty.frx":9827
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   15
      Top             =   720
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   1320
      Picture         =   "frmParty.frx":99AC
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   14
      Top             =   720
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   240
      Picture         =   "frmParty.frx":9B31
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   13
      Top             =   720
      Width           =   630
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   4200
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   3720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   5400
      Width           =   240
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   3480
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   -240
      Top             =   0
   End
   Begin VB.Label lblLeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo liderador por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4440
      TabIndex        =   24
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4440
      TabIndex        =   23
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
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
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
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
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
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
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
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
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sitio Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sitio Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sitio Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sitio Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Private Declare Sub ReleaseCapture Lib "user32" ()

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2
Public Function Transparent(Form As Form, Layout As Byte) As Boolean
    SetWindowLong Form.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Form.hWnd, 0, Layout, LWA_ALPHA
    Transparent = Err.LastDllError = 0
End Function

Private Sub Form_Load()
'Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
'SetLayeredWindowAttributes Me.hWnd, RGB(255, 0, 255), 0, 1

frmParty.Picture = LoadPicture(App.Path & "\GUI\Grupo.jpg")

    Me.Icon = frmMirage.Icon
    
    frmParty.picItems.Width = 480
    frmParty.picItems.Height = 720
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then

    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                                 x As Single, y As Single)
         Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(frmParty.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
      End Sub

Private Sub Label1_Click()
Call InvitePlayer
End Sub

Private Sub Label2_Click()
Call RemoveMember
End Sub

Private Sub Label3_Click()
Dim I As Byte
            Call SendLeaveParty
            If frmParty.Visible = True Then
            Unload frmParty
            For I = 1 To MAX_PARTY_INV_SLOTS
            'Player(MyIndex).Party.PartyItems(i).Num = 0
            Next I
            End If
            Unload Me
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then
''Call SendRoll(Index)
'ElseIf Button = 2 Then
'Call SendNoRoll(Index)
'End If
'frmToW.Text1.SetFocus
End Sub

Private Sub picSprite_Click(Index As Integer)
'Call SendTarget(index)
'frmToW.Text1.SetFocus
End Sub

Private Sub tmrSprite_Timer()
'Call PartyBltSprite
End Sub

