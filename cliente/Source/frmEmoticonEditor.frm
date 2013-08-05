VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Emoticonos"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4455
   ControlBox      =   0   'False
   Icon            =   "frmEmoticonEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   108
      Left            =   240
      Top             =   360
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   5054
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   2469
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Emoticonos"
      TabPicture(0)   =   "frmEmoticonEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEmoticon"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlEmoticon"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCommand"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtCommand 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "/"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.HScrollBar scrlEmoticon 
         Height          =   255
         Left            =   840
         Max             =   1000
         TabIndex        =   3
         Top             =   960
         Value           =   1
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   1920
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   1
         Top             =   360
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   4
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picEmoticons 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   6
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin Eclipse.jcbutton cmdOk 
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Aceptar"
         PictureNormal   =   "frmEmoticonEditor.frx":0FDE
         PictureHot      =   "frmEmoticonEditor.frx":17C2
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton Command1 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Cancelar"
         PictureNormal   =   "frmEmoticonEditor.frx":1FA6
         PictureHot      =   "frmEmoticonEditor.frx":28FA
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comando :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   7
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label lblEmoticon 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3480
         TabIndex        =   5
         Top             =   720
         Width           =   75
      End
      Begin VB.Label Label5 
         Caption         =   "Emoticono :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmoticonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOk_Click()
    Dim I As Long

    For I = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(I).Command) = Trim$(txtCommand.Text) And I <> EditorIndex - 1 And Trim$(txtCommand.Text) <> "/" Then
            MsgBox "Ese comando ya esta siendo usado por " & Trim$(txtCommand.Text) & " !"
            Exit Sub
        End If
    Next I
    Call EmoticonEditorOk
End Sub

Private Sub Command1_Click()
    Call EmoticonEditorCancel
End Sub

Private Sub Form_Load()
    picEmoticons.top = (scrlEmoticon.value * 32) * -1
End Sub

Private Sub scrlEmoticon_Change()
    picEmoticons.top = (scrlEmoticon.value * 32) * -1
    lblEmoticon.Caption = scrlEmoticon.value
End Sub

Private Sub Timer1_Timer()
    If picEmoticons.Left < -(10 * 32) Then
        picEmoticons.Left = 0
    End If
    picEmoticons.Left = picEmoticons.Left - 32
End Sub

Private Sub txtCommand_Change()
    Dim I As String
    I = txtCommand.Text
    If Mid$(I, 1, 1) <> "/" Then
        If Trim$(I) = vbNullString Then
            txtCommand.Text = "/"
            Exit Sub
        End If
        txtCommand.Text = "/" & I
        txtCommand.SelStart = Len(txtCommand.Text)
    End If
End Sub
