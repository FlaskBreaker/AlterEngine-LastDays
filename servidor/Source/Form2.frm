VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   1650
   ClientLeft      =   9270
   ClientTop       =   1860
   ClientWidth     =   2700
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2280
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   1335
      Width           =   2700
      Begin VB.Image salir 
         Height          =   240
         Left            =   2400
         MouseIcon       =   "Form2.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":030A
         Top             =   15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   60
         Picture         =   "Form2.frx":0694
         Top             =   40
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RSS Reader"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   420
         TabIndex        =   1
         Top             =   15
         Width           =   1455
      End
   End
   Begin VB.Label canti 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1140
      TabIndex        =   4
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivos Nuevos..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hay "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   675
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Me.Left = Screen.Width - Me.Width
Me.Top = 0 - Me.Height
End Sub

Private Sub Form_Load()
Alerta = LoadResData(101, "CUSTOM")
Timer1.Enabled = True
Timer1.Interval = 1
canti = Form1.Resultado
Call sndPlaySound(Alerta(0), SND_SYNC Or SND_MEMORY)
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top + 30
If Me.Top = 0 Then
Dim retraso As Long
retraso = 3500 + GetTickCount&
While retraso >= GetTickCount&
    DoEvents
Wend
Unload Form2
End If
End Sub
