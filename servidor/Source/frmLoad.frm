VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AlterEngine - Servidor v1.8"
   ClientHeight    =   4260
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9585
   ClipControls    =   0   'False
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   Picture         =   "frmLoad.frx":0FC2
   ScaleHeight     =   4260
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Server.ProgressBar loadProgressBar 
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   2160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Cargando..."
      TextEffect      =   5
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comenzando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "versión 1.8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Show
    Call InitServer
End Sub

