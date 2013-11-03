VERSION 5.00
Begin VB.Form frmLibro 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   " X"
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Me.Hide
End Sub
