VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creditos | www.AlterEngine.net"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0FC2
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   120
      Top             =   120
   End
   Begin VB.Image creditos 
      Height          =   4095
      Left            =   480
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   6000
      Width           =   3120
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String

    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
        
        ' Parte para cargar el recuadro de creditos.
        
                If FileExists("GUI\Creditos_Texto" & Ending) Then
            creditos.Picture = LoadPicture(App.Path & "\GUI\Creditos_Texto" & Ending)
        End If
    Next I
    

End Sub

Private Sub Image1_Click()

End Sub

Private Sub picCancel_Click()

    frmCredits.Visible = False
    frmMainMenu.Visible = True
    
    Unload Me
End Sub

Private Sub picCredits_Click()

End Sub

