VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Begin VB.Form frmFlash 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evento Flash"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFlash.frx":0FC2
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Check 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   4920
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   4800
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7890
      _cx             =   13917
      _cy             =   8467
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   2760
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Timer()
    If Flash.CurrentFrame > 0 Then
        If Flash.CurrentFrame >= Flash.TotalFrames - 1 Then
            Flash.FrameNum = 0
            Flash.Stop
            Check.Enabled = False
            WriteINI "CONFIG", "Music", frmMirage.chkMusic.value, App.Path & "\config.ini"
            Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
            WriteINI "CONFIG", "Sound", frmMirage.chkSound.value, App.Path & "\config.ini"
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String
    For I = 1 To 3
        If I = 1 Then
            Ending = ".gif"
        End If
        If I = 2 Then
            Ending = ".jpg"
        End If
        If I = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Reproductor_Flash" & Ending) Then
            frmFlash.Picture = LoadPicture(App.Path & "\GUI\Reproductor_Flash" & Ending)
        End If
    Next I
End Sub

Private Sub Label1_Click()
    Flash.FrameNum = 0
    Flash.Stop
    Check.Enabled = False
    WriteINI "CONFIG", "Music", frmMirage.chkMusic.value, App.Path & "\config.ini"
    Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
    WriteINI "CONFIG", "Sound", frmMirage.chkSound.value, App.Path & "\config.ini"
    Unload Me
End Sub

