VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Personaje"
   ClientHeight    =   7230
   ClientLeft      =   135
   ClientTop       =   315
   ClientWidth     =   9825
   ControlBox      =   0   'False
   Icon            =   "frmNewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0FC2
   ScaleHeight     =   482
   ScaleMode       =   0  'User
   ScaleWidth      =   650.112
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   5760
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   31
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   36
         Top             =   720
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   2
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   37
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   34
         Top             =   360
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   35
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   32
         Top             =   0
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   33
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   6480
      Max             =   200
      TabIndex        =   29
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   6480
      Max             =   200
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   6480
      Max             =   200
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mujer"
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
      Height          =   240
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3900
      UseMaskColor    =   -1  'True
      Width           =   2040
   End
   Begin VB.OptionButton optMale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hombre"
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
      Height          =   240
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   6240
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   6
      Top             =   2640
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   7
         Top             =   15
         Width           =   495
         Begin VB.PictureBox Picsprites 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DrawStyle       =   5  'Transparent
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   24
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   1080
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1560
      Top             =   720
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Height          =   270
      Left            =   5280
      TabIndex        =   40
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona una clase:"
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
      Height          =   270
      Left            =   2160
      TabIndex        =   39
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblClassDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label13"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   5280
      TabIndex        =   38
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Piernas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   30
      Top             =   2760
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerpo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Cabeza:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   23
      Top             =   5040
      Width           =   600
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   22
      Top             =   5040
      Width           =   840
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   21
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2040
      TabIndex        =   20
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fuerza:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   15
      Top             =   5280
      Width           =   600
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Magia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   5280
      Width           =   600
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      TabIndex        =   13
      Top             =   5040
      Width           =   600
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PS:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   5040
      Width           =   600
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      TabIndex        =   11
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PM:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   10
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      TabIndex        =   9
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PV:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   8
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de tu personaje:"
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
      Height          =   270
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label picAddChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   6360
      Width           =   2205
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   6360
      Width           =   2235
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public animi As Long

Private Sub cmbClass_Click()
    Dim I As Byte
   
    For I = 0 To Max_Classes
        If Trim(Class(I).name) = cmbClass.List(cmbClass.ListIndex) Then
            Exit For
        End If
    Next
    'MsgBox (cmbClass.ListIndex)
    lblHP.Caption = STR(Class(I).HP)
    lblMP.Caption = STR(Class(I).MP)
    lblSP.Caption = STR(Class(I).SP)

    lblSTR.Caption = STR(Class(I).STR)
    lblDEF.Caption = STR(Class(I).DEF)
    lblSPEED.Caption = STR(Class(I).SPEED)
    lblMAGI.Caption = STR(Class(I).MAGI)

    lblClassDesc.Caption = Class(I).desc
    
    
        If Class(I).gender = 0 Then
            frmNewChar.optMale.Visible = False
            frmNewChar.optFemale.Visible = False
        Else
            frmNewChar.optMale.Visible = True
            frmNewChar.optFemale.Visible = True
            frmNewChar.optMale.Caption = Class(I).gender1
            frmNewChar.optFemale.Caption = Class(I).gender2
        End If
End Sub

Private Sub HScroll1_Change()
    If HScroll1.value > MAX_HEAD Then
    HScroll1.value = 0
    End If
    If SpriteSize = 1 Then
        iconn(0).top = -Val(HScroll1.value * 64 + 15)
    Else
        iconn(0).top = -Val(HScroll1.value * PIC_Y)
    End If
End Sub

Private Sub HScroll2_Change()
    If HScroll2.value > MAX_BODY Then
    HScroll2.value = 0
    End If
    If SpriteSize = 1 Then
        iconn(1).top = -Val(HScroll2.value * 64 + 25)
    Else
        iconn(1).top = -Val(HScroll2.value * PIC_Y)
    End If
End Sub

Private Sub HScroll3_Change()
    If HScroll3.value > MAX_LEGS Then
    HScroll3.value = 0
    End If
    If SpriteSize = 1 Then
        iconn(2).top = -Val(HScroll3.value * 64 + 35)
    Else
        iconn(2).top = -Val(HScroll3.value * PIC_Y)
    End If
End Sub


Private Sub picAddChar_Click()
    Dim Msg As String
    Dim I As Long

    If Trim$(txtName.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)

        If Len(Trim$(txtName.Text)) < 3 Then
            MsgBox "El nombre del personaje debe tener más de 3 caracteres."
            Exit Sub
        End If

        ' Prevent high ascii chars
        For I = 1 To Len(Msg)
            If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 255 Then
                Call MsgBox("No puedes usar caracteres ASCII para tu nombre.", vbOKOnly, GAME_NAME)
                txtName.Text = vbNullString
                Exit Sub
            End If
        Next I

        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Me.Visible = False
End Sub


Private Sub Timer1_Timer()

    If cmbClass.ListIndex < 0 Then
        Exit Sub
    End If
    If 0 + CustomPlayers = 0 Then
        If SpriteSize = 1 Then
            If optMale.value = True Then
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.top = (Int(Class(cmbClass.ListIndex).MaleSprite) * 64) * -1
                Call BitBlt(picPic.hDC, 0, 0, PIC_X, 64, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * 64, SRCCOPY)
            Else
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.top = (Int(Class(cmbClass.ListIndex).FemaleSprite) * 64) * -1
                Call BitBlt(picPic.hDC, 0, 0, PIC_X, 64, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * 64, SRCCOPY)
            End If
        Else
            If optMale.value = True Then
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.top = (Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y) * -1

                Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
            Else
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.top = (Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y) * -1
                Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim I As Integer
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

        If FileExists("GUI\Nuevo_Personaje" & Ending) Then
            frmNewChar.Picture = LoadPicture(App.Path & "\GUI\Nuevo_Personaje" & Ending)
        End If
    Next I

    If CustomPlayers = 1 Then
        If FileExists("GFX\Sprites.bmp") Then
            picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")
        Else
            Call MsgBox("Error: Could not find Sprites.bmp.")
            End
        End If
    End If

' Set the size of the scrolling bars
' HScroll1.Max = LoadPicture(App.Path & "\GFX\Heads.bmp").Height / 64
' DOES NOT WORK, FIX LATER PLZ KTHX

End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
    If animi > 4 Then
        animi = 3
    End If
End Sub
