VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C8530F8A-C19C-11D2-99D6-9419F37DBB29}#1.1#0"; "ccrpprg6.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AlterEngine Reborn v1.8"
   ClientHeight    =   10005
   ClientLeft      =   555
   ClientTop       =   780
   ClientWidth     =   12870
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMirage.frx":A6AA
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   858
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   12000
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12480
      Top             =   0
   End
   Begin VB.Timer anunciotimer 
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   10080
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   102
      Top             =   5880
      Visible         =   0   'False
      Width           =   2400
      Begin VB.VScrollBar scrlInventory 
         Height          =   330
         Left            =   2640
         Max             =   3
         TabIndex        =   135
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picDown 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1230
         Picture         =   "frmMirage.frx":36A2F
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   134
         Top             =   3000
         Width           =   270
      End
      Begin VB.PictureBox picUp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         Picture         =   "frmMirage.frx":36CBA
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   133
         Top             =   3000
         Width           =   270
      End
      Begin VB.PictureBox picInventory2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2670
         Left            =   0
         ScaleHeight     =   178
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   103
         Top             =   0
         Width           =   2400
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   24
            Left            =   1815
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   132
            Top             =   2670
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   25
            Left            =   1245
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   131
            Top             =   2670
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   26
            Left            =   675
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   130
            Top             =   2670
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   27
            Left            =   105
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   129
            Top             =   2670
            Width           =   480
         End
         Begin VB.PictureBox picInventory3 
            Appearance      =   0  'Flat
            BackColor       =   &H00828B82&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   5700
            Left            =   0
            ScaleHeight     =   380
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   160
            TabIndex        =   104
            Top             =   0
            Width           =   2400
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   23
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   128
               Top             =   2670
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   22
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   127
               Top             =   2670
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   21
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   126
               Top             =   2670
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   20
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   125
               Top             =   2670
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   19
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   124
               Top             =   2145
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   18
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   123
               Top             =   2145
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   17
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   122
               Top             =   2145
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   16
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   121
               Top             =   2145
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   15
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   120
               Top             =   1620
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   14
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   119
               Top             =   1620
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   13
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   118
               Top             =   1620
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   12
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   117
               Top             =   1620
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   11
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   116
               Top             =   1095
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   10
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   115
               Top             =   1095
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   9
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   114
               Top             =   1095
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   113
               Top             =   45
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   112
               Top             =   45
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   2
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   111
               Top             =   45
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   3
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   110
               Top             =   45
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   4
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   109
               Top             =   570
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   108
               Top             =   570
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   6
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   107
               Top             =   570
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   7
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   106
               Top             =   570
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   8
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   105
               Top             =   1095
               Width           =   480
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   0
               Left            =   600
               Top             =   720
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   1
               Left            =   0
               Top             =   720
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   2
               Left            =   600
               Top             =   720
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   3
               Left            =   600
               Top             =   600
               Width           =   540
            End
            Begin VB.Shape SelectedItem 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   525
               Left            =   90
               Top             =   30
               Width           =   525
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   6
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   5
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   4
               Left            =   0
               Top             =   0
               Width           =   540
            End
         End
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Usar objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   137
         Top             =   3000
         Width           =   810
      End
      Begin VB.Label lblDropItem 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Soltar objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1560
         TabIndex        =   136
         Top             =   3000
         Width           =   795
      End
      Begin VB.Line Line1 
         X1              =   40
         X2              =   128
         Y1              =   192
         Y2              =   192
      End
   End
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00828B82&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   240
      ScaleHeight     =   279
      ScaleMode       =   0  'User
      ScaleWidth      =   319
      TabIndex        =   49
      Top             =   2880
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label descMagic 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Magic"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -240
         TabIndex        =   100
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   60
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label lblRequirements 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Requerimientos"
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
         Height          =   210
         Left            =   -240
         TabIndex        =   59
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -240
         TabIndex        =   58
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -240
         TabIndex        =   57
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   -240
         TabIndex        =   56
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label lblAdditions 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficios Adicionales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2040
         TabIndex        =   55
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2040
         TabIndex        =   54
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2040
         TabIndex        =   53
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   840
         TabIndex        =   52
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Height          =   210
         Left            =   1200
         TabIndex        =   51
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: XXXX Magic: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2040
         TabIndex        =   50
         Top             =   1200
         Width           =   2655
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4320
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   61
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CheckBox chkPlayerBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mini barra de PV"
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
         Height          =   225
         Left            =   120
         TabIndex        =   72
         Top             =   840
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nombres"
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
         Height          =   225
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkNpcName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nombres"
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
         Height          =   225
         Left            =   2760
         TabIndex        =   70
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkBubbleBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Texto en burbujas"
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
         Height          =   225
         Left            =   2760
         TabIndex        =   69
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkNpcBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Barras de PV"
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
         Height          =   225
         Left            =   2760
         TabIndex        =   68
         Top             =   840
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkPlayerDamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Daño encima de la cabeza"
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
         Height          =   225
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkNpcDamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Daño encima de la cabeza"
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
         Height          =   225
         Left            =   2760
         TabIndex        =   66
         Top             =   600
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkMusic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Musica"
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
         Height          =   225
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkSound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sonidos"
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
         Height          =   225
         Left            =   120
         TabIndex        =   64
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   4
         TabIndex        =   63
         Top             =   2280
         Value           =   6
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto Scroll"
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
         Height          =   225
         Left            =   2760
         TabIndex        =   62
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin Eclipse.jcbutton cmdSaveConfig 
         Height          =   495
         Left            =   360
         TabIndex        =   159
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         Caption         =   "Guardar"
         PictureNormal   =   "frmMirage.frx":36F52
         PictureHot      =   "frmMirage.frx":37736
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Eclipse.jcbutton cnfigcontrols 
         Height          =   735
         Left            =   2760
         TabIndex        =   160
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1296
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
         Caption         =   "Configurar Controles"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   4
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin VB.Label lblNPCData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración de NPCs"
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
         Height          =   210
         Left            =   2760
         TabIndex        =   101
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblLines 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lineas de texto en pantalla: 6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   480
         TabIndex        =   76
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblPlayerData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Información del Jugador"
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
         Height          =   210
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblSoundData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración de sonido"
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
         Height          =   210
         Left            =   120
         TabIndex        =   74
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblChatData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración de Chat"
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
         Height          =   210
         Left            =   2760
         TabIndex        =   73
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   105
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   48
      Top             =   675
      Width           =   9600
      Begin VB.PictureBox PicExtMen 
         Height          =   2415
         Left            =   5400
         Picture         =   "frmMirage.frx":37F1A
         ScaleHeight     =   2355
         ScaleWidth      =   1995
         TabIndex        =   161
         Top             =   3480
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Label Cancel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   165
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Exit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salir"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   164
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label MPrinc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "           Menu          Principal"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   163
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label SPer 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Seleccionar Personaje"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   162
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.PictureBox anunciobox 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   480
         ScaleHeight     =   825
         ScaleWidth      =   8625
         TabIndex        =   153
         Top             =   120
         Visible         =   0   'False
         Width           =   8655
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Anuncio:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3600
            TabIndex        =   156
            Top             =   0
            Width           =   975
         End
         Begin VB.Label anuncio 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   154
            Top             =   240
            Visible         =   0   'False
            Width           =   8415
         End
      End
      Begin VB.PictureBox barracasteo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         Picture         =   "frmMirage.frx":3B553
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   235
         TabIndex        =   157
         Top             =   120
         Visible         =   0   'False
         Width           =   3555
         Begin CCRProgressBar6.ccrpProgressBar ccrpProgressBar1 
            Height          =   390
            Left            =   135
            Top             =   165
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   688
            BackPicture     =   "frmMirage.frx":3F0B1
            Caption         =   " "
            FillColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IncrementSize   =   1
            Picture         =   "frmMirage.frx":41498
            Smooth          =   -1  'True
            Style           =   1
         End
         Begin VB.Label cancelcasteo 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   3000
            TabIndex        =   158
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.PictureBox mostrarbarra2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9360
         Picture         =   "frmMirage.frx":446C7
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   152
         Top             =   6960
         Width           =   255
      End
      Begin VB.PictureBox mostrarbarra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9360
         Picture         =   "frmMirage.frx":4684F
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   151
         Top             =   6960
         Width           =   255
      End
      Begin VB.PictureBox barrahechizos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3120
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   407
         TabIndex        =   150
         Top             =   6360
         Width           =   6135
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   9
            Left            =   5520
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   8
            Left            =   4920
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   7
            Left            =   4320
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   6
            Left            =   3720
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   5
            Left            =   3120
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   4
            Left            =   2520
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   3
            Left            =   1920
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   2
            Left            =   1320
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   1
            Left            =   720
            Top             =   120
            Width           =   495
         End
         Begin VB.Image Imagesb 
            Appearance      =   0  'Flat
            Height          =   495
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12960
      ScaleHeight     =   477.09
      ScaleMode       =   0  'User
      ScaleWidth      =   477.091
      TabIndex        =   47
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12960
      Top             =   2640
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12960
      Top             =   2160
   End
   Begin VB.Timer tmrGameClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12960
      Top             =   1680
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12960
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   29
      Top             =   4320
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   9840
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   2865
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2835
         ItemData        =   "frmMirage.frx":489D7
         Left            =   240
         List            =   "frmMirage.frx":489D9
         TabIndex        =   2
         Top             =   120
         Width           =   2310
      End
      Begin VB.Label lblForgetSpell 
         BackStyle       =   0  'Transparent
         Caption         =   "Olvidar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   43
         Top             =   3240
         Width           =   495
      End
   End
   Begin VB.PictureBox picWhosOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   9840
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   2880
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         ItemData        =   "frmMirage.frx":489DB
         Left            =   240
         List            =   "frmMirage.frx":489DD
         TabIndex        =   10
         Top             =   360
         Width           =   2430
      End
   End
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   9840
      ScaleHeight     =   3705
      ScaleWidth      =   2880
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   2880
      Begin VB.TextBox txtAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   18
         Top             =   585
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   17
         Top             =   345
         Width           =   1575
      End
      Begin VB.CommandButton cmdTrainee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hacer Aprendiz"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdMember 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hacer Miembro"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmdDisown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expulsar"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambiar Privilegio"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Privilegio:"
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
         Left            =   180
         TabIndex        =   20
         Top             =   615
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
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
         Left            =   210
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.PictureBox picEquipment 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   9840
      ScaleHeight     =   3705
      ScaleWidth      =   2880
      TabIndex        =   28
      Top             =   5760
      Width           =   2880
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   44
         Top             =   2280
         Width           =   555
         Begin VB.PictureBox LegsImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   45
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox AmuletImage2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   40
         Top             =   840
         Width           =   555
         Begin VB.PictureBox NecklaceImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   41
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1680
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   38
         Top             =   1560
         Width           =   555
         Begin VB.PictureBox RingImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   39
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   240
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   36
         Top             =   840
         Width           =   555
         Begin VB.PictureBox WeaponImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   37
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   34
         Top             =   1560
         Width           =   555
         Begin VB.PictureBox ArmorImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   35
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1680
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   32
         Top             =   840
         Width           =   555
         Begin VB.PictureBox ShieldImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   33
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   30
         Top             =   120
         Width           =   555
         Begin VB.PictureBox HelmetImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   31
            Top             =   15
            Width           =   495
         End
      End
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   240
      MaxLength       =   200
      TabIndex        =   7
      Top             =   8040
      Visible         =   0   'False
      Width           =   9360
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   12960
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   2434
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":489DF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picCharStatus 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   9840
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   80
      Top             =   5760
      Width           =   2880
      Begin VB.Label AddDEF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2520
         TabIndex        =   99
         Top             =   2400
         Width           =   165
      End
      Begin VB.Label AddSTR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2520
         TabIndex        =   98
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   97
         Top             =   2400
         Width           =   1170
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STRENGTH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   96
         Top             =   2040
         Width           =   1170
      End
      Begin VB.Label AddMAGI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2520
         TabIndex        =   95
         Top             =   1320
         Width           =   165
      End
      Begin VB.Label AddSPD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2520
         TabIndex        =   94
         Top             =   1680
         Width           =   165
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   93
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAGIC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   92
         Top             =   1320
         Width           =   1170
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1080
         TabIndex        =   91
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label lblSTATWINDOW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- Personaje -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "POINTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   1320
         TabIndex        =   89
         Top             =   2760
         Width           =   1170
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   88
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Magia :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   87
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   86
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fuerza :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   85
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defensa :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   84
         Top             =   2400
         Width           =   1125
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PUNTOS :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   83
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Energia :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   82
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   81
         Top             =   960
         Width           =   1125
      End
   End
   Begin VB.PictureBox picGuildMember 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   9840
      ScaleHeight     =   3705
      ScaleWidth      =   2880
      TabIndex        =   22
      Top             =   5760
      Visible         =   0   'False
      Width           =   2880
      Begin VB.Label cmdLeave 
         BackStyle       =   0  'Transparent
         Caption         =   "Salir del clan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   27
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblGuildRank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
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
         Left            =   1440
         TabIndex        =   26
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lblGuildName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guild"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tu Rango :"
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
         Height          =   165
         Left            =   570
         TabIndex        =   24
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Clan :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1140
      End
   End
   Begin VB.Label anunciocuenta 
      Height          =   495
      Left            =   1560
      TabIndex        =   155
      Top             =   10200
      Width           =   2895
   End
   Begin VB.Image Image16 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   10
      Left            =   11040
      TabIndex        =   149
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   9
      Left            =   10080
      TabIndex        =   148
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   8
      Left            =   9120
      TabIndex        =   147
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   7
      Left            =   8160
      TabIndex        =   146
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   145
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   144
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   4
      Left            =   5280
      TabIndex        =   143
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   142
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   141
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   140
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   139
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label mascota 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   10440
      TabIndex        =   138
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label lblEquipment 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
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
      Height          =   420
      Left            =   10080
      TabIndex        =   79
      Top             =   2400
      Width           =   2520
   End
   Begin VB.Label lblCharStats 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   10080
      TabIndex        =   78
      Top             =   1440
      Width           =   2520
   End
   Begin VB.Label lblMenuQuit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   10080
      TabIndex        =   77
      Top             =   4560
      Width           =   2520
   End
   Begin VB.Label lblGameClock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(tiempo)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   46
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   7155
      TabIndex        =   42
      Top             =   210
      Width           =   2250
   End
   Begin VB.Label lblGuild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   10080
      TabIndex        =   21
      Top             =   3480
      Width           =   2520
   End
   Begin VB.Label lblWhosOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   10080
      TabIndex        =   11
      Top             =   3960
      Width           =   2520
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   10080
      TabIndex        =   8
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label lblSpells 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   10080
      TabIndex        =   6
      Top             =   3000
      Width           =   2400
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   10080
      TabIndex        =   5
      Top             =   1920
      Width           =   2520
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB884B&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4200
      TabIndex        =   4
      Top             =   210
      Width           =   2250
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1530
      TabIndex        =   3
      Top             =   210
      Width           =   2250
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   1530
      Top             =   210
      Width           =   2250
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00CB884B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   4275
      Top             =   210
      Width           =   2250
   End
   Begin VB.Shape shpTNL 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      Height          =   225
      Left            =   7155
      Top             =   210
      Width           =   2250
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub AddStr_Click()
    Call SendUseStatPoint(0)
End Sub

Private Sub AddDef_Click()
    Call SendUseStatPoint(1)
End Sub

Private Sub AddMagi_Click()
    Call SendUseStatPoint(2)
End Sub

Private Sub AddSPD_Click()
    Call SendUseStatPoint(3)
End Sub

Private Sub anunciotimer_Timer()
     If anunciocuenta.Caption = 0 Then
        anunciotimer.Enabled = False
        anunciobox.Visible = False
        anuncio.Visible = False
     Else
          anunciocuenta.Caption = anunciocuenta.Caption - 1
     End If
End Sub

Private Sub barracasteo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub barracasteo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(frmMirage.barracasteo, Button, Shift, X, Y)
End Sub
Private Sub barrahechizos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub barrahechizos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(frmMirage.barrahechizos, Button, Shift, X, Y)
End Sub

Private Sub Cancel_Click()
PicExtMen.Visible = False
End Sub

Private Sub cancelcasteo_Click()
Timer1.Interval = 0
Timer1.Enabled = False
ccrpProgressBar1.min = 0
ccrpProgressBar1.value = 0
ccrpProgressBar1.Clear
frmMirage.barracasteo.Visible = False
End Sub

Private Sub chkSound_Click()
    Call WriteINI("CONFIG", "Sound", chkSound.value, App.Path & "\Config.ini")
End Sub

Private Sub chkBubbleBar_Click()
    Call WriteINI("CONFIG", "SpeechBubbles", chkBubbleBar.value, App.Path & "\Config.ini")
End Sub

Private Sub chkNpcBar_Click()
    Call WriteINI("CONFIG", "NPCBar", chkNpcBar.value, App.Path & "\Config.ini")
End Sub

Private Sub chkNpcDamage_Click()
    Call WriteINI("CONFIG", "NPCDamage", chkNpcDamage.value, App.Path & "\Config.ini")
End Sub

Private Sub chkNpcName_Click()
    Call WriteINI("CONFIG", "NPCName", chkNpcName.value, App.Path & "\Config.ini")
End Sub

Private Sub chkPlayerBar_Click()
    Call WriteINI("CONFIG", "PlayerBar", chkPlayerBar.value, App.Path & "\Config.ini")
End Sub

Private Sub chkPlayerDamage_Click()
    Call WriteINI("CONFIG", "PlayerDamage", chkPlayerDamage.value, App.Path & "\Config.ini")
End Sub

Private Sub chkAutoScroll_Click()
    Call WriteINI("CONFIG", "AutoScroll", chkAutoScroll.value, App.Path & "\Config.ini")
End Sub

Private Sub chkPlayerName_Click()
    Call WriteINI("CONFIG", "PlayerName", chkPlayerName.value, App.Path & "\Config.ini")
End Sub

Private Sub chkMusic_Click()
    If chkMusic = Checked Then
        Call WriteINI("CONFIG", "Music", 1, App.Path & "\Config.ini")
        Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
    Else
        Call WriteINI("CONFIG", "Music", 0, App.Path & "\Config.ini")
        Call StopBGM
    End If
End Sub

Private Sub cmdLeave_Click()
    Call SendGuildLeave
End Sub

Private Sub cmdMember_Click()
    Call SendGuildMember(txtName.Text)
End Sub

Private Sub cmdSaveConfig_Click()
    picOptions.Visible = False
End Sub

Private Sub cnfigcontrols_Click()
If FileExists("/ConfiguradorAE.exe") Then
    Call Shell(App.Path & "/ConfiguradorAE.exe", vbNormalFocus)
Else
    Call MsgBox("Error, No se encuentra el Archivo ConfiguradorAE", vbCritical)
End If

End Sub

Private Sub Exit_Click()
    Call TcpDestroy
    Call DestroyDirectX
    Call BASS_Free
    PicExtMen.Visible = False
    End
End Sub

Private Sub Form_Load()
Dim result As Long
    result = SetWindowLong(txtChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Dim I As Long
    Dim Ending As String

    For I = 1 To 4
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
        If I = 4 Then Ending = ".bmp"

        If FileExists("GUI\Interfaz_Juego" & Ending) Then
            frmMirage.Picture = LoadPicture(App.Path & "\GUI\Interfaz_Juego" & Ending)
        End If
    Next I
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Terminate()
        Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call GameDestroy
End Sub

Private Sub Imagesb_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Byte
If Button = 2 Then
        Imagesb(index).Picture = Image16.Picture
        Call WriteINI("SK" & (index + 1), "sid", "0", App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini")
        Call WriteINI("SKP" & SpellMemorized, "sid", "0", App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini")
        For I = 1 To MAX_PLAYER_SPELLS
            If (I - 1) = index Then
                Player(MyIndex).Spellpos(I) = 0
                Exit Sub
            End If
        Next I
ElseIf Button = 1 Then
    If Imagesb(index).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = index + 1

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
                Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
                Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
    End If
End If
End Sub

Private Sub Label1_Click()
If IsPlaying(MyIndex) = True Then
If Player(MyIndex).Pet.Alive = YES Then
frmPetMenu.Visible = True
End If
End If
End Sub



Private Sub lblOptions_Click()
    chkPlayerName.value = Trim$(ReadINI("CONFIG", "PlayerName", App.Path & "\Config.ini"))
    chkPlayerDamage.value = Trim$(ReadINI("CONFIG", "PlayerDamage", App.Path & "\Config.ini"))
    chkPlayerBar.value = Trim$(ReadINI("CONFIG", "PlayerBar", App.Path & "\Config.ini"))
    chkNpcName.value = Trim$(ReadINI("CONFIG", "NPCName", App.Path & "\Config.ini"))
    chkNpcDamage.value = Trim$(ReadINI("CONFIG", "NPCDamage", App.Path & "\Config.ini"))
    chkNpcBar.value = Trim$(ReadINI("CONFIG", "NPCBar", App.Path & "\Config.ini"))
    chkMusic.value = Trim$(ReadINI("CONFIG", "Music", App.Path & "\Config.ini"))
    chkSound.value = Trim$(ReadINI("CONFIG", "Sound", App.Path & "\Config.ini"))
    chkBubbleBar.value = Trim$(ReadINI("CONFIG", "SpeechBubbles", App.Path & "\Config.ini"))
    chkAutoScroll.value = Trim$(ReadINI("CONFIG", "AutoScroll", App.Path & "\Config.ini"))

    picOptions.Visible = True
End Sub

Private Sub lblGuild_Click()
    If LenB(GetPlayerGuild(MyIndex)) <> 0 Then
        lblGuildName.Caption = GetPlayerGuild(MyIndex)
        lblGuildRank.Caption = GetPlayerGuildAccess(MyIndex)
    Else
        lblGuildName.Caption = "Ninguno"
        lblGuildRank.Caption = "Ninguno"
    End If

    picInventory.Visible = False
    picPlayerSpells.Visible = False
    picEquipment.Visible = False
    picWhosOnline.Visible = False
    picCharStatus.Visible = False
    picGuildAdmin.Visible = False
    picGuildMember.Visible = True
End Sub

Private Sub lblEquipment_Click()
    Call UpdateVisInv

    picInventory.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picCharStatus.Visible = False
    picEquipment.Visible = True
End Sub

Private Sub lblInventory_Click()
    Call UpdateVisInv

    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picEquipment.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    picCharStatus.Visible = False
    picInventory.Visible = True
End Sub

Private Sub lblSpells_Click()
    Call SendRequestSpells

    picInventory.Visible = False
    picGuildAdmin.Visible = False
    picEquipment.Visible = False
    picGuildMember.Visible = False
    picWhosOnline.Visible = False
    picCharStatus.Visible = False
    picPlayerSpells.Visible = True
End Sub

Private Sub lblCharStats_Click()
    picWhosOnline.Visible = False
    picInventory.Visible = False
    picEquipment.Visible = False
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picPlayerSpells.Visible = False
    picCharStatus.Visible = True
End Sub

Private Sub lblForgetSpell_Click()
    Call SendForgetSpell(lstSpells.ListIndex + 1)
End Sub

Private Sub lblMenuQuit_Click()
    PicExtMen.Visible = True
    'frmMirage.Visible = False
    'frmSendGetData.Visible = False
    'frmMainMenu.Visible = True
    'Call TcpDestroy
    'Call DestroyDirectX
End Sub

Private Sub lblSTATWINDOW_Click()
    Call SendRequestMyStats
End Sub

Private Sub lblWhosOnline_Click()
    Call SendOnlineList

    picInventory.Visible = False
    picEquipment.Visible = False
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picPlayerSpells.Visible = False
    picCharStatus.Visible = False
    picWhosOnline.Visible = True
End Sub

Private Sub lstOnline_DblClick()
    'Call SendPlayerChat(Trim$(lstOnline.Text))
    Call SendProfile(Trim$(frmMirage.lstOnline.Text))
End Sub

Private Sub lstSpells_DblClick()
Dim O As String, R() As String, p As String, q As String
Dim I
Dim X
Dim exists As Boolean
Dim file As String

O = lstSpells.Text
R = Split(O, ":")
    If (StrComp(Left(O, 1), "-") <> 0) Then 'Agregar la Imagen al 1° cuadro vacio de la Spellbar y
                                            'guardar el ID dentro del .ini
        file = "\GUI\Hechizos\" & Trim(R(1)) & ".gif"
        
        For I = 1 To MAX_PLAYER_SPELLS
            If (Player(MyIndex).Spellpos(I) <> 0) Then
                If Trim(Spell(Player(MyIndex).Spellpos(I)).name) = Trim(R(1)) Then
                    Call AddText("Ya tienes este hechizo en tu barra!", Red)
                    Exit Sub
                End If
            End If
        Next I
        For I = 1 To 10
        If (Imagesb(I - 1).Picture = Image16.Picture) Then
            If FileExists(file) Then
            Imagesb(I - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & Trim(R(1)) & ".gif")
            Else
            Imagesb(I - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\X.gif")
            End If

            Call WriteINI("SK" & I, "sid", "" & Player(MyIndex).Spell(CInt(R(0))), App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini")
            Call WriteINI("SKP" & I, "sid", "" & CInt(R(0)), App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini")
            Player(MyIndex).Spellpos(I) = Player(MyIndex).Spell(CInt(R(0)))
            Exit Sub
        Else
            If I = 10 Then
                Call AddText("Todos tus huecos de hechizos están llenos.", Red)
            End If
        End If
        Next I
    End If
End Sub

Private Sub mascota_Click()
If IsPlaying(MyIndex) = True Then
If Player(MyIndex).Pet.Alive = YES Then
frmPetMenu.Visible = True
End If
End If
End Sub


Private Sub mostrarbarra_Click()
barrahechizos.Visible = False
mostrarbarra.Visible = False
mostrarbarra2.Visible = True

End Sub

Private Sub mostrarbarra2_Click()
barrahechizos.Visible = True
mostrarbarra.Visible = True
mostrarbarra2.Visible = False
End Sub

Private Sub MPrinc_Click()
    frmMirage.Visible = False
    frmSendGetData.Visible = False
    Call TcpDestroy
    Call DestroyDirectX
    Call BASS_Free
    Call Main
    PicExtMen.Visible = False
End Sub

Private Sub picInv_DblClick(index As Integer)
    Dim d As Long

    If Player(MyIndex).Inv(Inventory).Num <= 0 Then
        Exit Sub
    End If

    Call SendUseItem(Inventory)

    For d = 1 To MAX_INV
        If Player(MyIndex).Inv(d).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
                picInv(d - 1).Picture = LoadPicture()
            End If
        End If
    Next d
    Call UpdateVisInv
End Sub

Private Sub picInv_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim AMT As Integer

    Inventory = index + 1

    SelectedItem.top = picInv(Inventory - 1).top - 1
    SelectedItem.Left = picInv(Inventory - 1).Left - 1

    If frmNewShop.fixItems And frmNewShop.Visible = True Then
        frmNewShop.FixItem (GetPlayerInvItemNum(MyIndex, Inventory))
    Else
        ' We're selling items to a shop
        If frmNewShop.SellItems And frmNewShop.Visible = True Then
            If Item(GetPlayerInvItemNum(MyIndex, Inventory)).Stackable = YES Then
                AMT = Val(InputBox("¿Cuanta cantidad deseas vender?", "Vender Objetos")) + 0
                If AMT > 0 Then
                    ' Sell the items
                    frmNewShop.Buyback GetPlayerInvItemNum(MyIndex, Inventory), Inventory, AMT
                End If
            Else
                ' Sell the selected item
                frmNewShop.Buyback GetPlayerInvItemNum(MyIndex, Inventory), Inventory
            End If
        Else
            ' Regular click
            If Button = 1 Then
                Call UpdateVisInv
            ElseIf Button = 2 Then
                Call DropItem
            End If
        End If
    End If
End Sub

Private Sub picInv_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim d As Long
    d = index

    If Player(MyIndex).Inv(d + 1).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            If Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
                itmDesc.top = 390
                itmDesc.Height = 17
            Else
                itmDesc.top = 240
                itmDesc.Height = 251
            End If
        Else
            If Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
                itmDesc.top = 240
                itmDesc.Height = 209
            Else
                itmDesc.top = 240
                itmDesc.Height = 251
            End If
        End If
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            ElseIf GetPlayerNecklaceSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (Equipado)"
            Else
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name)
            End If
        End If

' Fix aplicado por Stream
        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " " & STAT1
        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " " & STAT2
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " " & STAT3
        descMagic.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).MagicReq & " " & STAT4
        descHpMp.Caption = "HP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = STAT1 & ": " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " " & STAT2 & ": " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = STAT3 & ": " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed & " " & STAT4 & ": " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi
        desc.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        

        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ScreenID As Long
    Dim I As Long

    Call CheckInput(0, KeyCode, Shift)

    If KeyCode = vbKeyF1 Then
        If Player(MyIndex).Access > 0 Then
            frmAdmin.Visible = False
            frmAdmin.Visible = True
        End If
    End If

    If KeyCode = vbKeyF2 Then
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
                If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_POTIONADDHP Then
                    Call AddText("Has usado una poción para restaurar tu vida.", YELLOW)
                    Call SendUseItem(I)
                    Exit Sub
                End If
            Else
                If I = MAX_INV Then
                    Call AddText("No tienes pociones para restaurar tu vida!", BRIGHTRED)
                End If
            End If
        Next I
    End If

    If KeyCode = vbKeyF3 Then
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
                If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_POTIONADDMP Then
                    Call AddText("Has usado una poción para restaurar el mana.", YELLOW)
                    Call SendUseItem(I)
                    Exit Sub
                End If
            Else
                If I = MAX_INV Then
                    Call AddText("No tienes pociones para restaurar tu mana!", BRIGHTRED)
                End If
            End If
        Next I
    End If

    If KeyCode = vbKeyF4 Then
        If Player(MyIndex).Access >= 3 Then
            frmGuild.Show vbModeless, frmMirage
        Else
         Call AddText("No tienes permisos para Utilizar esta opción", Red)
        End If
    End If

    If KeyCode = vbKeyF5 Then
        picInventory.Visible = False
        picGuildMember.Visible = False
        picEquipment.Visible = False
        picPlayerSpells.Visible = False
        picWhosOnline.Visible = False
        picGuildAdmin.Visible = True
    End If

    If KeyCode = vbKeyInsert Then
       ' If SpellMemorized > 0 Then
        '    If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
         '       If Player(MyIndex).Moving = 0 Then
          '          Call SendData("cast" & SEP_CHAR & SpellMemorized & END_CHAR)

           '         Player(MyIndex).Attacking = 1
            '        Player(MyIndex).AttackTimer = GetTickCount
             '       Player(MyIndex).CastedSpell = YES
              '  Else
               '     Call AddText("No puedes lanzar un hechizo cuando caminas!", BRIGHTRED)
                'End If
            'End If
        'Else
            'Call AddText("No tienes ningún hechizo memorizado aquí.", BRIGHTRED)
        'End If
    End If

    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)

        Do
            If FileExists("Screenshot_" & ScreenID & ".bmp") Then
                ScreenID = ScreenID + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot_" & ScreenID & ".bmp")
                Exit Do
            End If
        Loop
    End If

    If KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, picScreen.Left, picScreen.top, picScreen.Width, picScreen.Height)

        Do
            If FileExists("Screenshot_" & ScreenID & ".bmp") Then
                ScreenID = ScreenID + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot_" & ScreenID & ".bmp")
                Exit Do
            End If
        Loop
    End If

    If KeyCode = vbKeyPageUp Then
        Call SendHotScript(1)
    End If

    If KeyCode = vbKeyDelete Then
        Call SendHotScript(2)
    End If

    If KeyCode = vbKeyEnd Then
        Call SendHotScript(3)
    End If

    If KeyCode = vbKeyPageDown Then
        Call SendHotScript(4)
    End If
    
    If KeyCode = vbKeyA Then
        Call SendHotScript(5)
    End If
    If KeyCode = vbKeyB Then
        Call SendHotScript(6)
    End If
    If KeyCode = vbKeyC Then
        Call SendHotScript(7)
    End If
    If KeyCode = vbKeyD Then
        Call SendHotScript(8)
    End If
    If KeyCode = vbKeyE Then
        Call SendHotScript(9)
    End If
    If KeyCode = vbKeyF Then
        Call SendHotScript(10)
    End If
    If KeyCode = vbKeyG Then
        Call SendHotScript(11)
    End If
    If KeyCode = vbKeyH Then
        Call SendHotScript(12)
    End If
    If KeyCode = vbKeyI Then
        Call SendHotScript(13)
    End If
    If KeyCode = vbKeyJ Then
        Call SendHotScript(14)
    End If
    If KeyCode = vbKeyK Then
        Call SendHotScript(15)
    End If
    If KeyCode = vbKeyL Then
        Call SendHotScript(16)
    End If
    If KeyCode = vbKeyM Then
        Call SendHotScript(17)
    End If
    If KeyCode = vbKeyN Then
        Call SendHotScript(18)
    End If
    If KeyCode = vbKeyO Then
        Call SendHotScript(19)
    End If
    If KeyCode = vbKeyP Then
        Call SendHotScript(20)
    End If
    If KeyCode = vbKeyQ Then
        Call SendHotScript(21)
    End If
    If KeyCode = vbKeyR Then
        Call SendHotScript(22)
    End If
    If KeyCode = vbKeyS Then
        Call SendHotScript(23)
    End If
    If KeyCode = vbKeyT Then
        Call SendHotScript(24)
    End If
    If KeyCode = vbKeyU Then
        Call SendHotScript(25)
    End If
    If KeyCode = vbKeyV Then
        Call SendHotScript(26)
    End If
    If KeyCode = vbKeyW Then
        Call SendHotScript(27)
    End If
    If KeyCode = vbKeyX Then
        Call SendHotScript(28)
    End If
    If KeyCode = vbKeyY Then
        Call SendHotScript(29)
    End If
    If KeyCode = vbKeyZ Then
        Call SendHotScript(30)
    End If


    If KeyCode = vbKeyHome Then
        If Player(MyIndex).Moving = NO Then
            If Player(MyIndex).Dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
            ElseIf Player(MyIndex).Dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
            ElseIf Player(MyIndex).Dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
            ElseIf Player(MyIndex).Dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
            End If

            Call SendPlayerDir
        End If
    End If
    
    
    If KeyCode = vbKey0 And frmMirage.txtMyTextBox.Visible = False Then

        If Imagesb(9).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 10

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
   End If
    If KeyCode = vbKey1 And frmMirage.txtMyTextBox.Visible = False Then

        If Imagesb(0).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 1

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey2 And frmMirage.txtMyTextBox.Visible = False Then
   

        If Imagesb(1).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 2

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey3 And frmMirage.txtMyTextBox.Visible = False Then

        If Imagesb(2).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 3

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey4 And frmMirage.txtMyTextBox.Visible = False Then

        If Imagesb(3).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 4

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey5 And frmMirage.txtMyTextBox.Visible = False Then

        If Imagesb(4).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 5

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey6 And frmMirage.txtMyTextBox.Visible = False Then

        If Imagesb(5).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 6

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey7 And frmMirage.txtMyTextBox.Visible = False Then
    
        If Imagesb(6).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 7

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, SpellMemorized)
            Call BltSpellsBar(MyIndex, SpellMemorized)
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
    End If
    If KeyCode = vbKey8 And frmMirage.txtMyTextBox.Visible = False Then
        If Imagesb(7).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 8

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If
 
    End If
    If KeyCode = vbKey9 And frmMirage.txtMyTextBox.Visible = False Then
        If Imagesb(8).Picture <> Image16.Picture Then
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = 9

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
        End If

    End If
End Sub

Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(frmMirage.picOptions, Button, Shift, X, Y)
End Sub

Private Sub picScreen_GotFocus()
    On Error Resume Next

    txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long

    If (Button = 1 Or Button = 2) And InEditor Then
        Call EditorMouseDown(Button, Shift, CurX, CurY)
    End If

    If Button = 1 And Not InEditor Then
        Call PlayerSearch(Button, Shift, CurX, CurY)
    End If
    
    If Button = 1 And Player(MyIndex).Pet.Alive = YES Then
            Call PetMove(Button, Shift, (X + (NewPlayerX * PIC_X)), (Y + (NewPlayerY * PIC_Y)))
    End If
    
    If Shift = 1 And Not InEditor Then
        If GetPlayerAccess(MyIndex) > 0 Then
            Call LocalWarp(CurX, CurY)
        End If
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurX = Int((X + (NewPlayerX * PIC_X)) / PIC_X)
    CurY = Int((Y + (NewPlayerY * PIC_Y)) / PIC_Y)

    If (Button = 1 Or Button = 2) And InEditor Then
        Call EditorMouseDown(Button, Shift, CurX, CurY)
    End If

    frmMapEditor.Caption = "AlterEngine | Editor de mapas - " & "X: " & CurX & " Y: " & CurY
End Sub

Private Sub picInventory3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub
Private Sub scrlBltText_Change()
    Dim I As Long

    For I = 1 To MAX_BLT_LINE
        BattlePMsg(I).index = 1
        BattlePMsg(I).time = I
        BattleMMsg(I).index = 1
        BattleMMsg(I).time = I
    Next I

    MAX_BLT_LINE = scrlBltText.value

    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec

    lblLines.Caption = "Lineas de texto en pantalla: " & scrlBltText.value
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)

    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub SPer_Click()
    frmMirage.Visible = False
    frmSendGetData.Visible = False
    frmChars.Visible = True
    'Call DestroyDirectX
    Call BASS_Free
    Call ClearPlayer(MyIndex)
    Call SendData("ALLCHARS" & SEP_CHAR & MyIndex & END_CHAR)
    Call Sleep(1)
    frmChars.Visible = True
    PicExtMen.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim z As Variant
Dim PosID As Long

        If SCast = 0 Then
            Call LoadSCT(MyIndex, SpellMemorized)
            If Spell(Player(MyIndex).Spellpos(SpellMemorized)).MPCost > Player(MyIndex).MP Then
                Timer1.Interval = 0
                Timer1.Enabled = False
                ccrpProgressBar1.min = 0
                ccrpProgressBar1.value = 0
                ccrpProgressBar1.Clear
                frmMirage.barracasteo.Visible = False
                Call AddText("No Tienes Suficiente MP", BRIGHTRED)
                Exit Sub
            End If
        
            If GetTickCount < Player(MyIndex).SpellcdTimer(SpellMemorized) + Spell(Player(MyIndex).Spellpos(SpellMemorized)).TimeToCast * 1000 Then
                Timer1.Interval = 0
                Timer1.Enabled = False
                ccrpProgressBar1.min = 0
                ccrpProgressBar1.value = 0
                ccrpProgressBar1.Clear
                frmMirage.barracasteo.Visible = False
                Call AddText("El Hechizo Esta en Cooldown", BRIGHTRED)
                Exit Sub
            End If
        
            frmMirage.barracasteo.Visible = True
            If Spell(Player(MyIndex).Spellpos(SpellMemorized)).CastTimer <= 100 Then
                SCast = 0
                Call SendData("cast" & SEP_CHAR & ReadINI("SKP" & SpellMemorized, "sid", App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini") & END_CHAR)
                Player(MyIndex).SpellcdTimer(SpellMemorized) = GetTickCount
                Player(MyIndex).spellcdb(SpellMemorized) = True
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
                Timer1.Interval = 0
                Timer1.Enabled = False
                Exit Sub
            Else
                Player(MyIndex).AttackTimer = GetTickCount
                frmMirage.ccrpProgressBar1.max = GetTickCount + Spell(Player(MyIndex).Spellpos(SpellMemorized)).CastTimer
                frmMirage.ccrpProgressBar1.min = GetTickCount
                SCast = Spell(Player(MyIndex).Spellpos(SpellMemorized)).CastTimer + Player(MyIndex).AttackTimer
            End If
        End If
            z = GetTickCount
            'z = z / SCast
    
            If z < frmMirage.ccrpProgressBar1.max Then
                ccrpProgressBar1.value = z
            Else
                SCast = 0
                Call SendData("cast" & SEP_CHAR & ReadINI("SKP" & SpellMemorized, "sid", App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini") & END_CHAR)
                Player(MyIndex).SpellcdTimer(SpellMemorized) = GetTickCount
                Player(MyIndex).spellcdb(SpellMemorized) = True
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
                Timer1.Interval = 0
                Timer1.Enabled = False
                ccrpProgressBar1.min = 0
                ccrpProgressBar1.value = 0
                ccrpProgressBar1.Clear
                frmMirage.barracasteo.Visible = False
            End If
        'End If
End Sub
Sub Spellincd(SpellID As Byte) 'SpellID es la posición que ocupa dentro de la lista de Hechizos del jugador de 1 a 20



If Player(MyIndex).spellcdb(SpellID) = True Then
    If Player(MyIndex).SpellcdTimer(SpellID) + Spell(SpellID).TimeToCast * 1000 < GetTickCount Then
            Player(MyIndex).spellcdb(SpellID) = False
            If Player(MyIndex).Spellpos(SpellID) <> 0 Then
                If FileExist("GUI\Hechizos\" & Trim(Spell(Player(MyIndex).Spellpos(SpellID)).name) & ".gif") Then
                    Imagesb(SpellID - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & Trim(Spell(Player(MyIndex).Spellpos(SpellID)).name) & ".gif")
                Else
                    Imagesb(SpellID - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\X.gif")
                End If
            End If
    Else
        If Player(MyIndex).Spellpos(SpellID) <> 0 Then
            If Player(MyIndex).Spellpos(SpellID) = ReadINI("SK" & SpellID, "sid", App.Path & "\Scripts\" & GetPlayerName(MyIndex) & ".ini") Then
                If FileExists("GUI\Hechizos\" & Trim(Spell(Player(MyIndex).Spellpos(SpellID)).name) & "_CD.gif") Then
                        Imagesb(SpellID - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & Trim(Spell(Player(MyIndex).Spellpos(SpellID)).name) & "_CD.gif")
                Else
                        Imagesb(SpellID - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\X.gif")
                End If
            End If
        End If
    End If
Else
    If Player(MyIndex).SpellcdTimer(SpellID) + Spell(SpellID).TimeToCast * 1000 > GetTickCount Then
        Player(MyIndex).spellcdb(SpellID) = False
            If Player(MyIndex).Spellpos(SpellID) <> 0 Then
                If FileExists("GUI\Hechizos\" & Trim(Spell(Player(MyIndex).Spellpos(SpellID)).name) & ".gif") Then
                    Imagesb(SpellID - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & Trim(Spell(Player(MyIndex).Spellpos(SpellID)).name) & ".gif")
                Else
                    Imagesb(SpellID - 1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\X.gif")
                End If
            End If
    End If
End If

End Sub

Private Sub Timer2_Timer()
Dim I As Byte
For I = 1 To 10
Call Spellincd(I)
Next I

End Sub

Private Sub tmrGameClock_Timer()
    IncrementGameClock
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then
        tmrRainDrop.Enabled = False
        Exit Sub
    End If

    If BLT_RAIN_DROPS > 0 Then
        If DropRain(BLT_RAIN_DROPS).Randomized = False Then
            Call RNDRainDrop(BLT_RAIN_DROPS)
        End If
    End If

    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1

    If tmrRainDrop.Interval > 30 Then
        tmrRainDrop.Interval = tmrRainDrop.Interval - 10
    End If
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then
        tmrSnowDrop.Enabled = False
        Exit Sub
    End If

    If BLT_SNOW_DROPS > 0 Then
        If DropSnow(BLT_SNOW_DROPS).Randomized = False Then
            Call RNDSnowDrop(BLT_SNOW_DROPS)
        End If
    End If

    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1

    If tmrSnowDrop.Interval > 30 Then
        tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
    End If
End Sub

Private Sub txtChat_GotFocus()
    On Error Resume Next

    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub picInventory_Click()
    picInventory.Visible = True
End Sub

Private Sub lblUseItem_Click()
    Call UseItem
End Sub

Private Sub lblDropItem_Click()
    Call DropItem
End Sub

Private Sub lblCast_Click()
If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
             SpellMemorized = Player(MyIndex).Spell(lstSpells.ListIndex + 1)

            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
            If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
            Call LoadSCT(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Call BltSpellsBar(MyIndex, Player(MyIndex).Spellpos(SpellMemorized))
            Else
                Call AddText("No tienes un hechizo en este hueco.", BRIGHTRED)
            End If
End Sub


Private Sub cmdAccess_Click()
    Call SendChangeGuildAccess(txtName.Text, txtAccess.Text)
End Sub

Private Sub cmdDisown_Click()
    Call SendGuildDisown(txtName.Text)
End Sub

Private Sub cmdTrainee_Click()
    Call SendSetTrainee(txtName.Text)
End Sub

Private Sub picUp_Click()
    If scrlInventory.value <> 0 Then
        scrlInventory.value = scrlInventory.value - 1
        picInventory3.top = scrlInventory.value * -PIC_Y
    End If
End Sub

Private Sub picDown_Click()
    If scrlInventory.value <> 1 Then
        scrlInventory.value = scrlInventory.value + 1
        picInventory3.top = scrlInventory.value * -PIC_Y
    End If
End Sub

Private Sub lstSpells_GotFocus()
    On Error Resume Next
    picScreen.SetFocus
End Sub
