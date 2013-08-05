VERSION 5.00
Begin VB.Form frmFixItem 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparar Objeto"
   ClientHeight    =   5700
   ClientLeft      =   135
   ClientTop       =   315
   ClientWidth     =   8070
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFixItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFixItem.frx":0FC2
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbItem 
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
      ItemData        =   "frmFixItem.frx":124FA
      Left            =   2640
      List            =   "frmFixItem.frx":124FC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2640
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciona el objeto a reparar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   4320
      TabIndex        =   2
      Top             =   4440
      Width           =   2355
   End
   Begin VB.Label picFix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   4440
      Width           =   2325
   End
End
Attribute VB_Name = "frmFixItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picFix_Click()
    Call SendData("fixitem" & SEP_CHAR & snumber & SEP_CHAR & cmbItem.ListIndex + 1 & END_CHAR)
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

        If FileExists("GUI\Arreglar_Objeto" & Ending) Then
            frmFixItem.Picture = LoadPicture(App.Path & "\GUI\Arreglar_Objeto" & Ending)
        End If
    Next I

    frmFixItem.cmbItem.Clear
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Then
                frmFixItem.cmbItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Then
                    frmFixItem.cmbItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (worn)"
                Else
                    frmFixItem.cmbItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
                End If
            End If
        Else
            frmFixItem.cmbItem.addItem I & "> Vacio"
        End If
    Next I

    frmFixItem.cmbItem.ListIndex = 0
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub
