VERSION 5.00
Begin VB.Form frmSellItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vender Objetos"
   ClientHeight    =   6555
   ClientLeft      =   465
   ClientTop       =   660
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSellItem.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   0
   End
   Begin VB.ListBox lstSellItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4515
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin VB.Timer tmrClear 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   0
   End
   Begin Eclipse.jcbutton lblSellItem 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   49152
      Caption         =   "Vender Objeto"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Eclipse.jcbutton Label1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744576
      Caption         =   "Refrescar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Eclipse.jcbutton CloseSell 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421631
      Caption         =   "Cerrar"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vender Objetos"
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
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   3240
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label lblSold 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   3255
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Dim I As Long


    ' frmBank.lblBank.Caption = Trim$(Map(GetPlayerMap(MyIndex)).Name)
    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Then
                frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Then
                    frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (equipado)"
                Else
                    frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name)
                End If
            End If
        Else
            frmBank.lstInventory.addItem I & "> Vacio"
        End If

    Next I
    frmSellItem.lstSellItem.Clear
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                frmSellItem.lstSellItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Then
                    frmSellItem.lstSellItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (equipado)"
                Else
                    frmSellItem.lstSellItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name)
                End If
            End If
        Else
            frmSellItem.lstSellItem.addItem I & "> Vacio"
        End If
    Next I
    frmSellItem.lstSellItem.ListIndex = 0
End Sub

Private Sub lblSellItem_Click()
    Dim packet As String
    Dim ItemNum As Long
    Dim ItemSlot As Integer
    Dim AMT As Long

    ItemNum = GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))
    ItemSlot = lstSellItem.ListIndex + 1
    If GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1)) > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Then
            Exit Sub
        Else
            If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
                Exit Sub
            Else
                If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price > 0 Then
                    If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Stackable = 1 Then
                        AMT = InputBox("Cuanta cantidad de " & Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).name & " te gustaria vender?", "Vender " & Trim$(Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).name), 0)
                        If IsNumeric(AMT) Then
                            packet = "sellitem" & SEP_CHAR & snumber & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & AMT & END_CHAR
                            Call SendData(packet)
                            lblSold.Caption = "Has vendido " & AMT & " " & Trim$(Item(ItemNum).name) & " ."
                        End If
                    Else
                        packet = "sellitem" & SEP_CHAR & snumber & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & 1 & END_CHAR
                        Call SendData(packet)
                        lblSold.Caption = "Has vendido 1 " & Trim$(Item(ItemNum).name) & "."
                    End If
                    tmrClear.Enabled = True

                Else
                    Exit Sub
                End If
            End If
        End If
    Else
        Exit Sub
    End If
    Timer1.Enabled = True

End Sub



Private Sub lstSellItem_Click()
    If GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1)) > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Then
            lblPrice.Caption = "Eso no es una selección valida"
        Else
            If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
                lblPrice.Caption = "Por favor, desequipate el objeto primero."
            Else
                If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price > 0 Then
                    lblPrice.Caption = "Precio: " & Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price & " Monedas"
                Else
                    lblPrice.Caption = "No esta en venta"
                End If
            End If
        End If
    Else
        lblPrice.Caption = "No es una selección valida."
    End If
End Sub
Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String
    For I = 1 To 3
        If I = 1 Then
            Ending = ".GIF"
        End If
        If I = 2 Then
            Ending = ".JPG"
        End If
        If I = 3 Then
            Ending = ".PNG"
        End If

        If FileExists("GUI\Vender" & Ending) Then
            frmSellItem.Picture = LoadPicture(App.Path & "\GUI\Vender" & Ending)
        End If
    Next I
    lblSold.Caption = vbNullString
    lblPrice.Caption = vbNullString
    frmSellItem.lstSellItem.Clear
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                frmSellItem.lstSellItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Then
                    frmSellItem.lstSellItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (equipado)"
                Else
                    frmSellItem.lstSellItem.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name)
                End If
            End If
        Else
            frmSellItem.lstSellItem.addItem I & "> Vacio"
        End If
    Next I
    frmSellItem.lstSellItem.ListIndex = 0
End Sub

Private Sub Timer1_Timer()
    Call Label1_Click
    Timer1.Enabled = False
End Sub

Private Sub tmrClear_Timer()
    lblSold.Caption = vbNullString

End Sub


Private Sub CloseSell_Click()
    Unload Me
End Sub
