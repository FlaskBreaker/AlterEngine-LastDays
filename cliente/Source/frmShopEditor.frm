VERSION 5.00
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Tiendas"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5910
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
   Icon            =   "frmShopEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSellsItems 
      Caption         =   "Compra Objetos"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Ver info del objeto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Objetos de la tienda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5655
      Begin VB.Frame frmAddEditItem 
         Caption         =   "Objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   5295
         Begin VB.CommandButton cmdAECancel 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2160
            TabIndex        =   20
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdAddEdit 
            Caption         =   "Añadir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   19
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtPrice 
            Appearance      =   0  'Flat
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
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtNumber 
            Appearance      =   0  'Flat
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
            Left            =   960
            TabIndex        =   17
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cmbItemList 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Precio:"
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
            Left            =   2160
            TabIndex        =   15
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Numero:"
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
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Objeto:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdDelItem 
         Caption         =   "Eliminar Objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "Editar Objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Añadir Objeto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox lstItems 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CheckBox chkFixesItems 
      Caption         =   "Arregla Objetos"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
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
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propiedades Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cmbCurrency 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmShopEditor.frx":0FC2
         Left            =   1200
         List            =   "frmShopEditor.frx":0FC4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblCurrency 
         Alignment       =   1  'Right Justify
         Caption         =   "Moneda Usada:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin Eclipse.jcbutton cmdOk 
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
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
      PictureNormal   =   "frmShopEditor.frx":0FC6
      PictureHot      =   "frmShopEditor.frx":17AA
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Eclipse.jcbutton cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
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
      PictureNormal   =   "frmShopEditor.frx":1F8E
      PictureHot      =   "frmShopEditor.frx":28E2
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   4
      TooltipBackColor=   0
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmShopEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private addItem As Boolean

' Temporary array so that we don't modify the shop while editing
Private ShopItemList(1 To MAX_SHOP_ITEMS) As ShopItemRec

' Loads shop item data into our temp array
Public Sub LoadShopItemData(shopNum As Integer)
    Dim I As Integer

    For I = 1 To 25
        ShopItemList(I).Amount = Shop(shopNum).ShopItem(I).Amount
        ShopItemList(I).ItemNum = Shop(shopNum).ShopItem(I).ItemNum
        ShopItemList(I).Price = Shop(shopNum).ShopItem(I).Price
    Next I
End Sub

' Adds the specified item to the shop list and temporary array
Public Sub AddShopItem(ByVal itemN As Integer, ByVal prc As Integer, ByVal cItem As Integer, Optional ByVal AMT As Integer = 0)
    Dim itemStr As String
    If itemN > 0 And itemN <= MAX_ITEMS Then

        If Item(itemN).Stackable = 1 Then
            ' It's stackable so add the amount
            itemStr = AMT & " "
        End If

        ' Add the rest
        itemStr = itemStr & Trim$(Item(itemN).name) & " por " & prc & " " & Trim$(Item(cItem).name)

        lstItems.addItem itemStr

        ' Add to the temp array
        ShopItemList(lstItems.ListCount).Amount = AMT
        ShopItemList(lstItems.ListCount).ItemNum = itemN
        ShopItemList(lstItems.ListCount).Price = prc
    End If
End Sub

' Edits the shop item in the list and array
Public Sub EditShopItem(ByVal index As Integer, ByVal itemN As Integer, ByVal prc As Integer, ByVal cItem As Integer, Optional ByVal AMT As Integer = 0)
    Dim itemStr As String

    If itemN > 0 And itemN <= MAX_ITEMS Then
        If index >= 0 And index <= MAX_SHOP_ITEMS Then

            ' Delete the existing item
            Call lstItems.RemoveItem(index)

            If Item(itemN).Stackable = 1 Then
                ' It's stackable so add the amount
                itemStr = AMT & " "
            End If

            ' Add the rest
            itemStr = itemStr & Trim$(Item(itemN).name) & " por " & prc & " " & Trim$(Item(cItem).name)

            lstItems.addItem itemStr, index

            ' Add to temp array
            ShopItemList(index + 1).Amount = AMT
            ShopItemList(index + 1).ItemNum = itemN
            ShopItemList(index + 1).Price = prc
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
If lstItems.ListCount + 1 > MAX_SHOP_ITEMS Then
    Call MsgBox("Máximo de Items Alcanzado!", vbOKOnly)
    Exit Sub
End If
    frmAddEditItem.Visible = True
    frmAddEditItem.Caption = "Añadir Objeto"
    cmdAddEdit.Caption = "Añadir Objeto"
    addItem = True

    ' Make all the values blank
    cmbItemList.ListIndex = 0
    txtNumber.Text = vbNullString
    txtPrice.Text = vbNullString
End Sub

Private Sub cmdAddEdit_Click()
    Dim currencyItem As Integer
    If cmbItemList.ListIndex + 1 > MAX_SHOP_ITEMS Then
        Call MsgBox("Máximo de Items Alcanzado!", vbOKOnly)
        Exit Sub
    End If

    currencyItem = cmbCurrency.ItemData(cmbCurrency.ListIndex)
    ' Check for invalid input
    If Not IsNumeric(txtNumber.Text) Or Not IsNumeric(txtPrice.Text) Then
        Call MsgBox("Entrada invalida - porfavor introduce un numero!", vbExclamation)
    Else
        ' Input was okay
        If addItem Then
            Call AddShopItem(cmbItemList.ListIndex + 1, Val(txtPrice.Text), currencyItem, Val(txtNumber.Text))
        Else
            ' Edit the item - make sure something was selected
            If lstItems.ListIndex >= 0 Then
                Call EditShopItem(lstItems.ListIndex, cmbItemList.ListIndex + 1, Val(txtPrice.Text), cmbCurrency.ListIndex + 1, Val(txtNumber.Text))
            End If
        End If
        frmAddEditItem.Visible = False
    End If

End Sub

Private Sub cmdAECancel_Click()
    frmAddEditItem.Visible = False
End Sub

Private Sub cmdDelItem_Click()
    If lstItems.ListIndex > 0 Then
        ' Remove the item
        Call lstItems.RemoveItem(lstItems.ListIndex)
    End If
End Sub

Private Sub cmdEditItem_Click()
    If lstItems.ListIndex > -1 Then
        frmAddEditItem.Visible = True
        addItem = False
        cmdAddEdit.Caption = "Ok"

        ' Set all the values
        cmbItemList.ListIndex = ShopItemList(lstItems.ListIndex + 1).ItemNum - 1
        txtNumber.Text = ShopItemList(lstItems.ListIndex + 1).Amount
        txtPrice.Text = ShopItemList(lstItems.ListIndex + 1).Price
    Else
        MsgBox "Selecciona un objeto primero!"
    End If
End Sub

Private Sub cmdOk_Click()
    Call ShopEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ShopEditorCancel
End Sub

' Returns amount of shopitem in the temp array
Public Function GetShopItemAmt(ByVal Item As Integer) As Integer
    If Item > 0 And Item < lstItems.ListCount + 1 Then
        GetShopItemAmt = ShopItemList(Item).Amount
    ElseIf Item < 26 Then
        GetShopItemAmt = 0
    End If
End Function

' Returns item num of shopitem in temp array
Public Function GetShopItemNum(ByVal Item As Integer) As Integer
    If Item > 0 And Item < lstItems.ListCount + 1 Then
        GetShopItemNum = ShopItemList(Item).ItemNum
    ElseIf Item < 26 Then
        GetShopItemNum = 0
    End If
End Function

' Returns item price of shopitem in temp array
Public Function GetShopItemPrice(ByVal Item As Integer) As Integer
    If Item > 0 And Item < lstItems.ListCount + 1 Then
        GetShopItemPrice = ShopItemList(Item).Price
    ElseIf Item < 26 Then
        GetShopItemPrice = 0
    End If
End Function
