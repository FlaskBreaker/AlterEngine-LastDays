VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColorPjs 
   Caption         =   "Color Pjs -"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   8
   End
   Begin Server.jcbutton BtnOk 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Guardar"
      CaptionEffects  =   0
   End
   Begin Server.jcbutton BtnCancel 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancelar"
      CaptionEffects  =   0
   End
   Begin Server.jcbutton UsersCr 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
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
      Caption         =   "Users"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Server.jcbutton ModsCr 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
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
      Caption         =   "Mods"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Server.jcbutton MappersCr 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
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
      Caption         =   "Mappers"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Server.jcbutton DevelopersCr 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
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
      Caption         =   "Developers"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Server.jcbutton AdminsCr 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2040
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
      Caption         =   "Admins"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Server.jcbutton OwnersCr 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
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
      Caption         =   "Owners"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Server.jcbutton PKsCr 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3000
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
      Caption         =   "Pks"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4455
   End
End
Attribute VB_Name = "frmColorPjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R, G, B As Byte
Dim SelectCr As Byte
Dim CrUser(1 To 3) As Byte
Dim CrMod(1 To 3) As Byte
Dim CrMapper(1 To 3) As Byte
Dim CrDeveloper(1 To 3) As Byte
Dim CrAdmin(1 To 3) As Byte
Dim CrOwner(1 To 3) As Byte
Dim CrPK(1 To 3) As Byte
'Evita el Cambio de Colores Accidentales
Dim Slidedisabled As Boolean


Private Sub AdminsCr_Click()
Slidedisabled = True
SelectCr = 5
Slider1.Value = CrAdmin(1)
Slider2.Value = CrAdmin(2)
Slider3.Value = CrAdmin(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False
End Sub

Private Sub BtnCancel_Click()
Unload Me
End Sub

Function comparecolors(ByVal index As Long, ByVal cchar As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    comparecolors = False
    If Player(index).Char(cchar).color(1) = Red Then
        If Player(index).Char(cchar).color(2) = Green Then
            If Player(index).Char(cchar).color(3) = Blue Then
            comparecolors = True
            End If
        End If
    End If
End Function

Private Sub BtnOk_Click()
Dim x As Long
Dim i As Long
Dim y As Byte
Dim z As Byte
Dim color As String

x = MsgBox("Estás Seguro de querer usar esta conbinación de colores?", vbYesNo)

If x = vbNo Then
    Exit Sub
Else
    i = 1
    Do While i < MAX_PLAYERS - 1
        For y = 1 To 3
   
            color = Player(i).Char(y).color
            'If (color = UserCr) Or (color = ModCr) Or (color = MapperCr) Or (color = DeveloperCr) Or (color = AdminCr) Or (color = OwnerCr) Or (color = PKCr) Then
            If comparecolors(i, y, UserCr(1), UserCr(2), UserCr(3)) Or comparecolors(i, y, ModCr(1), ModCr(2), ModCr(3)) Or comparecolors(i, y, MapperCr(1), MapperCr(2), MapperCr(3)) Or comparecolors(i, y, DeveloperCr(1), DeveloperCr(2), DeveloperCr(3)) Or comparecolors(i, y, AdminCr(1), AdminCr(2), AdminCr(3)) Or comparecolors(i, y, OwnerCr(1), OwnerCr(2), OwnerCr(3)) Or comparecolors(i, y, PKCr(1), PKCr(2), PKCr(3)) Then
                Call changeplayercolorname(i, y)
                If IsPlaying(i) Then
                    Call SendDataToMap(GetPlayerMap(i), "namecolor" & SEP_CHAR & i & SEP_CHAR & Player(i).Char(Player(i).CharNum).color(1) & SEP_CHAR & Player(i).Char(Player(i).CharNum).color(2) & SEP_CHAR & Player(i).Char(Player(i).CharNum).color(3) & END_CHAR)
                End If
            End If
        i = i + 1
        Next y
    Loop

    For y = 1 To 3
    UserCr(y) = CrUser(y)
    ModCr(y) = CrMod(y)
    MapperCr(y) = CrMapper(y)
    DeveloperCr(y) = CrDeveloper(y)
    AdminCr(y) = CrAdmin(y)
    OwnerCr(y) = CrOwner(y)
    PKCr(y) = CrPK(y)
    Next y
    'Paso a String para evitar Errores de referencia
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "UserCrR", CStr(CrUser(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "ModCrR", CStr(CrMod(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "MapperCrR", CStr(CrMapper(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrR", CStr(CrDeveloper(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "AdminCrR", CStr(CrAdmin(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrR", CStr(CrOwner(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKCrR", CStr(CrPK(1))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "UserCrG", CStr(CrUser(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "ModCrG", CStr(CrMod(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "MapperCrG", CStr(CrMapper(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrG", CStr(CrDeveloper(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "AdminCrG", CStr(CrAdmin(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrG", CStr(CrOwner(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKCrG", CStr(CrPK(2))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "UserCrB", CStr(CrUser(3))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "ModCrB", CStr(CrMod(3))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "MapperCrB", CStr(CrMapper(3))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrB", CStr(CrDeveloper(3))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "AdminCrB", CStr(CrAdmin(3))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrB", CStr(CrOwner(3))
    PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKCrB", CStr(CrPK(3))

    Call MsgBox("Configuración Guardada Correctamente!", vbOKOnly)
    Unload Me
End If
End Sub
Sub changeplayercolorname(ByVal index As Long, ByVal bchar As Byte)

If comparecolors(index, bchar, UserCr(1), UserCr(2), UserCr(3)) Then
        Player(index).Char(bchar).color(1) = CrUser(1)
        Player(index).Char(bchar).color(2) = CrUser(2)
        Player(index).Char(bchar).color(3) = CrUser(3)
        Exit Sub
ElseIf comparecolors(index, bchar, ModCr(1), ModCr(2), ModCr(3)) Then
        Player(index).Char(bchar).color(1) = CrMod(1)
        Player(index).Char(bchar).color(2) = CrMod(2)
        Player(index).Char(bchar).color(3) = CrMod(3)
        Exit Sub
ElseIf comparecolors(index, bchar, MapperCr(1), MapperCr(2), MapperCr(3)) Then
        Player(index).Char(bchar).color(1) = CrMapper
        Player(index).Char(bchar).color(2) = CrMapper(2)
        Player(index).Char(bchar).color(3) = CrMapper(3)
        Exit Sub
ElseIf comparecolors(index, bchar, DeveloperCr(1), DeveloperCr(2), DeveloperCr(3)) Then
        Player(index).Char(bchar).color(1) = CrDeveloper
        Player(index).Char(bchar).color(2) = CrDeveloper(2)
        Player(index).Char(bchar).color(3) = CrDeveloper(3)
        Exit Sub
ElseIf comparecolors(index, bchar, AdminCr(1), AdminCr(2), AdminCr(3)) Then
        Player(index).Char(bchar).color(1) = CrAdmin(1)
        Player(index).Char(bchar).color(2) = CrAdmin(2)
        Player(index).Char(bchar).color(3) = CrAdmin(3)
        Exit Sub
ElseIf comparecolors(index, bchar, OwnerCr(1), OwnerCr(2), OwnerCr(3)) Then
        Player(index).Char(bchar).color(1) = CrOwner(1)
        Player(index).Char(bchar).color(2) = CrOwner(2)
        Player(index).Char(bchar).color(3) = CrOwner(3)
        Exit Sub
ElseIf comparecolors(index, bchar, PKCr(1), PKCr(2), PKCr(3)) Then
        Player(index).Char(bchar).color(1) = CrPK(1)
        Player(index).Char(bchar).color(2) = CrPK(2)
        Player(index).Char(bchar).color(3) = CrPK(3)
        Exit Sub
End If
        
End Sub
Private Sub DevelopersCr_Click()
Slidedisabled = True
SelectCr = 4
Slider1.Value = CrDeveloper(1)
Slider2.Value = CrDeveloper(2)
Slider3.Value = CrDeveloper(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False
End Sub

Private Sub Form_Load()
Dim y As Byte

For y = 1 To 3
CrUser(y) = UserCr(y)
CrMod(y) = ModCr(y)
CrMapper(y) = MapperCr(y)
CrDeveloper(y) = DeveloperCr(y)
CrAdmin(y) = AdminCr(y)
CrOwner(y) = OwnerCr(y)
CrPK(y) = PKCr(y)
Next y

UsersCr.ForeColor = RGB(CrUser(1), CrUser(2), CrUser(3))
ModsCr.ForeColor = RGB(CrMod(1), CrMod(2), CrMod(3))
MappersCr.ForeColor = RGB(CrMapper(1), CrMapper(2), CrMapper(3))
DevelopersCr.ForeColor = RGB(CrDeveloper(1), CrDeveloper(2), CrDeveloper(3))
AdminsCr.ForeColor = RGB(CrAdmin(1), CrAdmin(2), CrAdmin(3))
OwnersCr.ForeColor = RGB(CrOwner(1), CrOwner(2), CrOwner(3))
PKsCr.ForeColor = RGB(CrPK(1), CrPK(2), CrPK(3))

SelectCr = 1
Shape1.BackColor = RGB(CrUser(1), CrUser(2), CrUser(3))
Slider1.Value = CrUser(1)
Slider2.Value = CrUser(2)
Slider3.Value = CrUser(3)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected

End Sub

Private Sub MappersCr_Click()
Slidedisabled = True
SelectCr = 3

Slider1.Value = CrMapper(1)
Slider2.Value = CrMapper(2)
Slider3.Value = CrMapper(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False
End Sub

Private Sub ModsCr_Click()
Slidedisabled = True
SelectCr = 2

Slider1.Value = CrMod(1)
Slider2.Value = CrMod(2)
Slider3.Value = CrMod(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False

End Sub

Private Sub OwnersCr_Click()
Slidedisabled = True
SelectCr = 6
Slider1.Value = CrOwner(1)
Slider2.Value = CrOwner(2)
Slider3.Value = CrOwner(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False
End Sub

Private Sub PKsCr_Click()
Slidedisabled = True
SelectCr = 7
Slider1.Value = CrPK(1)
Slider2.Value = CrPK(2)
Slider3.Value = CrPK(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False
End Sub

Private Sub Slider1_Change()
R = Slider1.Value
Shape1.BackColor = RGB(R, G, B)
If Slidedisabled Then Exit Sub
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
End Sub

Private Sub Slider3_Change()
B = Slider3.Value
Shape1.BackColor = RGB(R, G, B)
If Slidedisabled Then Exit Sub
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
End Sub

Private Sub Slider2_Change()
G = Slider2.Value
Shape1.BackColor = RGB(R, G, B)
If Slidedisabled Then Exit Sub
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
End Sub

Sub ChangeColorName(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)

    
    Select Case SelectCr
        Case 1
            UsersCr.ForeColor = RGB(Red, Green, Blue)
            CrUser(1) = Red
            CrUser(2) = Green
            CrUser(3) = Blue
            Exit Sub
            
        Case 2
            ModsCr.ForeColor = RGB(Red, Green, Blue)
            CrMod(1) = Red
            CrMod(2) = Green
            CrMod(3) = Blue
            Exit Sub
            
        Case 3
            MappersCr.ForeColor = RGB(Red, Green, Blue)
            CrMapper(1) = Red
            CrMapper(2) = Green
            CrMapper(3) = Blue
            Exit Sub
            
        Case 4
            DevelopersCr.ForeColor = RGB(Red, Green, Blue)
            CrDeveloper(1) = Red
            CrDeveloper(2) = Green
            CrDeveloper(3) = Blue
            Exit Sub
            
        Case 5
            AdminsCr.ForeColor = RGB(Red, Green, Blue)
            CrAdmin(1) = Red
            CrAdmin(2) = Green
            CrAdmin(3) = Blue
            Exit Sub
            
        Case 6
            OwnersCr.ForeColor = RGB(Red, Green, Blue)
            CrOwner(1) = Red
            CrOwner(2) = Green
            CrOwner(3) = Blue
            Exit Sub
            
        Case 7
            PKsCr.ForeColor = RGB(R, Green, Blue)
            CrPK(1) = Red
            CrPK(2) = Green
            CrPK(3) = Blue
            Exit Sub
            
    End Select
    
End Sub

Function checkcolorselected() As String
Select Case SelectCr
        Case 1
            checkcolorselected = "User- R: " & CrUser(1) & " G: " & CrUser(2) & " B: " & CrUser(3)
            Exit Function
            
        Case 2
            checkcolorselected = "Mod- R: " & CrMod(1) & " G: " & CrMod(2) & " B: " & CrMod(3)
            
        Case 3
            checkcolorselected = "Mapper- R: " & CrMapper(1) & " G: " & CrMapper(2) & " B: " & CrMapper(3)
            Exit Function
            
        Case 4
           checkcolorselected = "Developer- R: " & CrDeveloper(1) & " G: " & CrDeveloper(2) & " B: " & CrDeveloper(3)
            Exit Function
            
        Case 5
            checkcolorselected = "Admin- R: " & CrAdmin(1) & " G: " & CrAdmin(2) & " B: " & CrAdmin(3)
            Exit Function
            
        Case 6
            checkcolorselected = "Owner- R: " & CrOwner(1) & " G: " & CrOwner(2) & " B: " & CrOwner(3)
            Exit Function
            
        Case 7
            checkcolorselected = "PK- R: " & CrPK(1) & " G: " & CrPK(2) & " B: " & CrPK(3)
            Exit Function
            
    End Select

End Function

Private Sub UsersCr_Click()
Slidedisabled = True
SelectCr = 1

Slider1.Value = CrUser(1)
Slider2.Value = CrUser(2)
Slider3.Value = CrUser(3)
Call ChangeColorName(R, G, B)
frmColorPjs.Caption = "Colores Pjs - " & checkcolorselected
Slidedisabled = False
End Sub
