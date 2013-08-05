VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "RSS Reader"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   7650
   ForeColor       =   &H80000013&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Estadisticas:"
      Height          =   1395
      Left            =   60
      TabIndex        =   11
      Top             =   2220
      Width           =   3255
      Begin VB.Label fu 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1380
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label fp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1380
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label totar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1380
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total de archivos:"
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Primero:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ultimo:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.TextBox lin 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      MouseIcon       =   "Form1.frx":038A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1680
      Width           =   2235
   End
   Begin VB.TextBox nom 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Datos del Archivo:"
      Height          =   1875
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   3615
      Begin VB.TextBox fech 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   2235
      End
      Begin VB.TextBox tam 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Link Descarga:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño KB:"
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1620
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Image actualizar 
      Height          =   480
      Left            =   5400
      MouseIcon       =   "Form1.frx":0694
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":099E
      ToolTipText     =   "Actualizar"
      Top             =   2700
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   3780
      X2              =   3780
      Y1              =   360
      Y2              =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de archivos Subidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   2415
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu con 
         Caption         =   "Actualizar"
         Index           =   0
      End
      Begin VB.Menu con 
         Caption         =   "Ver Ultimos"
         Index           =   1
      End
      Begin VB.Menu con 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu con 
         Caption         =   "Salir"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RSS Reader,By Sebastián Andrés Perdomo(seba123neo) 2008*
'********************************************************
Option Explicit
Dim TotalArchivos As Integer
Public Resultado As Integer
Private WithEvents Icono As cSystray
Attribute Icono.VB_VarHelpID = -1
Dim Contador As Integer
Dim Flag As Boolean

Private Sub actualizar_Click()
Flag = True
Call Form_Load
End Sub

Private Sub con_Click(Index As Integer)
Select Case Index
Case 0
Call actualizar_Click
Case 1
With Me
.WindowState = 0
.Show
.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End With
archivo.Visible = False
Case 3
Unload Me
End Select
End Sub

Private Sub Form_Initialize()
Call XP
End Sub

Private Sub Form_Load()
Set Doc = New DOMDocument
Doc.resolveExternals = True
Doc.async = False
If Doc.Load("http://www.uploadsourcecode.com.ar/rss.php") = False Then
MsgBox "No se pudo conectar,compruebe su conexion a internet", vbExclamation
Unload Me
Else
TotalArchivos = Doc.selectNodes("/rss/channel/item").Length
If Flag = True Then
If CInt(totar.Caption) = TotalArchivos Then
MsgBox "No hay Nuevos Archivos...", vbInformation
End If
If CInt(totar.Caption) < TotalArchivos Then
Resultado = TotalArchivos - CInt(totar.Caption)
Form2.Show
End If
End If
totar.Caption = TotalArchivos
With Info
.FPrimero = GetElement("rss/channel/item/pubDate", TotalArchivos - 1)
.FUltimo = GetElement("rss/channel/item/pubDate", 0)
fu.Caption = .FUltimo
fp.Caption = .FPrimero
End With
List1.Clear
For Contador = 0 To TotalArchivos - 1
With Info
.NombreArchivo = GetElement("rss/channel/item/title", Contador)
List1.AddItem .NombreArchivo
End With
Next Contador
End If
List1.ListIndex = 0
End Sub

Private Function GetElement(NodeKey As String, Indice As Integer)
GetElement = Doc.getElementsByTagName(NodeKey).Item(Indice).Text
End Function

Private Sub Icono_MouseUp(Button As Integer)
If Button = 2 Then
Icono.BeforePopup
archivo.Visible = True
Me.PopupMenu archivo
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Icono = Nothing
Set Doc = Nothing
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
Me.Hide
Set Icono = New cSystray
With Icono
If .IsBalloonCapable Then
.BalloonIcon = TTIconUser
.BalloonText = "RSS Reader esta minimizado"
.BalloonTitle = Me.Caption
End If
.SysTrayIconFromHandle Me.Icon
.SysTrayToolTip = Me.Caption
.SysTrayShow True
If .IsBalloonCapable Then .BalloonShow True, 3000
End With
End If
If Me.WindowState = 0 Then
Set Icono = Nothing
End If
End Sub

Private Sub lin_Click()
Dim res As Long
res = ShellExecute(Me.hwnd, vbNullString, lin.Text, 0, 0, 1)
End Sub

Private Sub List1_Click()
With Info
    .NombreArchivo = GetElement("rss/channel/item/title", List1.ListIndex)
    .Tamaño = GetElement("rss/channel/item/description", List1.ListIndex)
    .Link = GetElement("rss/channel/item/link", List1.ListIndex)
    .Fecha = GetElement("rss/channel/item/pubDate", List1.ListIndex)
    nom.Text = .NombreArchivo
    lin.Text = .Link
    tam.Text = Replace(.Tamaño, " kb", "")
    fech.Text = .Fecha
End With
End Sub
