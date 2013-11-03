VERSION 5.00
Begin VB.Form frmClases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Clases"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Agregar Clase"
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
      Height          =   975
      Left            =   4560
      TabIndex        =   39
      Top             =   240
      Width           =   3975
      Begin Server.jcbutton AgregarClase 
         Height          =   495
         Left            =   960
         TabIndex        =   41
         Top             =   360
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
         Caption         =   "Agregar Clase "
         PictureNormal   =   "frmClases.frx":0000
         PictureHot      =   "frmClases.frx":0954
         CaptionEffects  =   4
         ColorScheme     =   2
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Mapa Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1935
      Left            =   4680
      TabIndex        =   22
      Top             =   1440
      Width           =   3735
      Begin VB.TextBox mapa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   25
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox coordenadax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox coordenaday 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   1320
         Width           =   375
      End
      Begin Server.jcbutton jcbutton1 
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "?"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   4
         ToolTip         =   "Desde aqui se selecciona en el numero de mapa que aparecera la raza, y las coordenadas precisas."
         TooltipType     =   1
         TooltipIcon     =   1
         TooltipTitle    =   "Mapa Inicial"
         TooltipBackColor=   0
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Coordenada Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Coordenada X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Estados Iniciales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   4680
      TabIndex        =   13
      Top             =   3600
      Width           =   3735
      Begin VB.TextBox fuerza 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox defensa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox velocidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox magia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Magia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Defensa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuerza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clase a editar"
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
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3975
      Begin VB.ComboBox CP 
         Height          =   315
         Left            =   360
         TabIndex        =   40
         Top             =   360
         Width           =   3255
      End
   End
   Begin Server.jcbutton salvarclase 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   5280
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
      Caption         =   "Guardar Clase"
      PictureNormal   =   "frmClases.frx":12A8
      PictureHot      =   "frmClases.frx":1BFC
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton cmdCancel 
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   5280
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Cancelar"
      PictureNormal   =   "frmClases.frx":2550
      PictureHot      =   "frmClases.frx":2EA4
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton salvarclase0 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 0"
      PictureNormal   =   "frmClases.frx":37F8
      PictureHot      =   "frmClases.frx":414C
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información Basica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3975
      Begin Server.jcbutton ayudita 
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   3000
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16641248
         Caption         =   "?"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   4
         ToolTip         =   "Escribe 1 para bloquearla o 0 para desbloquearla. Si la bloqueas no se podrá elegir esta clase al inicio."
         TooltipType     =   1
         TooltipIcon     =   1
         TooltipTitle    =   "Clase Bloqueada"
         TooltipBackColor=   0
      End
      Begin VB.TextBox bloqueado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   29
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox nombreclase 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox sieschicasprite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox sieschicosprite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox descripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Clase Bloqueada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción de la clase:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite para mujer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite para hombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de la clase:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin Server.jcbutton salvarclase2 
      Height          =   495
      Left            =   3240
      TabIndex        =   31
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 2"
      PictureNormal   =   "frmClases.frx":4AA0
      PictureHot      =   "frmClases.frx":53F4
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton salvarclase3 
      Height          =   495
      Left            =   3240
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 3"
      PictureNormal   =   "frmClases.frx":5D48
      PictureHot      =   "frmClases.frx":669C
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton salvarclase4 
      Height          =   495
      Left            =   3240
      TabIndex        =   33
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 4"
      PictureNormal   =   "frmClases.frx":6FF0
      PictureHot      =   "frmClases.frx":7944
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton salvarclase5 
      Height          =   495
      Left            =   3240
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 5"
      PictureNormal   =   "frmClases.frx":8298
      PictureHot      =   "frmClases.frx":8BEC
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton salvarclase6 
      Height          =   495
      Left            =   3240
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 6"
      PictureNormal   =   "frmClases.frx":9540
      PictureHot      =   "frmClases.frx":9E94
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton salvarclase7 
      Height          =   495
      Left            =   3240
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
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
      Caption         =   "Guardar Clase 7"
      PictureNormal   =   "frmClases.frx":A7E8
      PictureHot      =   "frmClases.frx":B13C
      CaptionEffects  =   4
      ColorScheme     =   2
   End
   Begin Server.jcbutton recargarclase 
      Height          =   495
      Left            =   480
      TabIndex        =   42
      Top             =   5280
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
      Caption         =   "Recargar Clases"
      PictureNormal   =   "frmClases.frx":BA90
      PictureHot      =   "frmClases.frx":C3E4
      CaptionEffects  =   4
      ColorScheme     =   2
   End
End
Attribute VB_Name = "frmClases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub AgregarClase_Click()
MAX_CLASSES = MAX_CLASSES + 1

   Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Name", nombreclase.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "MaleSprite", sieschicasprite.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "FemaleSprite", sieschicosprite.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Desc", descripcion.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "STR", fuerza.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "DEF", defensa.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "SPEED", velocidad.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "MAGI", magia.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "MAP", mapa.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "X", coordenadax.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Y", coordenaday.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Locked", bloqueado.Text)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Gender", 1)
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Gender1", "Hombre")
    Call PutVar(App.Path & "\Clases\Class" & MAX_CLASSES & ".ini", "CLASS", "Gender2", "Mujer")
    
    Call MsgBox("La Class" & MAX_CLASSES & ", ha sido guardada con exito.")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CP_Click()
' Comprobamos que el archivo de la clase existe.
    If FileExists("\Clases\" & CP.List(CP.ListIndex) & ".ini") Then
        
    nombreclase.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Name")
    sieschicasprite.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "MaleSprite")
    sieschicosprite.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "FemaleSprite")
    descripcion.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Desc")
    fuerza.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "STR")
    defensa.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "DEF")
    velocidad.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "SPEED")
    magia.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "MAGI")
    
    mapa.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "MAP")
    coordenadax.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "X")
    coordenaday.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Y")
    
    bloqueado.Text = GetVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Locked")
    
    Else
        MsgBox "No se ha encontrado el archivo", vbInformation
    End If
End Sub

Private Sub recargarclase_Click()
Dim i As Long
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "Todas las Clases Fueron Recargadas.", True)
    MsgBox ("Todas las Clases Fueron Recargadas.")
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
End Sub

Private Sub salvarclase_Click()
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Name", nombreclase.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "MaleSprite", sieschicasprite.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "FemaleSprite", sieschicosprite.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Desc", descripcion.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "STR", fuerza.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "DEF", defensa.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "SPEED", velocidad.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "MAGI", magia.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "MAP", mapa.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "X", coordenadax.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Y", coordenaday.Text)
    Call PutVar(App.Path & "\Clases\" & CP.List(CP.ListIndex) & ".ini", "CLASS", "Locked", bloqueado.Text)
    
    Call MsgBox("La " & CP.List(CP.ListIndex) & ", ha sido guardada con exito.")
End Sub
