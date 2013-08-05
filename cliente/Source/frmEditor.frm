VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmEditor 
   Caption         =   "AlterEngine - Editor de scripts online"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12000
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   1193
      ScaleHeight     =   7425
      ScaleWidth      =   9585
      TabIndex        =   2
      Top             =   773
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   4
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox txtCommands 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   -20
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmEditor.frx":058A
         Top             =   0
         Width           =   9615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de comandos disponibles para usar en AlterEngine"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   7080
         Width           =   3945
      End
   End
   Begin CodeSenseCtl.CodeSense RT 
      Height          =   4455
      Left            =   0
      OleObjectBlob   =   "frmEditor.frx":574F
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnuSave 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Editar"
      Begin VB.Menu mnuUndo 
         Caption         =   "Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Rehacer"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cortar"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Pegar"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Encontrar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Encontrar siguiente"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Remplazar"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGoLine 
         Caption         =   "Ir a la linea"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Seleccionar todo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelLine 
         Caption         =   "Seleccionar linea actual"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuSC 
         Caption         =   "Comandos de Scripts"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Dim sLastWord As String
Dim sIntellText As String
Dim LBoxPos As Long

Public Sub EditorSetVals()
    RT.Language = "Basic"
    
    RT.SetColor cmClrBookmark, ClrData(0).frClr
    RT.SetColor cmClrBookmarkBk, ClrData(0).bgClr
    RT.SetColor cmClrCommentBk, ClrData(1).bgClr
    RT.SetColor cmClrComment, ClrData(1).frClr
    RT.SetColor cmClrHDividerLines, ClrData(2).frClr
    RT.SetColor cmClrVDividerLines, ClrData(3).frClr
    RT.SetColor cmClrHighlightedLine, ClrData(4).frClr
    RT.SetColor cmClrKeyword, ClrData(5).frClr
    RT.SetColor cmClrKeywordBk, ClrData(5).bgClr
    RT.SetColor cmClrLeftMargin, ClrData(6).frClr
    RT.SetColor cmClrLineNumber, ClrData(7).frClr
    RT.SetColor cmClrLineNumberBk, ClrData(7).bgClr
    RT.SetColor cmClrNumber, ClrData(8).frClr
    RT.SetColor cmClrNumberBk, ClrData(8).bgClr
    RT.SetColor cmClrOperator, ClrData(9).frClr
    RT.SetColor cmClrOperatorBk, ClrData(9).bgClr
    RT.SetColor cmClrScopeKeyword, ClrData(10).frClr
    RT.SetColor cmClrScopeKeywordBk, ClrData(10).bgClr
    RT.SetColor cmClrString, ClrData(11).frClr
    RT.SetColor cmClrStringBk, ClrData(11).bgClr
    RT.SetColor cmClrTagElementName, ClrData(12).frClr
    RT.SetColor cmClrTagElementNameBk, ClrData(12).bgClr
    RT.SetColor cmClrTagEntity, ClrData(13).frClr
    RT.SetColor cmClrTagEntityBk, ClrData(13).bgClr
    RT.SetColor cmClrTagAttributeName, ClrData(14).frClr
    RT.SetColor cmClrTagAttributeNameBk, ClrData(14).bgClr
    RT.SetColor cmClrTagText, ClrData(15).frClr
    RT.SetColor cmClrTagTextBk, ClrData(15).bgClr
    RT.SetColor cmClrText, ClrData(16).frClr
    RT.SetColor cmClrTextBk, ClrData(16).bgClr
    RT.SetColor cmClrWindow, ClrData(17).frClr
      
    Dim iHG As Integer
    iHG = CInt(GetSetting(App.EXEName, "EditOptions", "Highlight", "1"))
    If iHG = 0 Then
        RT.HighlightedLine = -1
    End If
    RT.LineNumbering = CBool(GetSetting(App.EXEName, "EditOptions", "linenumber", "1"))
    RT.DisplayLeftMargin = CBool(GetSetting(App.EXEName, "EditOptions", "leftmargin", "1"))
    RT.DisplayWhitespace = CBool(GetSetting(App.EXEName, "EditOptions", "whitespace", "0"))
    RT.SmoothScrolling = CBool(GetSetting(App.EXEName, "EditOptions", "smoothscroll", "1"))
    RT.LineNumberStart = 1
    RT.EnableDragDrop = True
    RT.ExpandTabs = True
    RT_SelChange RT

End Sub

Private Sub Command1_Click()
    Picture1.Visible = False
End Sub

Private Sub Form_Load()
ResetAllEditVals
GetEditColors

'If Dir(App.Path & "\" & AFileName) <> "" Then
If Dir(App.Path & "\Scripts\Main.txt") <> "" Then
    hFile = FreeFile
    'Open App.Path & "\" & AFileName For Input As #hFile
    Open App.Path & "\Scripts\Main.txt" For Input As #hFile
        frmEditor.RT.Text = input$(LOF(hFile), hFile)
    Close #hFile
End If
End Sub

Public Sub DoHighLight()
    Dim R As CodeSenseCtl.Range
    Set R = RT.GetSel(True)
    If CInt(GetSetting(App.EXEName, "EditOptions", "Highlight", "1")) = 1 Then
      RT.HighlightedLine = R.EndLineNo
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        RT.Width = Me.Width - 120
        RT.Height = (Me.Height) - 800
    End If
End Sub

Private Sub mnuBClearALL_Click()
    RT.DisplayLeftMargin = True
    RT.ExecuteCmd cmCmdBookmarkClearAll
End Sub

Private Sub mnuBGoPrev_Click()
    RT.DisplayLeftMargin = True
    RT.ExecuteCmd cmCmdBookmarkPrev
End Sub

Private Sub mnuBJumpFirst_Click()
    RT.DisplayLeftMargin = True
    RT.ExecuteCmd cmCmdBookmarkJumpToFirst
End Sub

Private Sub mnuBJumpLast_Click()
    RT.DisplayLeftMargin = True
    RT.ExecuteCmd cmCmdBookmarkJumpToLast
End Sub

Private Sub mnuBNext_Click()
    RT.DisplayLeftMargin = True
    RT.ExecuteCmd cmCmdBookmarkNext
End Sub

Private Sub mnuBToggle_Click()
    RT.DisplayLeftMargin = True
    RT.ExecuteCmd cmCmdBookmarkToggle
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText RT.SelText
End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText RT.SelText
    RT.SelText = ""
End Sub

Private Sub mnuDelete_Click()
    RT.SelText = ""
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFind_Click()
    RT.ExecuteCmd cmCmdFind
End Sub

Private Sub mnuFindNext_Click()
    RT.ExecuteCmd cmCmdFindNext
End Sub

Private Sub mnuGoLine_Click()
    RT.ExecuteCmd cmCmdGotoLine, -1
End Sub

Private Sub mnuPaste_Click()
    RT.Paste
End Sub

Private Sub mnuRedo_Click()
    RT.Redo
End Sub

Private Sub mnuReplace_Click()
    RT.ExecuteCmd cmCmdFindReplace
End Sub

Private Sub mnuSave_Click()
    'Open App.Path & "\" & AFileName For Output As #1
    Open App.Path & "\Scripts\Main.txt" For Output As #1
        Print #1, RT.Text
    Close #1
    
    ' If you are making this editor edit more than scripts...
    ' Find a way to change this!
    Call SendData("NEWMAIN" & SEP_CHAR & RT.Text & SEP_CHAR & END_CHAR)

    Unload Me
End Sub

Private Sub mnuSC_Click()
    Picture1.Visible = True
End Sub

Private Sub mnuSelAll_Click()
    RT.ExecuteCmd cmCmdSelectAll
End Sub

Private Sub mnuSelLine_Click()
    RT.ExecuteCmd cmCmdSelectLine
End Sub

Private Sub mnuUndo_Click()
    RT.Undo
End Sub

Private Function RT_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
    GetRange
End Function

Private Function RT_KeyUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
Dim R As CodeSenseCtl.Range

    If KeyCode = 9 Or KeyCode = 13 Then
        AddIntellWord
    End If

    If RT.CurrentWord <> "." Then
        sLastWord = RT.CurrentWord
    End If
    
    If KeyCode = 190 Then

    Set R = RT.GetSel(False)

    LBoxPos = R.EndColNo
        RT.ExecuteCmd cmCmdCodeList
    End If
End Function
Private Function RT_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    'ListCtrl.hImageList = IMGIntellisence.hImageList
End Function

Private Function RT_CodeListCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    AddIntellWord
    RT_CodeListCancel = False
End Function
Private Function RT_CodeListChar(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal wChar As Long, ByVal lKeyData As Long) As Boolean
    RT_CodeListChar = False
End Function
Private Function RT_CodeListSelChange(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As String
    sIntellText = ListCtrl.GetItemText(lItem)
    RT_CodeListSelChange = ""
End Function
Private Function RT_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    AddIntellWord
    RT_CodeListSelMade = False
End Function
Private Function RT_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
    RT_CodeListSelWord = True
End Function
Private Function RT_CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
    RT_CodeTip = cmToolTipTypeNormal
End Function
Private Function RT_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long) As Boolean
    GetRange
    If Button = 2 Then
        Me.PopupMenu Me.mnuEdit
    End If
End Function

Private Function RT_MouseUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long) As Boolean
    GetRange
End Function

Private Sub RT_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
    DoHighLight
End Sub

Private Sub GetRange()
Dim R As CodeSenseCtl.Range
Dim LLine As Long
Dim LCurrent As Long
    Set R = RT.GetSel(False)
    LLine = R.EndLineNo
    LCurrent = R.EndColNo
    LLine = LLine + 1
    LCurrent = LCurrent + 1
End Sub

Private Sub AddIntellWord()
Dim R As CodeSenseCtl.Range
    If sIntellText <> "" Then
        Set R = RT.GetSel(False)
        R.StartColNo = LBoxPos
        R.EndColNo = R.EndColNo
        RT.SetSel R, False
        R.StartColNo = R.EndColNo + Len(sIntellText)
        R.EndColNo = R.EndColNo + Len(sIntellText)
        RT.SelText = sIntellText
        RT.SetSel R, False

        sIntellText = ""
    End If
End Sub
