Attribute VB_Name = "modGeneral"
Option Explicit

' Client Executes Here.
Public Sub Main()

    Call SetStatus("Cargando Sonido...")
    
    ' change and set the current path, to prevent from VB not finding BASS.DLL
    ChDrive App.Path
    ChDir App.Path

    ' check the correct BASS was loaded
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
        End
    End If

    ' Initialize output - default device, 44100hz, stereo, 16 bits
    If BASS_Init(-1, 44100, 0, frmMirage.hWnd, 0) = BASSFALSE Then
        Call Error_("Can't initialize digital sound system")
        End
    End If
    
    If FileExists("debug") Then
        frmDebug.Visible = True
    End If

    frmSendGetData.Visible = True

    ' Check to make sure all the folder exist.
    Call SetStatus("Comprobando Carpetas...")
    Call CheckFolders

    ' Check to make sure all the files exist.
    Call SetStatus("Comprobando Archivos...")
    Call SystemFileChecker
    
    If Not FileExists("config.ini") Then
        Call FileCreateConfigINI
    End If

    If Not FileExists("Noticias.ini") Then
        Call FileCreateNewsINI
    End If

    If Not FileExists("Fuente.ini") Then
        Call FileCreateFontINI
    End If

    If Not FileExists("GUI\Colors.txt") Then
        Call FileCreateColorsTXT
    End If
    
    ' Initialize global variables.
    LAST_DIR = -1

    ' Load the configuration settings.
    Call SetStatus("Cargando Configuración...")
    Call LoadConfig
    Call LoadColors
    Call LoadFont
    
    InQuestEditor = False
    ' Prepare the socket for communication.
    Call SetStatus("Preparando conexión...")
    Call TcpInit

    'frmMainMenu.lblVersion.Caption = "Version: " & App.Major & "." & App.Minor

    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
End Sub

Public Sub Error_(ByVal es As String)
    Call MsgBox(es & vbCrLf & "(error code: " & BASS_ErrorGetCode() & ")", vbExclamation, "Error")
End Sub

Private Sub CheckFolders()

    If LCase$(Dir$(App.Path & "\Mapas", vbDirectory)) <> "mapas" Then
        Call MkDir$(App.Path & "\Mapas")
    End If

    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir$(App.Path & "\GFX")
    End If

    If UCase$(Dir$(App.Path & "\GUI", vbDirectory)) <> "GUI" Then
        Call MkDir$(App.Path & "\GUI")
    End If

    If UCase$(Dir$(App.Path & "\Musica", vbDirectory)) <> "MUSICA" Then
        Call MkDir$(App.Path & "\Musica")
    End If

    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir$(App.Path & "\SFX")
    End If

    If UCase$(Dir$(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then
        Call MkDir$(App.Path & "\Flashs")
    End If

    If UCase$(Dir$(App.Path & "\BGS", vbDirectory)) <> "BGS" Then
        Call MkDir$(App.Path & "\BGS")
    End If

    If UCase$(Dir$(App.Path & "\DATA", vbDirectory)) <> "DATA" Then
        Call MkDir$(App.Path & "\Data")
    End If

End Sub

Private Sub LoadConfig()
    Dim filename As String

    On Error GoTo ErrorHandle

    filename = App.Path & "\config.ini"

    frmMirage.chkBubbleBar.value = CLng(ReadINI("CONFIG", "SpeechBubbles", filename))
    frmMirage.chkNpcBar.value = CLng(ReadINI("CONFIG", "NpcBar", filename))
    frmMirage.chkNpcName.value = CLng(ReadINI("CONFIG", "NPCName", filename))
    frmMirage.chkPlayerBar.value = CLng(ReadINI("CONFIG", "PlayerBar", filename))
    frmMirage.chkPlayerName.value = CLng(ReadINI("CONFIG", "PlayerName", filename))
    frmMirage.chkPlayerDamage.value = CLng(ReadINI("CONFIG", "NPCDamage", filename))
    frmMirage.chkNpcDamage.value = ReadINI("CONFIG", "PlayerDamage", filename)
   ' frmMirage.chkMusic.Value = CLng(ReadINI("CONFIG", "Music", FileName)) <-- This caused connectivity issues upon disabling music [Devil Of Duce]
    frmMirage.chkSound.value = CLng(ReadINI("CONFIG", "Sound", filename))
    frmMirage.chkAutoScroll.value = CLng(ReadINI("CONFIG", "AutoScroll", filename))
    AutoLogin = CLng(ReadINI("CONFIG", "Auto", filename))

    Exit Sub

ErrorHandle:
    Call MsgBox("Error cargando config. Recreando config.ini.")
    Kill "config.ini"
    Call FileCreateConfigINI

End Sub

Private Sub FileCreateConfigINI()
    WriteINI "IPCONFIG", "IP", "127.0.0.1", App.Path & "\config.ini"
    WriteINI "IPCONFIG", "PORT", 4001, App.Path & "\config.ini"

    WriteINI "CONFIG", "Account", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "Password", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "WebSite", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "SpeechBubbles", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "NpcBar", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "NPCName", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "NPCDamage", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerBar", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerName", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerDamage", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "MapGrid", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "Music", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "Sound", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "AutoScroll", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "Auto", 0, App.Path & "\config.ini"
End Sub

Private Sub FileCreateNewsINI()
    WriteINI "DATA", "News", vbNullString, App.Path & "\Noticias.ini"
    WriteINI "DATA", "Desc", vbNullString, App.Path & "\Noticias.ini"

    WriteINI "COLOR", "Red", 255, App.Path & "\Noticias.ini"
    WriteINI "COLOR", "Green", 255, App.Path & "\Noticias.ini"
    WriteINI "COLOR", "Blue", 255, App.Path & "\Noticias.ini"

    WriteINI "FONT", "Font", "Arial", App.Path & "\Noticias.ini"
    WriteINI "FONT", "Size", "14", App.Path & "\Noticias.ini"
End Sub

Private Sub FileCreateFontINI()
    Call WriteINI("FONT", "Font", "fixedsys", App.Path & "\Fuente.ini")
    Call WriteINI("FONT", "Size", 18, App.Path & "\Fuente.ini")
End Sub

Private Sub LoadColors()
    Dim r1 As Long
    Dim g1 As Long
    Dim b1 As Long

    On Error GoTo ErrorHandle

    ' chat box color
    r1 = CInt(ReadINI("CHATBOX", "R", App.Path & "\GUI\Colors.txt"))
    g1 = CInt(ReadINI("CHATBOX", "G", App.Path & "\GUI\Colors.txt"))
    b1 = CInt(ReadINI("CHATBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtChat.BackColor = RGB(r1, g1, b1)

    ' chat box text color
    r1 = CInt(ReadINI("CHATTEXTBOX", "R", App.Path & "\GUI\Colors.txt"))
    g1 = CInt(ReadINI("CHATTEXTBOX", "G", App.Path & "\GUI\Colors.txt"))
    b1 = CInt(ReadINI("CHATTEXTBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtMyTextBox.BackColor = RGB(r1, g1, b1)

    r1 = CInt(ReadINI("SPELLLIST", "R", App.Path & "\GUI\Colors.txt"))
    g1 = CInt(ReadINI("SPELLLIST", "G", App.Path & "\GUI\Colors.txt"))
    b1 = CInt(ReadINI("SPELLLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstSpells.BackColor = RGB(r1, g1, b1)

    r1 = CInt(ReadINI("WHOLIST", "R", App.Path & "\GUI\Colors.txt"))
    g1 = CInt(ReadINI("WHOLIST", "G", App.Path & "\GUI\Colors.txt"))
    b1 = CInt(ReadINI("WHOLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstOnline.BackColor = RGB(r1, g1, b1)

    r1 = CInt(ReadINI("NEWCHAR", "R", App.Path & "\GUI\Colors.txt"))
    g1 = CInt(ReadINI("NEWCHAR", "G", App.Path & "\GUI\Colors.txt"))
    b1 = CInt(ReadINI("NEWCHAR", "B", App.Path & "\GUI\Colors.txt"))
    frmNewChar.optMale.BackColor = RGB(r1, g1, b1)
    frmNewChar.optFemale.BackColor = RGB(r1, g1, b1)

    r1 = CInt(ReadINI("BACKGROUND", "R", App.Path & "\GUI\Colors.txt"))
    g1 = CInt(ReadINI("BACKGROUND", "G", App.Path & "\GUI\Colors.txt"))
    b1 = CInt(ReadINI("BACKGROUND", "B", App.Path & "\GUI\Colors.txt"))

    frmMirage.picInventory3.BackColor = RGB(r1, g1, b1)
    frmMirage.picInventory.BackColor = RGB(r1, g1, b1)
    frmMirage.itmDesc.BackColor = RGB(r1, g1, b1)
    frmMirage.picWhosOnline.BackColor = RGB(r1, g1, b1)
    frmMirage.picGuildAdmin.BackColor = RGB(r1, g1, b1)
    frmMirage.picGuildMember.BackColor = RGB(r1, g1, b1)
    frmMirage.picEquipment.BackColor = RGB(r1, g1, b1)
    frmMirage.picPlayerSpells.BackColor = RGB(r1, g1, b1)
    frmMirage.picOptions.BackColor = RGB(r1, g1, b1)

    frmMirage.chkBubbleBar.BackColor = RGB(r1, g1, b1)
    frmMirage.chkNpcBar.BackColor = RGB(r1, g1, b1)
    frmMirage.chkNpcName.BackColor = RGB(r1, g1, b1)
    frmMirage.chkPlayerBar.BackColor = RGB(r1, g1, b1)
    frmMirage.chkPlayerName.BackColor = RGB(r1, g1, b1)
    frmMirage.chkPlayerDamage.BackColor = RGB(r1, g1, b1)
    frmMirage.chkNpcDamage.BackColor = RGB(r1, g1, b1)
    frmMirage.chkMusic.BackColor = RGB(r1, g1, b1)
    frmMirage.chkSound.BackColor = RGB(r1, g1, b1)
    frmMirage.chkAutoScroll.BackColor = RGB(r1, g1, b1)

    Exit Sub

ErrorHandle:
    Call MsgBox("Error loading colors.txt")

End Sub

Private Sub FileCreateColorsTXT()
    WriteINI "CHATBOX", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATBOX", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATBOX", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "CHATTEXTBOX", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATTEXTBOX", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATTEXTBOX", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "BACKGROUND", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "BACKGROUND", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "BACKGROUND", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "SPELLLIST", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "SPELLLIST", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "SPELLLIST", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "WHOLIST", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "WHOLIST", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "WHOLIST", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "NEWCHAR", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "NEWCHAR", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "NEWCHAR", "B", 120, App.Path & "\GUI\Colors.txt"
End Sub

Private Sub LoadFont()
    On Error GoTo ErrorHandle

    Font = ReadINI("FONT", "Font", App.Path & "\Fuente.ini")
    fontsize = CByte(ReadINI("FONT", "Size", App.Path & "\Fuente.ini"))

    If Font = vbNullString Then
        Font = "fixedsys"
    End If

    If fontsize <= 0 Or fontsize > 32 Then
        fontsize = 18
    End If

    Call SetFont(Font, fontsize)

    Exit Sub

ErrorHandle:
    Call WriteINI("FONT", "Font", "fixedsys", App.Path & "\Fuente.ini")
    Call WriteINI("FONT", "Size", 18, App.Path & "\Fuente.ini")

    Call SetFont("fixedsys", 18)

End Sub

' Función para poder utilizar arrays en los custom menus
Public Function IsInArray(CtrlArray As Variant, index As Integer) As Boolean
    On Error Resume Next
    Dim X As String
    X = CtrlArray(index).name
    If Err.Number = 0 Then IsInArray = True
End Function

Public Sub ClearSpells()
Dim I As Byte

For I = 1 To 20
Player(MyIndex).spellcdb(I) = False
Player(MyIndex).SpellcdTimer(I) = 0
Player(MyIndex).Spellpos(I) = 0
Next I
End Sub
