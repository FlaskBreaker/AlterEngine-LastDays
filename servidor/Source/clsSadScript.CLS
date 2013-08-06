VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSadScript"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##############################
'# TIPOS
'##############################
Private Type SSDefine
    sVari As String
    sValue As String
End Type
Private Type SSClass
    sName As String
    sClassVars() As String
    sCodeBlock As Collection
End Type
Private Type SSModule
    sName As String
    sGlobals() As String
    sClasses() As SSClass
    sCodeBlock As Collection
End Type

'##############################
'# VARIABLES
'##############################
Public WithEvents SControl      As ScriptControl
Attribute SControl.VB_VarHelpID = -1
Private sAllCode()              As String
Private sModule              As SSModule
Private sIncludes()          As String

'##############################
'# PROCEDIMIENTOS Y FUNCIONES
'##############################
Public Function ReadInCode(sFile As String, sModTitle As String, msc As ScriptControl)
    Dim sTempCode As String
    sTempCode = ParseCompSeq(sFile)
    sModule.sName = sModTitle
    sAllCode = Split(sTempCode, vbNewLine)
    ReDim Preserve sAllCode(UBound(sAllCode) - 1)
    ParseCode sAllCode()
    msc.Modules.Add sModTitle
    AddCode msc
End Function
'Parsea los comandos de compilador: #loquesea
Private Function ParseCompSeq(sFile As String)
    Dim sTemp As String
    Dim iTemp As Integer
    Dim sTotalTemp As String
    Dim iFileNum As Integer
    Dim sDefines() As SSDefine
    Dim Path As String
    Dim AlreadyIncluded As Boolean
    
    Path = App.Path & "\Scripts"
    
    ReDim sIncludes(0)
    ReDim sDefines(0)
    
    sIncludes(0) = ""
    
    iFileNum = FreeFile
    
    Open sFile For Input As iFileNum
    Do Until EOF(iFileNum)
        Line Input #iFileNum, sTemp

        sTemp = Trim$(Replace(sTemp, vbTab, vbNullString))

        If LenB(sTemp) <> 0 Then
            If Left$(sTemp, 1) <> "#" Then
                sTotalTemp = sTotalTemp & sTemp & vbNewLine
            Else
                If LCase$(Left$(sTemp, 8)) = "#include" Then
                    sTemp = Mid$(sTemp, InStr(sTemp, "<") + 1, Len(sTemp) - InStr(sTemp, "<") - 1)
                    Dim i As Integer
                    For i = 0 To UBound(sIncludes)
                        If sIncludes(i) = LCase(sTemp) Then
                            AlreadyIncluded = True
                            Exit For
                        End If
                    Next i
                    If Not AlreadyIncluded Then
                        sIncludes(UBound(sIncludes)) = LCase(sTemp)
                        ReDim Preserve sIncludes(UBound(sIncludes) + 1)
                        sTemp = sTemp & vbNewLine & ParseCompSeq(Path & "\" & sTemp)
                        sTotalTemp = sTotalTemp & vbNewLine & sTemp
                    End If
                ElseIf LCase$(Left$(sTemp, 7)) = "#define" Then
                    sTemp = Right$(sTemp, Len(sTemp) - 8)
                    sDefines(UBound(sDefines)).sVari = Mid$(sTemp, 2, InStr(sTemp, "> <") - 2)
                    sDefines(UBound(sDefines)).sValue = Mid$(sTemp, InStr(sTemp, "> <") + 3, Len(sTemp) - InStr(sTemp, "> <") - 3)
                    ReDim Preserve sDefines(UBound(sDefines) + 1)
                End If
            End If
        End If
    Loop

    Close iFileNum
    
    If UBound(sDefines) <> 0 Then
        ReDim Preserve sDefines(UBound(sDefines) - 1)
    End If
    
    For iTemp = 0 To UBound(sDefines)
        sTotalTemp = Replace(sTotalTemp, sDefines(iTemp).sVari, sDefines(iTemp).sValue)
    Next
    ParseCompSeq = sTotalTemp
End Function
'Parsea el código
Private Sub ParseCode(sCode() As String)
    Dim iCount As Integer
    Dim iTemp As Integer
    Dim sTitle As String
    Dim sClass As SSClass
    Dim sCB As String
    
    ReDim sModule.sGlobals(0)
    ReDim sModule.sClasses(0)
    Set sModule.sCodeBlock = New Collection
    
    ReDim sClass.sClassVars(0)
    Set sClass.sCodeBlock = New Collection

    For iCount = 0 To UBound(sCode)
        sCB = ""
        If LCase$(Left$(sCode(iCount), 6)) = "class " Then 'Si es una clase
            sClass.sName = Mid$(sCode(iCount), 7, Len(sCode(iCount)) - 6)
'            For iTemp = 7 To Len(sCode(iCount)) 'Buscamos el nombre de la misma
'                If Mid$(sCode(iCount), iTemp, 1) = " " Then
'                    sClass.sName = Mid$(sCode(iCount), 7, iTemp - 7)
'                    Exit For
'                End If
'            Next

            Do Until LCase$(sCode(iCount)) = "end class"
                sCB = ""
                If (LCase$(Left$(sCode(iCount), 9)) = "function ") Or (LCase$(Left$(sCode(iCount), 4)) = "sub ") Then 'Procedimiento o función en una clase
                    If (LCase$(Left$(sCode(iCount), 9)) = "function ") Then 'Función en una clase
                        For iTemp = 10 To Len(sCode(iCount)) 'Buscamos el nombre
                            If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                                sTitle = Mid$(sCode(iCount), 10, iTemp - 10)
                                Exit For
                            End If
                        Next
            
                        Do Until LCase$(sCode(iCount)) = "end function" 'Conseguimos el código
                            sCB = sCB & sCode(iCount) & vbNewLine
                            iCount = iCount + 1
                        Loop
                        sCB = sCB & sCode(iCount)
                       'On Error Resume Next
                        sClass.sCodeBlock.Add sCB, sTitle
                        Err.Clear
                    ElseIf (LCase$(Left$(sCode(iCount), 4)) = "sub ") Then 'Procedimiento en una clase
                        For iTemp = 5 To Len(sCode(iCount)) 'Buscamos el nombre
                            If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                                sTitle = Mid$(sCode(iCount), 5, iTemp - 5)
                                Exit For
                            End If
                        Next
            
                        Do Until LCase$(sCode(iCount)) = "end sub" 'Conseguimos el código
                            sCB = sCB & sCode(iCount) & vbNewLine
                            iCount = iCount + 1
                        Loop
                        sCB = sCB & sCode(iCount)
                       'On Error Resume Next
                        sClass.sCodeBlock.Add sCB, sTitle
                        Err.Clear
                    End If
                ElseIf (LCase$(Left$(sCode(iCount), 7)) = "public ") Or (LCase$(Left$(sCode(iCount), 8)) = "private ") Or (LCase$(Left$(sCode(iCount), 4)) = "dim ") Then 'Global en una clase
                    If LCase$(Left$(sCode(iCount), 7)) = "public " Then
                        If Not ((LCase$(Left$(sCode(iCount), 16)) = "public function ") Or (LCase$(Left$(sCode(iCount), 11)) = "public sub ")) Then 'Se nos coló una función o un sub
                            If (LCase$(Left$(sCode(iCount), 16)) = "public function ") Then 'Función pública en una clase
                                For iTemp = 17 To Len(sCode(iCount)) 'Buscamos el nombre
                                    If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                                        sTitle = Mid$(sCode(iCount), 17, iTemp - 17)
                                        Exit For
                                    End If
                                Next
                    
                                Do Until LCase$(sCode(iCount)) = "end function" 'Conseguimos el código
                                    sCB = sCB & sCode(iCount) & vbNewLine
                                    iCount = iCount + 1
                                Loop
                                sCB = sCB & sCode(iCount)
                               'On Error Resume Next
                                sClass.sCodeBlock.Add sCB, sTitle
                                Err.Clear
                            ElseIf (LCase$(Left$(sCode(iCount), 11)) = "public sub ") Then 'Procedimiento público en una clase
                                For iTemp = 12 To Len(sCode(iCount)) 'Buscamos el nombre
                                    If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                                        sTitle = Mid$(sCode(iCount), 12, iTemp - 12)
                                        Exit For
                                    End If
                                Next
                    
                                Do Until LCase$(sCode(iCount)) = "end sub" 'Conseguimos el código
                                    sCB = sCB & sCode(iCount) & vbNewLine
                                    iCount = iCount + 1
                                Loop
                                sCB = sCB & sCode(iCount)
                               'On Error Resume Next
                                sClass.sCodeBlock.Add sCB, sTitle
                                Err.Clear
                            End If
                        Else 'public
                            sClass.sClassVars(UBound(sClass.sClassVars)) = sCode(iCount)
                            ReDim Preserve sClass.sClassVars(UBound(sClass.sClassVars) + 1)
                        End If
                    ElseIf LCase$(Left$(sCode(iCount), 8)) = "private " Then
                        If Not ((LCase$(Left$(sCode(iCount), 17)) = "private function ") Or (LCase$(Left$(sCode(iCount), 12)) = "private sub ")) Then 'Se nos coló una función o un sub
                            If (LCase$(Left$(sCode(iCount), 17)) = "private function ") Then 'Función privada en una clase
                                For iTemp = 18 To Len(sCode(iCount)) 'Buscamos el nombre
                                    If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                                        sTitle = Mid$(sCode(iCount), 18, iTemp - 18)
                                        Exit For
                                    End If
                                Next
                    
                                Do Until LCase$(sCode(iCount)) = "end function" 'Conseguimos el código
                                    sCB = sCB & sCode(iCount) & vbNewLine
                                    iCount = iCount + 1
                                Loop
                                sCB = sCB & sCode(iCount)
                               'On Error Resume Next
                                sClass.sCodeBlock.Add sCB, sTitle
                                Err.Clear
                            ElseIf (LCase$(Left$(sCode(iCount), 12)) = "private sub ") Then 'Procedimiento privado en una clase
                                For iTemp = 13 To Len(sCode(iCount)) 'Buscamos el nombre
                                    If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                                        sTitle = Mid$(sCode(iCount), 13, iTemp - 13)
                                        Exit For
                                    End If
                                Next
                    
                                Do Until LCase$(sCode(iCount)) = "end sub" 'Conseguimos el código
                                    sCB = sCB & sCode(iCount) & vbNewLine
                                    iCount = iCount + 1
                                Loop
                                sCB = sCB & sCode(iCount)
                               'On Error Resume Next
                                sClass.sCodeBlock.Add sCB, sTitle
                                Err.Clear
                            End If
                        Else 'private
                            sClass.sClassVars(UBound(sClass.sClassVars)) = sCode(iCount)
                            ReDim Preserve sClass.sClassVars(UBound(sClass.sClassVars) + 1)
                        End If
                    ElseIf LCase$(Left$(sCode(iCount), 4)) = "dim " Then 'dim
                        sClass.sClassVars(UBound(sClass.sClassVars)) = sCode(iCount)
                        ReDim Preserve sClass.sClassVars(UBound(sClass.sClassVars) + 1)
                    End If
                End If
                iCount = iCount + 1
            Loop
            
            If UBound(sClass.sClassVars) <> 0 Then
                ReDim Preserve sClass.sClassVars(UBound(sClass.sClassVars) - 1)
            End If
           'On Error Resume Next

            sModule.sClasses(UBound(sModule.sClasses)) = sClass
            ReDim Preserve sModule.sClasses(UBound(sModule.sClasses) + 1)
            Err.Clear
            'Reseteo de Class
            ReDim sClass.sClassVars(0)
            Set sClass.sCodeBlock = New Collection
        ElseIf (LCase$(Left$(sCode(iCount), 9)) = "function ") Or (LCase$(Left$(sCode(iCount), 4)) = "sub ") Then
            If (LCase$(Left$(sCode(iCount), 9)) = "function ") Then 'Función fuera de una clase
                For iTemp = 10 To Len(sCode(iCount)) 'Buscamos el nombre
                    If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                        sTitle = Mid$(sCode(iCount), 10, iTemp - 10)
                        Exit For
                    End If
                Next
    
                Do Until LCase$(sCode(iCount)) = "end function" 'Conseguimos el código
                    sCB = sCB & sCode(iCount) & vbNewLine
                    iCount = iCount + 1
                Loop
                sCB = sCB & sCode(iCount)
               'On Error Resume Next
                sModule.sCodeBlock.Add sCB, sTitle
                Err.Clear
            ElseIf (LCase$(Left$(sCode(iCount), 4)) = "sub ") Then 'Procedimiento fuera de una clase
                For iTemp = 5 To Len(sCode(iCount)) 'Buscamos el nombre
                    If Mid$(sCode(iCount), iTemp, 1) = "(" Then
                        sTitle = Mid$(sCode(iCount), 5, iTemp - 5)
                        Exit For
                    End If
                Next
    
                Do Until LCase$(sCode(iCount)) = "end sub" 'Conseguimos el código
                    sCB = sCB & sCode(iCount) & vbNewLine
                    iCount = iCount + 1
                Loop
                sCB = sCB & sCode(iCount)
               'On Error Resume Next
                sModule.sCodeBlock.Add sCB, sTitle
                Err.Clear
            End If
        ElseIf (LCase$(Left$(sCode(iCount), 7)) = "public ") Or (LCase$(Left$(sCode(iCount), 8)) = "private ") Then
            sModule.sGlobals(UBound(sModule.sGlobals)) = sCode(iCount)
            ReDim Preserve sModule.sGlobals(UBound(sModule.sGlobals) + 1)
        End If
    Next iCount
    If UBound(sModule.sGlobals) <> 0 Then
        ReDim Preserve sModule.sGlobals(UBound(sModule.sGlobals) - 1)
    End If
    If UBound(sModule.sClasses) <> 0 Then
        ReDim Preserve sModule.sClasses(UBound(sModule.sClasses) - 1)
    End If
End Sub
'Añade el código al ScriptControl
Private Sub AddCode(mscControl As ScriptControl)
    Dim i As Integer
    Dim j As Integer
    Dim aClass As SSClass
    Dim tmpClass As String
   'On Error Resume Next
   If sModule.sGlobals(0) <> "" Then
        For i = 0 To UBound(sModule.sGlobals)
            mscControl.Modules(sModule.sName).AddCode sModule.sGlobals(i)
        Next i
    End If
   i = 0
    If sModule.sClasses(0).sName <> "" Then
        For i = 0 To UBound(sModule.sClasses)
            aClass = sModule.sClasses(i)
            tmpClass = tmpClass & vbNewLine & "Class " & aClass.sName
            If aClass.sClassVars(0) <> "" Then
                For j = 0 To UBound(aClass.sClassVars)
                    tmpClass = tmpClass & vbNewLine & aClass.sClassVars(j)
                Next j
            End If
            j = 0
            For j = 1 To aClass.sCodeBlock.Count
                tmpClass = tmpClass & vbNewLine & aClass.sCodeBlock(j)
            Next j
            tmpClass = tmpClass & vbNewLine & "End Class"
            mscControl.Modules(sModule.sName).AddCode tmpClass
            tmpClass = ""
        Next i
    End If
    i = 0
    For i = 1 To sModule.sCodeBlock.Count
        mscControl.Modules(sModule.sName).AddCode sModule.sCodeBlock(i)
    Next i
End Sub
Private Sub Class_Initialize()
    Set SControl = New ScriptControl
    SControl.Language = "VBScript"
End Sub

Private Sub Class_Terminate()
    Set SControl = Nothing
End Sub

'PRIVATE BECAUSE I CANT GET IT TO WORK
Private Function RunCodeReturn(sModule As String, sCode As String, ParamArray abc() As Variant) As Variant
    Dim sError As String
    Dim aTemp() As String
    Dim bTemp As Variant

    'How the heck to pass a paramarray array to another function using a
    'paramarray array?? if ya know email me andyd@vuetura.com

    If IsMissing(abc) Then
        bTemp = SControl.Modules(sModule).Run(sCode)
    Else
        bTemp = SControl.Modules(sModule).Run(sCode, abc)
    End If

        RunCodeReturn = bTemp
End Function

Public Function ExecuteStatement(sModule As String, sCode As String)
    'On Error Resume Next

    SControl.Modules(sModule).ExecuteStatement sCode
End Function

Private Sub SControl_Error()
    Dim sError As String

    If Val(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ScriptErrors")) = 1 Then
        MsgBox "TYPE: " & SControl.Error.Description & vbNewLine & "LINE: " & SControl.Error.Line & vbNewLine & "COLUMN: " & SControl.Error.Column & vbNewLine & "CODE: " & vbNewLine & "    " & SControl.Error.Text
    End If

    Err.Clear
End Sub

