Attribute VB_Name = "mdCodeLines"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long

Private Const STR_ATTRIB_START      As String = "Attribute VB_Name "
Private Const STR_ATTRIB            As String = "Attribute "
Private Const STR_PRIVATE           As String = "Private "
Private Const STR_PUBLIC            As String = "Public "
Private Const STR_FRIEND            As String = "Friend "
Private Const STR_STATIC            As String = "Static "
Private Const STR_SUB               As String = "Sub "
Private Const STR_FUNCTION          As String = "Function "
Private Const STR_PROPERTY          As String = "Property "
Private Const STR_END               As String = "End "
Private Const STR_SELECT_CASE       As String = "Select Case "
Private Const STR_CASE              As String = "Case "
Private Const STR_AS                As String = "As "
Private Const STR_STOP              As String = "{stop_code_lines}"
Private Const STR_ON_ERROR          As String = "On Error GoTo"

Private Enum UcsStageEnum
    ucsSearchAttribStart
    ucsSearchAttribEnd
    ucsSearchProcStart
    ucsSearchProcStartContinue
    ucsSearchProcEnd
    ucsSearchProcEndContinue
    ucsSearchFirstCase
    ucsSearchFirstCaseContinue
    ucsSearchSelectCaseContinue
    ucsSearchOnError
End Enum

Public Function ProcessProject( _
            ByVal sVbpFileName As String, _
            Optional lProcessedLines As Long) As Long
    Dim oVbp            As New cVbpFile
    Dim cSources        As Collection
    Dim lIdx            As Long
    
    On Error GoTo EH
    oVbp.Init sVbpFileName
    Set cSources = oVbp.Sources
    lProcessedLines = 0
    For lIdx = 1 To cSources.Count
        lProcessedLines = lProcessedLines + pvProcessFile(oVbp.GetSourceFileName(cSources(lIdx)))
        ProcessProject = ProcessProject + 1
    Next
    Exit Function
EH:
    MsgAlert Error
End Function

Private Function pvProcessFile(sSourceFile As String) As Long
    Dim iFile           As Integer
    Dim cLines          As Variant
    Dim lIdx            As Long
    Dim sLine           As String
    Dim lStage          As UcsStageEnum
    Dim lCurrentLine    As Long
    Dim lLen            As Long
    Dim sTrimLine       As String
    Dim vNonCodeLines   As Variant
    
    On Error GoTo EH
    vNonCodeLines = Split("Dim|Const|Static|On Error|Else|End If|Goto|Resume|Case|End Select|Loop|Next|End With", "|")
    '--- read source file
    cLines = Split(ReadTextFile(sSourceFile), vbCrLf)
    '--- process
    lCurrentLine = 1
    lStage = ucsSearchAttribStart
    For lIdx = 1 To UBound(cLines) + 1
        sLine = cLines(lIdx - 1)
        sTrimLine = Trim$(sLine)
        If lStage = ucsSearchAttribStart Then '--- searching for attrib start
            If Left$(sLine, Len(STR_ATTRIB_START)) = STR_ATTRIB_START Then
                lStage = ucsSearchAttribEnd
            End If
        End If
        If lStage = ucsSearchAttribEnd Then '--- searching for attrib end
            If Left$(sLine, Len(STR_ATTRIB)) <> STR_ATTRIB Then
                lStage = ucsSearchProcStart
            End If
        End If
        If lStage >= ucsSearchProcStart Then
            '--- skip procedure attributes and don't increment current line
            If StrComp(Left$(sTrimLine, Len(STR_ATTRIB)), STR_ATTRIB, vbTextCompare) = 0 Then
                GoTo LoopNext
            End If
            If StrComp(Right$(sTrimLine, Len(STR_STOP)), STR_STOP, vbTextCompare) = 0 Then
                Exit For
            End If
            If lStage = ucsSearchProcStart Then
                '--- searching procedure start
                If StrComp(Left$(sTrimLine, Len(STR_PRIVATE)), STR_PRIVATE, vbTextCompare) = 0 Then
                    sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_PRIVATE)))
                ElseIf StrComp(Left$(sTrimLine, Len(STR_PUBLIC)), STR_PUBLIC, vbTextCompare) = 0 Then
                    sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_PUBLIC)))
                ElseIf StrComp(Left$(sTrimLine, Len(STR_FRIEND)), STR_FRIEND, vbTextCompare) = 0 Then
                    sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_FRIEND)))
                ElseIf StrComp(Left$(sTrimLine, Len(STR_STATIC)), STR_STATIC, vbTextCompare) = 0 Then
                    sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_STATIC)))
                End If
                If StrComp(Left$(sTrimLine, Len(STR_SUB)), STR_SUB, vbTextCompare) = 0 _
                        Or StrComp(Left$(sTrimLine, Len(STR_FUNCTION)), STR_FUNCTION, vbTextCompare) = 0 _
                        Or StrComp(Left$(sTrimLine, Len(STR_FUNCTION)), STR_PROPERTY, vbTextCompare) = 0 Then
                    sTrimLine = Trim$(Mid$(sTrimLine, InStr(sTrimLine, " ")))
                    If StrComp(Left$(sTrimLine, Len(STR_AS)), STR_AS, vbTextCompare) <> 0 Then
                        lStage = ucsSearchOnError
                    End If
                End If
                If Right$(sTrimLine, 2) = " _" Then
                    '--- start continuation
                    If lStage = ucsSearchProcStart Then
                        lStage = ucsSearchProcStartContinue
                    ElseIf lStage = ucsSearchProcEnd Then
                        lStage = ucsSearchProcEndContinue
                    End If
                End If
                lCurrentLine = lCurrentLine + 1
                GoTo LoopNext
            End If
            If lStage = ucsSearchProcStartContinue Then
                If Right$(sTrimLine, 2) <> " _" Then
                    '--- end of continuation
                    lStage = ucsSearchProcStart
                End If
                lCurrentLine = lCurrentLine + 1
            End If
            If lStage = ucsSearchProcEnd Then
                '--- searching procedure end
                If StrComp(Left$(sTrimLine, Len(STR_END)), STR_END, vbTextCompare) = 0 Then
                    sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_END))) & " "
                    If StrComp(Left$(sTrimLine, Len(STR_SUB)), STR_SUB, vbTextCompare) = 0 _
                            Or StrComp(Left$(sTrimLine, Len(STR_FUNCTION)), STR_FUNCTION, vbTextCompare) = 0 _
                            Or StrComp(Left$(sTrimLine, Len(STR_FUNCTION)), STR_PROPERTY, vbTextCompare) = 0 Then
                        lStage = ucsSearchProcStart
                    End If
                End If
                If lStage = ucsSearchProcEnd Then
                    If StrComp(Left$(sTrimLine, Len(STR_PRIVATE)), STR_PRIVATE, vbTextCompare) = 0 Then
                        sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_PRIVATE)))
                    ElseIf StrComp(Left$(sTrimLine, Len(STR_PUBLIC)), STR_PUBLIC, vbTextCompare) = 0 Then
                        sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_PUBLIC)))
                    ElseIf StrComp(Left$(sTrimLine, Len(STR_FRIEND)), STR_FRIEND, vbTextCompare) = 0 Then
                        sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_FRIEND)))
                    ElseIf StrComp(Left$(sTrimLine, Len(STR_STATIC)), STR_STATIC, vbTextCompare) = 0 Then
                        sTrimLine = Trim$(Mid$(sTrimLine, Len(STR_STATIC)))
                    End If
                    If StrComp(Left$(sTrimLine, Len(STR_SUB)), STR_SUB, vbTextCompare) = 0 _
                            Or StrComp(Left$(sTrimLine, Len(STR_FUNCTION)), STR_FUNCTION, vbTextCompare) = 0 _
                            Or StrComp(Left$(sTrimLine, Len(STR_FUNCTION)), STR_PROPERTY, vbTextCompare) = 0 Then
                        sTrimLine = vbNullString
                    End If
                    '--- skip empty lines and preprocessor directives
                    If sTrimLine <> "" And Left$(sTrimLine, 1) <> "#" And Left$(sTrimLine, 1) <> "'" And Right$(sTrimLine, 1) <> ":" Then
                        '--- clean previous line numbers (and next colon/space)
                        If Val(sLine) <> 0 Then
                            lLen = Len(CStr(Val(sLine))) + 2
                            sLine = Mid$(sLine, lLen)
                        End If
                        If Not pvBeginsWith(Trim$(sLine), vNonCodeLines) Then
                            '--- add current line number
                            sLine = lCurrentLine & " " & sLine
                            '--- change current line in collection
                            cLines(lIdx - 1) = sLine
                            '--- count processed lines
                            pvProcessFile = pvProcessFile + 1
                        End If
                    End If
                    '--- special treatment of "select case"
                    If InStr(1, sLine, STR_SELECT_CASE, vbTextCompare) > 0 Then
                        If InStrRev(sLine, "'", InStr(1, sLine, STR_SELECT_CASE, vbTextCompare)) = 0 Then
                            lStage = ucsSearchFirstCase
                        End If
                    End If
                End If
                If Right$(Trim$(sLine), 2) = " _" Then
                    '--- start continuation
                    If lStage = ucsSearchProcEnd Then
                        lStage = ucsSearchProcEndContinue
                    ElseIf lStage = ucsSearchFirstCase Then
                        lStage = ucsSearchSelectCaseContinue
                    End If
                End If
                lCurrentLine = lCurrentLine + 1
                GoTo LoopNext
            End If
            If lStage = ucsSearchProcEndContinue Then
                If Right$(Trim$(sLine), 2) <> " _" Then
                    '--- end of continuation
                    lStage = ucsSearchProcEnd
                End If
                lCurrentLine = lCurrentLine + 1
            End If
            If lStage = ucsSearchFirstCaseContinue Then
                If Right$(Trim$(sLine), 2) <> " _" Then
                    '--- end of continuation
                    lStage = ucsSearchProcEnd
                End If
                lCurrentLine = lCurrentLine + 1
            End If
            If lStage = ucsSearchFirstCase Then
                If StrComp(Left$(sTrimLine, Len(STR_CASE)), STR_CASE, vbTextCompare) = 0 Then
                    If Right$(Trim$(sLine), 2) <> " _" Then
                        '--- end of continuation
                        lStage = ucsSearchProcEnd
                    Else
                        lStage = ucsSearchFirstCaseContinue
                    End If
                End If
                lCurrentLine = lCurrentLine + 1
            End If
            If lStage = ucsSearchSelectCaseContinue Then
                If Right$(Trim$(sLine), 2) <> " _" Then
                    '--- end of continuation
                    lStage = ucsSearchFirstCase
                End If
                lCurrentLine = lCurrentLine + 1
            End If
            If lStage = ucsSearchOnError Then
                If StrComp(Left$(sTrimLine, Len(STR_ON_ERROR)), STR_ON_ERROR, vbTextCompare) = 0 Then
                    lStage = ucsSearchProcEnd
                End If
                lCurrentLine = lCurrentLine + 1
            End If
        End If
LoopNext:
    Next
    '--- save source file
    SetAttr sSourceFile, vbArchive
    WriteTextFile sSourceFile, Join(cLines, vbCrLf)
    Exit Function
EH:
    MsgAlert Error
    Resume NextLine
NextLine:
    On Error Resume Next
    Close iFile
End Function

Private Function pvBeginsWith(sText As String, vMatch As Variant) As Boolean
    Dim vElem           As Variant
    
    For Each vElem In vMatch
        If StrComp(Left$(sText, Len(vElem)), vElem, vbTextCompare) = 0 Then
            If Len(sText) = Len(vElem) Then
                pvBeginsWith = True
                Exit Function
            End If
            If Mid$(sText, Len(vElem) + 1, 1) = " " Then
                pvBeginsWith = True
                Exit Function
            End If
        End If
    Next
End Function

Public Sub Main()
    Dim vElem           As Variant
    
    If Command <> "" Then
        g_bNoUI = True
        For Each vElem In SplitArgs(Command$)
            ProcessProject vElem, 0
        Next
    Else
        frmMain.Show
    End If
End Sub

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function
