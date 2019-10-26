Attribute VB_Name = "mdCommon"
Option Explicit

Public Const STR_APPNAME            As String = "VB Code Lines"

Public g_bNoUI         As Boolean

Public Function FileExists(sFile As String) As Long
    FileExists = (Dir(sFile) <> "")
End Function

Public Function EscapeFile(sFile As String) As String
    EscapeFile = sFile
    If Len(EscapeFile) > 0 Then
        If Right(EscapeFile, 1) = "\" Then
            EscapeFile = Left(EscapeFile, Len(EscapeFile) - 1)
        End If
    End If
    If InStr(1, EscapeFile, " ") > 0 Then
        EscapeFile = """" & EscapeFile & """"
    End If
End Function

Public Sub MsgAlert(sText As String)
    If Not g_bNoUI Then
        MsgBox sText, vbExclamation, STR_APPNAME
    End If
End Sub

Public Function MsgConfirm(sText As String) As Boolean
    If Not g_bNoUI Then
        MsgConfirm = MsgBox(sText, vbQuestion Or vbYesNo, STR_APPNAME) = vbYes
    End If
End Function

Public Function ReadTextFile(sFile As String) As String
    Dim lSize           As Long
    Dim nFile           As Integer
    
    On Error GoTo EH
    lSize = FileLen(sFile)
    nFile = FreeFile()
    Open sFile For Binary Access Read As nFile
    ReadTextFile = String$(lSize, 0)
    Get nFile, , ReadTextFile
    Close nFile
    Exit Function
EH:
    If nFile <> 0 Then
        Close nFile
    End If
End Function

Public Sub WriteTextFile(sFile As String, sText As String)
    Dim nFile           As Integer
    
    On Error GoTo EH
    nFile = FreeFile()
    Open sFile For Binary Access Write As nFile
    Put nFile, , sText
    Close nFile
    Exit Sub
EH:
    If nFile <> 0 Then
        Close nFile
    End If
End Sub
