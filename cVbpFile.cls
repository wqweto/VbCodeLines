VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVbpFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum UcsHKEYEnum
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
End Enum

'=========================================================================
' API
'=========================================================================

'--- Reg Key Security Options...
Private Const READ_CONTROL              As Long = &H20000
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_ALL_ACCESS            As Long = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                                                    KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                                                    KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'--- Reg Ket types
Private Const REG_SZ                    As Long = 1
Private Const REG_DWORD                 As Long = 4
'--- reg key error codes
Private Const ERROR_SUCCESS             As Long = 0
Private Const MAX_PATH                  As Long = 260

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, ByVal lpcbClass As Long, ByVal lpftLastWriteTime As Long) As Long

'=========================================================================
' Constants and variables
'=========================================================================

Private Const STR_REFERENCE         As String = "REFERENCE="
Private Const STR_OBJECT            As String = "OBJECT="
Private Const STR_USERCONTROL       As String = "USERCONTROL="
Private Const STR_FORM              As String = "FORM="
Private Const STR_RELATEDDOC        As String = "RELATEDDOC="
Private Const STR_MODULE            As String = "MODULE="
Private Const STR_DESIGNER          As String = "DESIGNER="
Private Const STR_CLASS             As String = "CLASS="
Private Const STR_PROPERTYPAGE      As String = "PROPERTYPAGE="
Private Const STR_TYPE              As String = "TYPE="
Private Const STR_BUILDEXE          As String = "EXENAME32="
Private Const STR_BUILDPATH         As String = "PATH32="
Private Const STR_COMPATIBLEEXE     As String = "COMPATIBLEEXE32="
Private Const STR_COMPATIBLEMODE    As String = "COMPATIBLEMODE="
Private Const STR_VERSIONCOMPATIBLE As String = "VERSIONCOMPATIBLE32="
Private Const STR_MAJORVER          As String = "MAJORVER="
Private Const STR_MINORVER          As String = "MINORVER="
Private Const STR_REVISIONVER       As String = "REVISIONVER="
Private Const STR_AUTOINCREMENTVER  As String = "AUTOINCREMENTVER="
Private Const STR_MISSING_REF       As String = "#MissingReference#"

Private m_oFileName             As cFileName
Private m_cLines                As Collection

'=========================================================================
' Properties
'=========================================================================

Property Get FileName() As cFileName
    Set FileName = m_oFileName
End Property

Property Get References() As Collection
    Dim lIdx            As Long
    Dim sLine           As String
    
    Set References = New Collection
    For lIdx = 1 To m_cLines.Count
        sLine = m_cLines(lIdx)
        '--- check if line specifies reference
        If UCase(Left(sLine, Len(STR_REFERENCE))) = STR_REFERENCE Then
            References.Add sLine
        End If
    Next
End Property

Property Get Objects() As Collection
    Dim lIdx            As Long
    Dim sLine           As String
    
    Set Objects = New Collection
    For lIdx = 1 To m_cLines.Count
        sLine = m_cLines(lIdx)
        If UCase(Left(sLine, Len(STR_OBJECT))) = STR_OBJECT Then
            Objects.Add sLine
        End If
    Next
End Property

Property Get Sources() As Collection
    Dim lIdx            As Long
    Dim sLine           As String
    
    Set Sources = New Collection
    For lIdx = 1 To m_cLines.Count
        sLine = m_cLines(lIdx)
        If UCase(Left(sLine, Len(STR_USERCONTROL))) = STR_USERCONTROL _
                Or UCase(Left(sLine, Len(STR_FORM))) = STR_FORM _
                Or UCase(Left(sLine, Len(STR_RELATEDDOC))) = STR_RELATEDDOC _
                Or UCase(Left(sLine, Len(STR_MODULE))) = STR_MODULE _
                Or UCase(Left(sLine, Len(STR_DESIGNER))) = STR_DESIGNER _
                Or UCase(Left(sLine, Len(STR_CLASS))) = STR_CLASS _
                Or UCase(Left(sLine, Len(STR_PROPERTYPAGE))) = STR_PROPERTYPAGE _
                    Then
            Sources.Add sLine
        End If
    Next
End Property

Property Get All() As Collection
    Set All = m_cLines
End Property

Property Get Type_() As String
    Type_ = GetValue(FindLine(STR_TYPE))
End Property

Property Let Type_(sValue As String)
    SetValue FindLine(STR_TYPE), sValue
End Property

Property Get BuildPath() As String
    BuildPath = GetStringValue(FindLine(STR_BUILDPATH))
End Property

Property Let BuildPath(sValue As String)
    SetStringValue FindLine(STR_BUILDPATH), sValue
End Property

Property Get BuildExe() As String
    BuildExe = GetStringValue(FindLine(STR_BUILDEXE))
End Property

Property Let BuildExe(sValue As String)
    SetStringValue FindLine(STR_BUILDEXE), sValue
End Property

Property Get CompatibleExe() As String
    CompatibleExe = GetStringValue(FindLine(STR_COMPATIBLEEXE))
End Property

Property Let CompatibleExe(sValue As String)
    SetStringValue FindLine(STR_COMPATIBLEEXE), sValue
End Property

Property Get CompatibleMode() As String
    CompatibleMode = GetStringValue(FindLine(STR_COMPATIBLEMODE))
End Property

Property Let CompatibleMode(sValue As String)
    SetStringValue FindLine(STR_COMPATIBLEMODE), sValue
End Property

Property Get VersionCompatible() As String
    VersionCompatible = GetStringValue(FindLine(STR_VERSIONCOMPATIBLE))
End Property

Property Let VersionCompatible(sValue As String)
    SetStringValue FindLine(STR_VERSIONCOMPATIBLE), sValue
End Property

Property Get MajorVer() As Long
    MajorVer = GetValue(FindLine(STR_MAJORVER))
End Property

Property Let MajorVer(ByVal lValue As Long)
    SetValue FindLine(STR_MAJORVER), lValue
End Property

Property Get MinorVer() As Long
    MinorVer = GetValue(FindLine(STR_MINORVER))
End Property

Property Let MinorVer(ByVal lValue As Long)
    SetValue FindLine(STR_MINORVER), lValue
End Property

Property Get RevisionVer() As Long
    RevisionVer = GetValue(FindLine(STR_REVISIONVER))
End Property

Property Let RevisionVer(ByVal lValue As Long)
    SetValue FindLine(STR_REVISIONVER), lValue
End Property

Property Get AutoIncrementVer() As Long
    RevisionVer = GetValue(FindLine(STR_AUTOINCREMENTVER))
End Property

Property Let AutoIncrementVer(ByVal lValue As Long)
    SetValue FindLine(STR_AUTOINCREMENTVER), lValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sVbpFileName As String) As Boolean
    Dim iFile           As Integer
    Dim sLine           As String
    
    '--- store filename
    Set m_oFileName = New cFileName
    m_oFileName.Init sVbpFileName
    '--- open vbp file
    iFile = FreeFile
    Open m_oFileName.FullName() For Input As #iFile
    Set m_cLines = New Collection
    '--- process
    Do While Not EOF(iFile)
        '--- line by line
        Line Input #iFile, sLine
        m_cLines.Add sLine
    Loop
    Close #iFile
    '--- success
    Init = True
End Function

Public Function Save(sVbpFileName As String)
    Dim iFile           As Integer
    Dim lIdx            As Long
    Dim sLine           As String
    
    '--- kill output file (if necessary)
    If FileExists(sVbpFileName) Then
        SetAttr sVbpFileName, vbArchive
        Kill sVbpFileName
    End If
    iFile = FreeFile
    Open sVbpFileName For Output As #iFile
    For lIdx = 1 To m_cLines.Count
        sLine = m_cLines(lIdx)
        Print #iFile, sLine
    Next
    Close #iFile
End Function

Public Function DeleteAllReferences() As Boolean
    Dim lLine       As Long

    '--- references
    lLine = FindLine(STR_REFERENCE)
    Do While lLine > 0
        m_cLines.Remove lLine
        lLine = FindLine(STR_REFERENCE)
    Loop
    '--- objects
    lLine = FindLine(STR_OBJECT)
    Do While lLine > 0
        m_cLines.Remove lLine
        lLine = FindLine(STR_OBJECT)
    Loop
End Function

Public Function AddReference(sFileName As String, Optional ByVal bObject As Boolean)
    Dim oTliInfo        As TypeLibInfo
    Dim sLine           As String
    Dim oFile           As cFileName
    
    On Error Resume Next
    Set oTliInfo = TLI.TypeLibInfoFromFile(sFileName)
    If Err.Number = 0 Then
        If Not bObject Then
            sLine = "Reference=*\G" & oTliInfo.Guid & "#" & Hex(oTliInfo.MajorVersion) & "." & Hex(oTliInfo.MinorVersion) & "#0#" & sFileName & "#" & oTliInfo.Name
        Else
            Set oFile = New cFileName
            oFile.Init sFileName
            sLine = "Object=" & oTliInfo.Guid & "#" & oTliInfo.MajorVersion & "." & oTliInfo.MinorVersion & "#0; " & oFile.FileName & oFile.Suffix
        End If
        '--- always insert as second line
        m_cLines.Add sLine, , , 1
    End If
End Function

Public Function FindLine(sPrefix As String) As Long
    Dim lIdx            As Long
    Dim sLine           As String
    
    For lIdx = 1 To m_cLines.Count
        sLine = m_cLines(lIdx)
        If UCase(Left(sLine, Len(sPrefix))) = UCase(sPrefix) Then
            FindLine = lIdx
            Exit Function
        End If
    Next
End Function

Public Function GetValue(ByVal lLine As Long) As String
    Dim lPos        As Long
    Dim sLine       As String
    
    If lLine > 0 And lLine <= m_cLines.Count Then
        sLine = m_cLines(lLine)
        lPos = InStr(1, sLine, "=")
        If lPos > 0 Then
            GetValue = Mid(sLine, lPos + 1)
        End If
    End If
End Function

Public Function GetStringValue(ByVal lLine As Long) As String
    GetStringValue = GetValue(lLine)
    If Len(GetStringValue) >= 2 Then
        '--- strip double quotes
        GetStringValue = _
                Replace(Left(GetStringValue, 1), Chr(34), "") & _
                Mid(GetStringValue, 2, Len(GetStringValue) - 2) & _
                Replace(Right(GetStringValue, 1), Chr(34), "")
    End If
End Function

Public Function SetValue(ByVal lLine As Long, ByVal sValue As String) As Boolean
    Dim lPos            As Long
    Dim sLine           As String
    
    If lLine > 0 And lLine <= m_cLines.Count Then
        sLine = m_cLines(lLine)
        lPos = InStr(1, sLine, "=")
        If lPos > 0 Then
            '--- if anything is written then change
            If sValue <> "" Then
                sLine = Left(sLine, lPos) & sValue
                m_cLines.Remove lLine
                If m_cLines.Count < lLine Then
                    m_cLines.Add sLine
                Else
                    m_cLines.Add sLine, , lLine
                End If
            Else
                '--- else delete line
                m_cLines.Remove lLine
            End If
            '--- success
            SetValue = True
        End If
    End If
End Function

Public Function SetStringValue(ByVal lLine As Long, ByVal sValue As String) As Boolean
    '--- add double quotes
    If sValue <> "" Then
        sValue = Chr(34) & sValue & Chr(34)
    End If
    SetStringValue = SetValue(lLine, sValue)
End Function

Public Function ExtractVersion(sLine As String) As String
    Dim lPos1       As Long
    Dim lPos2       As Long
    
    lPos1 = InStr(1, sLine, "#")
    lPos2 = InStr(lPos1 + 1, sLine, "#")
    If lPos1 > 0 And lPos2 > 0 Then
        ExtractVersion = Mid(sLine, lPos1 + 1, lPos2 - lPos1 - 1)
    End If
End Function

Public Function ExtractGuid(sLine As String) As String
    Dim lPos1       As Long
    Dim lPos2       As Long
    
    lPos1 = InStr(1, sLine, "{")
    lPos2 = InStr(1, sLine, "}")
    If lPos1 > 0 And lPos2 > 0 Then
        ExtractGuid = Mid(sLine, lPos1, lPos2 - lPos1 + 1)
    End If
End Function

Public Function GetSourceFileName(sLine As String) As String
    Dim lPos1           As Long
    Dim lPos2           As Long
    Dim sFile           As String
    Dim sBinnFile       As String
    
    lPos1 = InStr(1, sLine, "=")
    lPos2 = InStr(lPos1, sLine, ";")
    If lPos1 > 0 Then
        '--- two possible formats of source files in vbp
        If lPos2 > 0 Then
            sFile = Mid(sLine, lPos2 + 2)
        Else
            sFile = Mid(sLine, lPos1 + 1)
        End If
        GetSourceFileName = m_oFileName.FullName("", sFile)
'        '--- these source have a binary companion file
'        If UCase(Left(sLine, Len(STR_USERCONTROL))) = STR_USERCONTROL _
'                Or UCase(Left(sLine, Len(STR_FORM))) = STR_FORM _
'                Or UCase(Left(sLine, Len(STR_DESIGNER))) = STR_DESIGNER _
'                Or UCase(Left(sLine, Len(STR_PROPERTYPAGE))) = STR_PROPERTYPAGE _
'                    Then
'            '--- produce binary filename by replacing last
'            '---   char in source filename with an 'x'
'            sBinnFile = Left(sFile, Len(sFile) - 1) & "x"
'            '--- check for existance
'            If Not FileExists(m_oFileName.FullName("", sBinnFile)) Then
'                sBinnFile = ""
'            End If
'        End If
'        GetSourceFileName = EscapeFile(sFile) & IIf(sBinnFile <> "", " " & EscapeFile(sBinnFile), "")
    End If
End Function

Public Function GetDllFileName(sLine As String) As String
'--- search registry for reference dll/ocx
    Dim sRefPath        As String
    Dim sGuid           As String
    Dim sVersion        As String
    Dim sLastVersion    As String
    
    sGuid = ExtractGuid(sLine)
    sVersion = ExtractVersion(sLine)
    If GetKeyValue(HKEY_CLASSES_ROOT, "TypeLib\" & sGuid & "\" & sVersion & "\0\Win32", "", sRefPath) Then
        GetDllFileName = sRefPath
    Else
        If GetLastKey(HKEY_CLASSES_ROOT, "TypeLib\" & sGuid, sLastVersion) Then
            sVersion = sLastVersion
            If GetKeyValue(HKEY_CLASSES_ROOT, "TypeLib\" & sGuid & "\" & sVersion & "\0\Win32", "", sRefPath) Then
                GetDllFileName = sRefPath
            End If
        End If
    End If
    If GetDllFileName = "" Then
        MsgAlert "Can't find " & sLine
        GetDllFileName = STR_MISSING_REF
    End If
End Function

Public Function GetCompatDllFileName(sLine As String) As String
'--- guess reference compatible dll/ocx
    Dim sFileName       As String
    Dim oFileName       As New cFileName
    
    sFileName = GetDllFileName(sLine)
    '--- search in subfolder "compatible"
    oFileName.Init sFileName
    GetCompatDllFileName = oFileName.FullName(".cmp", , oFileName.Folder & "Compatible\")
    '--- if does not exist try with original suffix
    If Not FileExists(GetCompatDllFileName) Then
        GetCompatDllFileName = oFileName.FullName(, , oFileName.Folder & "Compatible\")
        '--- if does not exist revert to orig reference dll/ocx
        If Not FileExists(GetCompatDllFileName) Then
            GetCompatDllFileName = sFileName
        End If
    End If
End Function

Public Function GetKeyValue( _
            ByVal KeyRoot As UcsHKEYEnum, _
            ByVal KeyName As String, _
            ByVal SubKeyRef As String, _
            ByRef KeyVal As String) As Boolean
'---
'--- ugly, but no time to cleanup the code
'---
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Public Function GetLastKey( _
            KeyRoot As UcsHKEYEnum, _
            KeyName As String, _
            sLastKey As String) As Boolean
    Dim rc              As Long
    Dim hKey            As Long
    Dim sBuffer         As String
    Dim lBufSize        As Long
    Dim sKey            As String
    Dim lIdx            As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    If (rc <> ERROR_SUCCESS) Then GoTo QH
    sBuffer = String(1024, Chr(0))
    lBufSize = Len(sBuffer)
    rc = RegEnumKeyEx(hKey, lIdx, sBuffer, lBufSize, 0, 0, 0, 0)
    Do While rc = ERROR_SUCCESS
        If Mid(sBuffer, lBufSize, 1) = Chr(0) Then
            sKey = Left(sBuffer, lBufSize - 1)
        Else
            sKey = Left(sBuffer, lBufSize)
        End If
        lIdx = lIdx + 1
        rc = RegEnumKeyEx(hKey, lIdx, sBuffer, lBufSize, 0, 0, 0, 0)
    Loop
    rc = RegCloseKey(hKey)
    If sKey <> "" Then
        '--- success
        sLastKey = sKey
        GetLastKey = True
    End If
QH:
    Exit Function
End Function



