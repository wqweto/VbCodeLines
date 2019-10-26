VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' API
'=========================================================================

Private Const MAX_PATH              As Long = 260
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, lpFilePart As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal nBufferLength As Long) As Long

'=========================================================================
' Constants and variables
'=========================================================================

Private m_sDrive                As String
Private m_sFolder               As String
Private m_sFileName             As String
Private m_sSuffix               As String

'=========================================================================
' Properties
'=========================================================================

Property Get Drive() As String
    Drive = m_sDrive
End Property

Property Get Folder() As String
    Folder = m_sFolder
End Property

Property Get FileName() As String
    FileName = m_sFileName
End Property

Property Get Suffix() As String
    Suffix = m_sSuffix
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sFile As String)
    Dim sBuffer         As String
    Dim lPos1           As Long
    Dim lPos2           As Long
    
    '--- canonize filename
    sBuffer = ExpandFileName(sFile)
    '--- find trailing . and \
    lPos1 = InStrRev(sBuffer, ".")
    lPos2 = InStrRev(sBuffer, "\")
    '--- extract filename and suffix
    If lPos1 > lPos2 Then
        m_sSuffix = Mid(sBuffer, lPos1)
        m_sFileName = Mid(sBuffer, lPos2 + 1, lPos1 - lPos2 - 1)
    Else
        m_sSuffix = ""
        m_sFileName = Mid(sBuffer, lPos2 + 1)
    End If
    '--- extract folder
    If lPos2 > 0 Then
        m_sFolder = Left(sBuffer, lPos2)
    Else
        m_sFolder = ""
    End If
    '--- extract drive
    If Mid(sBuffer, 2, 1) = ":" Then
        m_sDrive = Left(sBuffer, 2)
    Else
        m_sDrive = ""
    End If
End Function

Public Function FullName( _
            Optional ByVal NewSuffix, _
            Optional ByVal NewFileName, _
            Optional ByVal NewFolder) As String
    If IsMissing(NewSuffix) Then
        NewSuffix = m_sSuffix
    End If
    '--- prepend . (if not already present)
    If Len(NewSuffix) > 0 Then
        If Left(NewSuffix, 1) <> "." Then
            NewSuffix = "." & NewSuffix
        End If
    End If
    If IsMissing(NewFileName) Then
        NewFileName = m_sFileName
    End If
    If IsMissing(NewFolder) Then
        NewFolder = m_sFolder
    Else
        '--- append trailing \ (if necessary)
        If Right(NewFolder, 1) <> "\" Then
            NewFolder = NewFolder & "\"
        End If
        '--- check if relative path
        If Mid(NewFolder, 2, 1) <> ":" And Left(NewFolder, 2) <> "\\" Then
            '--- check only drive prepend needed
            If Left(NewFolder, 1) = "\" Then
                NewFolder = m_sDrive & NewFolder
            Else
                NewFolder = m_sFolder & NewFolder
            End If
        End If
    End If
    '--- Canonicalize full path
    FullName = ExpandFileName(NewFolder & NewFileName & NewSuffix)
End Function

Public Function ShortName( _
            Optional ByVal NewSuffix, _
            Optional ByVal NewFileName, _
            Optional ByVal NewFolder) As String
    Dim sFile           As String
    Dim lSize           As Long
    Dim sBuffer         As String
    
    sFile = FullName(NewSuffix, NewFileName, NewFolder)
    lSize = GetShortPathName(sFile, vbNullString, 0)
    sBuffer = String(lSize, Chr(0))
    lSize = GetShortPathName(sFile, sBuffer, Len(sBuffer))
    '--- success
    ShortName = Left(sBuffer, lSize)
End Function

Private Function ExpandFileName(sFile As String) As String
    Dim lSize           As Long
    Dim sBuffer         As String
    
    lSize = GetFullPathName(sFile, 0, vbNullString, 0)
    sBuffer = String(lSize, Chr(0))
    lSize = GetFullPathName(sFile, Len(sBuffer), sBuffer, 0)
    '--- success
    ExpandFileName = Left(sBuffer, lSize)
End Function



