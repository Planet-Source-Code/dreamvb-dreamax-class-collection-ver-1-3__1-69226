VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dFileOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const ERROR_SUCCESS As Long = 0&
Private Const REG_EXPAND_SZ = 2

Private Enum EFileParts
    fDrive = 0
    fDir = 1
    fPath = 2
    fFilename = 3
    fExt = 4
End Enum

Public Function ChangeFileExt(ByVal Filename As String, ByVal Extension As String) As String
    'Change FileExt
    ChangeFileExt = RemoveFileExt(Filename) & Extension
End Function

Function RemoveFileExt(ByVal Filename As String) As String
Dim sPos As Integer

    'Remove FileExt from Filename.
    sPos = InStrRev(Filename, ".", Len(Filename), vbBinaryCompare)
    
    If (sPos = 0) Then
        RemoveFileExt = Filename
    Else
        RemoveFileExt = Left$(Filename, sPos - 1)
    End If
End Function

Function ExtractFileDrive(ByVal Filename As String) As String
    'Extract Drive eg C:\windows\readme.txt, returns C:\
    ExtractFileDrive = GetfParts(Filename, fDrive)
End Function

Function ExtractFileDir(ByVal Filename As String) As String
    ExtractFileDir = GetfParts(Filename, fDir)
End Function

Function ExtractFilePath(ByVal Filename As String) As String
    ExtractFilePath = GetfParts(Filename, fPath)
End Function

Function ExtractFileName(ByVal Filename As String) As String
    ExtractFileName = GetfParts(Filename, fFilename)
End Function

Function ExtractFileTitle(ByVal Filename As String) As String
    If FileExtExisits(Filename) Then
       ExtractFileTitle = RemoveFileExt(GetfParts(Filename, fFilename))
    Else
        ExtractFileTitle = Filename
    End If
End Function

Function ExtractFileExt(ByVal Filename As String) As String
    ExtractFileExt = GetfParts(Filename, fExt)
End Function

Function FileExtExisits(ByVal Filename As String) As Boolean
    FileExtExisits = Not (GetfParts(Filename, fExt)) = vbNullString
End Function

Function FileExists(ByVal Filename As String) As Boolean
On Error GoTo nFound:
    FileExists = (GetAttr(Filename) And vbDirectory) = 0
    Exit Function
nFound:
    FileExists = False
End Function

Function FileAge(ByVal Filename As String) As Date
    FileAge = FileDateTime(Filename)
End Function

Sub ForceDirectories(ByVal FilePath As String)

End Sub

Function DirectoryExists(ByVal PathName As String) As Boolean
On Error GoTo nFound:
    DirectoryExists = GetAttr(PathName) = vbDirectory
    Exit Function
nFound:
    DirectoryExists = False
End Function

Function GetAppDir() As String
    GetAppDir = AddBackSlash(App.Path)
End Function

Function AddBackSlash(ByVal FilePath As String) As String
    If Right$(FilePath, 1) = "\" Then
        AddBackSlash = FilePath
    Else
        AddBackSlash = FilePath & "\"
    End If
End Function

Function RenameFile(ByVal OldFilename As String, ByVal NewFilename As String) As Boolean
    If Not FileExists(OldFilename) Then
        Exit Function
    Else
        Name OldFilename As NewFilename
    End If
    
    RenameFile = True
End Function

Function GetShortName(ByVal Filename As String) As String
Dim Ret As Long
Dim sBuff As String

    'Get the length needed to store the name
    Ret = GetShortPathName(Filename, vbNullString, ByVal 0&)
    
    If (Ret = 0) Then
        GetShortName = Filename
    Else
        'Create space for the filename
        sBuff = Space$(Ret)
        'Get Filename into sBuff
        GetShortPathName Filename, sBuff, Ret
        'Return new name striping nullchar away
        GetShortName = Left(sBuff, Ret - 1)
        sBuff = vbNullString
    End If
    
End Function

Function GetFileType(ByVal Filename As String) As String
Dim FileExt As String
Dim StrA As String

    FileExt = GetfParts(Filename, fExt)
    
    If (FileExt <> vbNullString) Then
        'Get the keyname
        StrA = ReadStringValue(HKEY_CLASSES_ROOT, FileExt, vbNullString, 2)
        
        If (StrA <> vbNullString) Then
            'GetFile Type
            StrA = ReadStringValue(HKEY_CLASSES_ROOT, StrA, vbNullString, 2)
        End If
        'Return Filetype
        GetFileType = StrA
    End If
    
    StrA = vbNullString
    FileExt = vbNullString
End Function

Private Function GetfParts(ByVal Filename As String, fPart As EFileParts) As String
Dim Pos1 As Integer
Dim Pos2 As Integer
Dim Pos3 As Integer
Dim RetStr As String
    
    If Len(Filename) = 0 Then
        Exit Function
    End If
    
    Pos1 = InStr(1, Filename, ":\", vbBinaryCompare)
    Pos2 = InStrRev(Filename, "\", Len(Filename), vbBinaryCompare)
    Pos3 = InStrRev(Filename, ".", Len(Filename), vbBinaryCompare)
    '
    RetStr = vbNullString
    
    Select Case fPart
        Case fDrive
            'Get Drive
            If (Pos1 = 0) Then
                Exit Function
            Else
                RetStr = Left$(Filename, Pos1)
            End If
        Case fDir, fPath
            'Get FileDir and FilePath
            If (Pos2 = 0) Then
                Exit Function
            ElseIf (fPart = fDir) Then
                'Get FileDirName
                RetStr = Left$(Filename, Pos2 - 1)
            Else
                RetStr = Left$(Filename, Pos2)
            End If
        Case fFilename
            'Get Filename
            If (Pos2 = 0) Or (Pos3 = 0) Then
                RetStr = Filename
            Else
                RetStr = Mid$(Filename, Pos2 + 1, Pos3)
            End If
        Case fExt
            'Get FileExt
            If (Pos3 = 0) Then
                Exit Function
            Else
                RetStr = Mid$(Filename, Pos3)
            End If
    End Select
    
    GetfParts = RetStr
End Function

Private Function ReadStringValue(hKey As Long, ByVal KeyPath As String, KeyName As String, RegType As Integer) As String
Dim sBuff As String
Dim lBuffSize As Long
Dim sRegKey As Long
    
    'Check if Key is found.
    If RegOpenKeyEx(hKey, KeyPath, 0&, KEY_QUERY_VALUE, sRegKey) <> ERROR_SUCCESS Then
        Exit Function
    End If
    'Check for Value Name is found.
    If RegQueryValueEx(sRegKey, KeyName, 0, REG_EXPAND_SZ, ByVal 0&, lBuffSize) <> ERROR_SUCCESS Then
        RegCloseKey sRegKey
        Exit Function
    End If
    
    'Create buffer for Value.
    sBuff = Space$(lBuffSize - 1)
    RegQueryValueEx sRegKey, KeyName, 0&, REG_EXPAND_SZ, ByVal sBuff, lBuffSize
    'Close the Open Key
    RegCloseKey sRegKey
    '
    ReadStringValue = sBuff
    sBuff = vbNullString
End Function

