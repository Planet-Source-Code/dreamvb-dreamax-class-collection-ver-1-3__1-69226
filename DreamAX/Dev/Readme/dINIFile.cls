VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dINIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private m_IniFilename As String

Public Function INIReadKey(ByVal Selection As String, ByVal KeyName As String, Optional Default = "") As String
'Used to read a value from a given keyname
Dim Ret As Long
Dim sBuff As String
    sBuff = Space(128)
    Ret = GetPrivateProfileString(Selection, KeyName, "{ERR}", sBuff, 128, m_IniFilename)
    
    If (Ret) Then
        'Return the keyvalue
        sBuff = Left(sBuff, Ret)
        If (sBuff = "{ERR}") Then
            INIReadKey = Default
        Else
            INIReadKey = sBuff
        End If
    Else
        INIReadKey = Default
    End If
    
    sBuff = vbNullString
End Function

Public Function INISetKey(ByVal Selection As String, ByVal KeyName As String, ByVal lpValue As String) As Long
    'Used to set a new keynames value
    INISetKey = WritePrivateProfileString(Selection, KeyName, lpValue, m_IniFilename)
End Function

Public Function INIDeleteSelection(ByVal Selection As String) As Long
    'Used to delete an INI Selection, includeing keynames and values
    INIDeleteSelection = WritePrivateProfileString(Selection, 0&, 0&, m_IniFilename)
End Function

Public Function INIRenameSelection(ByVal OldSelection As String, ByVal NewSelection As String) As Boolean
Dim fp As Long
Dim sLine As String
Dim sSelection As String
Dim sLines() As String
Dim Count As Integer

    'Check if the ini is found.
    If (IniFound) Then
        fp = FreeFile
        '
        Open FileName For Input As #fp
            Do Until EOF(fp)
                Line Input #fp, sLine
                If IsSelection(sLine) Then
                    'Lower case selection
                    sSelection = LCase(sLine)
                    'Remove [ and ]
                    sSelection = Left(sSelection, Len(sSelection) - 1)
                    sSelection = Right(sSelection, Len(sSelection) - 1)
                    'Check if the old selection exsits
                    If (sSelection = LCase(OldSelection)) Then
                        'Add new selection
                        sLine = "[" & NewSelection & "]"
                    End If
                End If
                
                ReDim Preserve sLines(Count)
                'Store current line
                sLines(Count) = sLine
                'INC Line counter
                Count = (Count + 1)
            Loop
        Close #fp
        
        'Write new chnages to the file.
        Open FileName For Output As #fp
            Print #fp, Join(sLines, vbCrLf)
        Close #fp
    End If
    
End Function

Public Function INIDeleteKey(ByVal Selection As String, ByVal KeyName As String) As Long
    'Deletes a keyname and it's value from the INI
    INIDeleteKey = WritePrivateProfileString(Selection, KeyName, ByVal 0&, m_IniFilename)
End Function

Function GetSelectionNames() As Collection
'Used to collect selections name of an INI File
On Error Resume Next
Dim col As New Collection
Dim fp As Long
Dim sLine As String
    
    fp = FreeFile
    Open m_IniFilename For Input As #fp
        Do Until EOF(fp)
            Line Input #fp, sLine
            sLine = Trim(sLine)
            If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                col.Add Mid(sLine, 2, Len(sLine) - 2)
            End If
        Loop
    Close #fp
    
    Set GetSelectionNames = col
    Set col = Nothing
    
End Function

Function GetIniSelection(ByVal Selection As String) As Collection
'Collects all the keys in an INI's Selection
Dim Ret As Long
Dim Count As Integer
Dim sBuff As String
Dim vLst() As String
Dim Tmp As String
Dim col As New Collection

    sBuff = Space(8192)
    Ret = GetPrivateProfileString(Selection, 0&, "Err", sBuff, 8192, m_IniFilename)

    
    If (Ret > 0) Then
        Tmp = Left(sBuff, Ret)
        vLst = Split(Tmp, Chr(0))
        'Loop though all the keynames
        For Count = 0 To UBound(vLst)
            If Len(vLst(Count)) <> 0 Then
                'Add the found keyname
                col.Add vLst(Count)
            End If
        Next Count
    End If
    'Return the collection
    Set GetIniSelection = col
    Set col = Nothing
    Erase vLst
    sBuff = ""
    Tmp = ""
    Count = 0
End Function

Function SelectionExists(ByVal Selection As String) As Boolean
'Used to return wereever a selection exists.
Dim col As New Collection
    Set col = GetSelectionNames
    For Each Item In col
        If StrComp(Item, Selection, vbTextCompare) = 0 Then
            SelectionExists = True
            Exit For
        End If
    Next Item
    
    Set col = Nothing
    
End Function

Public Property Get FileName() As String
    FileName = m_IniFilename
End Property

Public Property Let FileName(ByVal NewFilename As String)
    m_IniFilename = NewFilename
End Property

Private Function FindFile(lzFileName As String) As Boolean
    If Trim(Len(lzFileName)) = 0 Then Exit Function
    FindFile = LenB(Dir(lzFileName)) <> 0
End Function

Public Property Get IniFound() As Boolean
    IniFound = FindFile(m_IniFilename)
End Property

Private Function IsSelection(Selection As String) As Boolean
Dim Tmp As String

    'Returns true if we find a selection.
    Tmp = Trim(Selection)
    IsSelection = (Left(Tmp, 1) = "[" And Right(Tmp, 1) = "]")
    
End Function
