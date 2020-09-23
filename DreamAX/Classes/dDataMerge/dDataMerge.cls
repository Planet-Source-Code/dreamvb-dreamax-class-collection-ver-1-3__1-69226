VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dDataMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Error()

Public Sub MergeTo(DestinationFile As String, InsertFile As String)
On Error GoTo WriteErr:
Dim fp As Long
Dim fOffset As Long
Dim fData() As Byte

    'Check if DestinationFile and InsertFile are found.
    If Not IsFileHere(DestinationFile) Then
        Err.Raise 53, DestinationFile
        Exit Sub
    ElseIf Not IsFileHere(InsertFile) Then
        Err.Raise 53, InsertFile
    Else
        'Read in the InsertFile Data
        fp = FreeFile
        Open InsertFile For Binary As #fp
            If LOF(fp) > 0 Then
                ReDim Preserve fData(0 To LOF(fp) - 1)
            End If
            Get #fp, , fData
        Close #fp
        'Put InsertData into DestinationFile
        Open DestinationFile For Binary As #fp
            'Get Length of the File
            fOffset = LOF(fp)
            'Move to the end of the file
            Seek #fp, fOffset
            'Put in header SIG
            Put #fp, , "dMerge"
            'Put the Insert Data
            Put #fp, , fData
            'put the offset
            Put #fp, , fOffset
        Close #fp
    End If
    'Clear up
    Erase fData
    fOffset = 0
    
    Exit Sub
    
'Error Trap
WriteErr:
    RaiseEvent Error
End Sub

Public Sub ExtractTo(MergeSrcFile As String, ExtractFilename As String)
On Error GoTo ReadErr:
Dim fStart As Long
Dim Sig As String
Dim fp As Long
Dim Bytes() As Byte

    'Check if the MergeSrcFile is found.
    If Not IsFileHere(MergeSrcFile) Then
        Err.Raise 53, MergeSrcFile
        Exit Sub
    Else
        fp = FreeFile
        Sig = Space(6)
        'Open the merge file.
        Open MergeSrcFile For Binary As #fp
            'Get File Data Start offset.
            Get #fp, LOF(fp) - 3, fStart
            'Check for vaild offset.
            If (fStart = 0) Then
                'Not Found.
                Err.Raise 54, MergeSrcFile, "DataOffSet Not Found"
            Else
                'Check for SIG
                Get #fp, fStart, Sig
                If LCase(Sig) <> "dmerge" Then
                    Err.Raise 54, MergeSrcFile, "DataOffSet Not Found"
                Else
                    'Resize and extract the data
                    ReDim Bytes((LOF(fp) - fStart))
                    Get #fp, , Bytes
                    'Extract the New Data.
                    Open ExtractFilename For Binary As #2
                        Put #2, , Bytes
                    Close #2
                End If
            End If
        Close #fp
    End If
    
    'Clear up
    Erase Bytes
    Sig = vbNullString
    fStart = 0
    
    Exit Sub
    
'Error Trap
ReadErr:
    RaiseEvent Error
    Close #fp
End Sub

Private Function IsFileHere(ByVal lzFilename As String) As Boolean
    
    If Len(lzFilename) = 0 Then
        IsFileHere = False
        Exit Function
    End If
    
    If LenB(Dir(lzFilename)) = 0 Then
        IsFileHere = False
        Exit Function
    Else
        IsFileHere = True
    End If
    
End Function
