VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dValueList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TValues
    Key As String
    Value As String
End Type

Private m_Items() As TValues

Public Sub Add(Key As String, ByVal Value)
Dim cSize As Long

    'Check for a Key
    If Not HasKeyData(Key) Then
        Err.Raise 94, "Add"
        Exit Sub
    'Check if the key been added is in the list.
    ElseIf IndexOfKey(Key) <> (-1) Then
        Err.Raise 457, "Add"
        Exit Sub
    Else
        'Get array size.
        cSize = SafeUBound
        'Resize the array to hold the item.
        ReDim Preserve m_Items(cSize + 1)
        'Set the array item.
        With m_Items(Count)
            .Key = Key
            .Value = Value
        End With
    End If
    'Clear up
    cSize = 0
    
End Sub

Public Sub Delete(ByVal Index)
Dim iCount As Long
Dim cTop As Long
Dim TmpIdx As Long
    
    TmpIdx = GetKeyIndex(Index)
    
    'Checks that the index is in the correct array bounds.
    If Not InRange(TmpIdx) Then
        Err.Raise 9, "Item"
        Exit Sub
    End If
    
    'Get array upper bound.
    cTop = (SafeUBound - 1)
    
    'Loop tho the array and shift all the items up
    For iCount = TmpIdx To cTop
        m_Items(iCount) = m_Items(iCount + 1)
    Next iCount
    
    'if top less then zero clear the array.
    If (cTop < 0) Then
        Call Clear
    Else
        'Resize the array remoevng the last index.
        ReDim Preserve m_Items(cTop)
    End If
    
    'Clear up
    iCount = 0
    TmpIdx = 0
    cTop = 0
End Sub

Public Property Get Count() As Long
    Count = SafeUBound
End Property

Public Sub Clear()
    'Clear all items from the array.
    Erase m_Items
End Sub

Public Property Get Item(ByVal Index) As cItem
Dim oItem As New cItem
Dim TmpIdx As Long

    TmpIdx = GetKeyIndex(Index) 'Get Index
    'Checks that the index is found.
    If Not InRange(TmpIdx) Then
        Err.Raise 9, "Item"
        Exit Property
    End If
    
    'Store the Item data
    oItem.Key = m_Items(TmpIdx).Key
    oItem.Value = m_Items(TmpIdx).Value
    
    'Return item.
    Set Item = oItem
    Set oItem = Nothing
    TmpIdx = 0
End Property

Public Property Let Item(ByVal Index, NewItem As cItem)
Dim TmpIdx As Long
    
    TmpIdx = GetKeyIndex(Index) 'Get Index
    'Checks that the index is in the correct array bounds.
    If Not InRange(TmpIdx) Then
        Err.Raise 9, "Item"
        Exit Property
    End If
    
    'Check if the key is found.
    If Not HasKeyData(NewItem.Key) Then
        Err.Raise 94, "Item"
        Exit Property
    Else
        'Set the new item.
        m_Items(TmpIdx).Key = NewItem.Key
        m_Items(TmpIdx).Value = NewItem.Value
        TmpIdx = 0
    End If
End Property

Public Function ToString() As String
Dim cnt As Long
Dim Tmp() As String

    'Returns the array as a whole string.
    ReDim Tmp(0 To Count) As String
    
    For cnt = 0 To Count
       ' MsgBox m_Items(Cnt).Key
        Tmp(cnt) = m_Items(cnt).Key & "=" _
        & m_Items(cnt).Value
    Next cnt
    
    ToString = Join(Tmp, vbCrLf)
    Erase Tmp
End Function

Public Function KeyEquals(ByVal lpStr As String) As Boolean
    'Returns true of false if an item in the array is equal to lpStr
    KeyEquals = IndexOfKey(lpStr, vbBinaryCompare) <> (-1)
End Function

Private Sub Class_Terminate()
    'Clear things up
    Call Clear
End Sub

'Public and private tools used in this class
'
'
Private Function SafeUBound() As Long
On Error GoTo ErrFlag
    'Return safe size of an error with out the error.
    SafeUBound = UBound(m_Items)
    Exit Function
ErrFlag:
    'Empty array
    SafeUBound = -1
End Function

Private Function InRange(ByVal Index As Long) As Boolean
    'Returns true if we are in a vaild range
    If (Index < 0) Or (Index > UBound(m_Items)) Then
        'We are out of range.
        InRange = False
    Else
        'Rnage ok.
        InRange = True
    End If
End Function

Private Function IndexOfKey(ByVal lpStr As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
Dim pos As Long
Dim X As Long
    'Returns the index position of the item found in the array.
    '-1 is returned if the item is not found.
    pos = -1
    For X = 0 To SafeUBound
        'Compare lpStr with the items in the array
        If StrComp(lpStr, m_Items(X).Key, Compare) = 0 Then
            'Return position found.
            pos = X
            'Skip out of the loop
            Exit For
        End If
    Next X
    'clear var.
    X = 0
    'Return the index
    IndexOfKey = pos
    
End Function

Private Function GetKeyIndex(ByVal Index) As Long
    'Check if Key is Numeric
    If IsNumeric(Index) Then
        'Numeric Key
        GetKeyIndex = Index
    Else
        'Key is Alpha
        GetKeyIndex = IndexOfKey(Index)
    End If
End Function

Function HasKeyData(Key As String) As Boolean
    HasKeyData = Len(Trim(Key)) <> 0
End Function

Public Sub WriteToFile(Filename As String)
On Error GoTo TFileErr:
Dim fp As Long
    
    'Free File Ptr.
    fp = FreeFile
    'Write contents to the file.
    Open Filename For Output As #fp
        Print #fp, ToString
    Close #fp
    
    Exit Sub
    
    'Error Flag
TFileErr:
    If Err Then Err.Raise Err.Number, "SaveWriteToFile", Err.Description, _
    Err.HelpFile, Err.HelpContext
    
End Sub

Public Sub LoadFromFile(Filename As String)
On Error GoTo TFileErr:
Dim fp As Long
Dim sLine As String
Dim s_pos As Integer
Dim sKey As String
    
    Call Clear
    
    'Free File Ptr.
    fp = FreeFile
    'Write contents to the file.
    Open Filename For Input As #fp
        Do Until EOF(fp)
            Line Input #fp, sLine
            sLine = Trim(sLine)
            If Len(sLine) > 0 Then
                'Get Position of assignment =
                s_pos = InStr(1, sLine, "=", vbBinaryCompare)
                'Check for assignment Sign.
                If (s_pos > 0) Then
                    'Extract the Key.
                    sKey = Trim(Left(sLine, s_pos - 1))
                    'Check that the key has a value, and it not aready in the List
                    If HasKeyData(sKey) And Not (KeyEquals(sKey)) Then
                        'Add the key and the value.
                        Add sKey, Mid(sLine, s_pos + 1)
                    End If
                End If
            End If
            DoEvents
        Loop
    Close #fp
    
    'Clear up
    sKey = vbNullString
    sLine = vbNullString
    s_pos = 0
    Exit Sub
    
    'Error Flag
TFileErr:
    If Err Then Err.Raise Err.Number, "LoadFromFile", Err.Description, _
    Err.HelpFile, Err.HelpContext
    
End Sub

